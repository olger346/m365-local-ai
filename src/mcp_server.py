from datetime import datetime, timedelta
import json
from pathlib import Path
import sys
from typing import Any, Dict, List, Optional
import httpx

# Garantizar que el directorio raíz del proyecto esté en sys.path aunque LM Studio no aplique "cwd"
PROJECT_ROOT_DIR = str(Path(__file__).resolve().parent.parent)
if PROJECT_ROOT_DIR not in sys.path:
    sys.path.insert(0, PROJECT_ROOT_DIR)

try:
    from mcp.server.mcpserver import MCPServer as FastMCP
except ImportError:
    from mcp.server.fastmcp import FastMCP

from src.config import PROJECT_ROOT, load_rules, settings
from src.connectors import get_mail_client
from src.connectors.graph_client import clear_saved_session
from src.engine.db import AuditDB
from src.engine.processor import decide_final_action
from src.llm.local_provider import LocalLLMClient

mcp = FastMCP("m365-local-ai")

SHORT_ID_CACHE_FILE = PROJECT_ROOT / ".m365_short_ids.json"


def _save_short_ids(messages: List[Dict[str, Any]]) -> List[Dict[str, Any]]:
    """
    Asigna alias cortos (E1, E2, E3...) a los correos listados para que modelos locales
    (como Gemma o Qwen) no fallen al copiar IDs largos de 150 caracteres de Microsoft Graph.
    """
    mapping: Dict[str, str] = {}
    if SHORT_ID_CACHE_FILE.exists():
        try:
            mapping = json.loads(SHORT_ID_CACHE_FILE.read_text(encoding="utf-8"))
        except Exception:
            mapping = {}

    enriched: List[Dict[str, Any]] = []
    for idx, msg in enumerate(messages, 1):
        short_key = f"E{idx}"
        full_id = msg.get("id", "")
        if full_id:
            mapping[short_key] = full_id
            mapping[str(idx)] = full_id
            if msg.get("sender_email"):
                mapping[msg["sender_email"].lower()] = full_id
        item = {"short_id": short_key, **msg}
        enriched.append(item)

    try:
        SHORT_ID_CACHE_FILE.write_text(json.dumps(mapping, indent=2), encoding="utf-8")
    except Exception:
        pass
    return enriched


def _resolve_message_id(client: Any, raw_id: str) -> str:
    """
    Permite que el modelo pase 'E1', '1', el correo del remitente o el ID completo de Graph.
    """
    candidate = (raw_id or "").strip()
    if SHORT_ID_CACHE_FILE.exists():
        try:
            mapping = json.loads(SHORT_ID_CACHE_FILE.read_text(encoding="utf-8"))
            if candidate.upper() in mapping:
                return mapping[candidate.upper()]
            if candidate.lower() in mapping:
                return mapping[candidate.lower()]
        except Exception:
            pass

    # Si el modelo pasó una dirección de correo en vez de un ID, buscar el último correo de ese remitente
    if "@" in candidate and len(candidate) < 80:
        found = client.search_emails(query=candidate, limit=1)
        if found:
            return found[0].id

    return candidate


def _parse_email_list(raw_emails: str) -> List[str]:
    """Convierte una cadena separada por comas o punto y coma en lista de correos limpios."""
    normalized = raw_emails.replace(";", ",")
    return [addr.strip() for addr in normalized.split(",") if addr.strip() and "@" in addr]


def _format_exception(exc: Exception) -> Dict[str, Any]:
    """Formatea cualquier excepción (incluyendo errores HTTP de Graph) sin romper la ejecución del MCP."""
    if isinstance(exc, httpx.HTTPStatusError):
        try:
            err_json = exc.response.json()
            graph_msg = err_json.get("error", {}).get("message") or exc.response.text
        except Exception:
            graph_msg = exc.response.text
        return {
            "status": "error",
            "http_status": exc.response.status_code,
            "error_message": f"Microsoft Graph API error ({exc.response.status_code}): {graph_msg}",
        }
    return {
        "status": "error",
        "error_message": f"{type(exc).__name__}: {str(exc)}",
    }


@mcp.tool()
def get_account_profile() -> Dict[str, Any]:
    """
    Verifica la sesión activa de Microsoft 365 (o abre el navegador para iniciar sesión)
    y devuelve el nombre y correo de la cuenta conectada.
    """
    try:
        client = get_mail_client()
        profile = client.get_my_profile()
        return {
            "status": "ok",
            "displayName": profile.get("displayName"),
            "email": profile.get("mail") or profile.get("userPrincipalName"),
            "id": profile.get("id"),
        }
    except Exception as exc:
        return _format_exception(exc)


@mcp.tool()
def logout_and_delete_token() -> Dict[str, Any]:
    """
    Cierra la sesión de Microsoft 365 sobrescribiendo con ceros y eliminando el archivo
    de token local (.m365_token_cache.bin) para que nadie pueda copiarlo.
    """
    try:
        removed = clear_saved_session()
        return {
            "status": "logged_out",
            "token_deleted": removed,
            "message": "El token local (.m365_token_cache.bin) ha sido sobrescrito y eliminado de forma segura.",
        }
    except Exception as exc:
        return _format_exception(exc)


@mcp.tool()
def list_emails(unread_only: bool = False, limit: int = 5, folder: str = "inbox") -> List[Dict[str, Any]]:
    """
    Lista los correos más recientes de Microsoft 365 (Outlook).
    Devuelve un 'short_id' (ej. 'E1', 'E2') además del 'id' completo para usar fácilmente en otras herramientas.
    - unread_only: Si es True, solo devuelve correos no leídos. Por defecto False.
    - limit: Cantidad máxima de correos a devolver (por defecto 5).
    - folder: Carpeta de Outlook ('inbox', 'drafts', 'sentitems').
    """
    try:
        client = get_mail_client()
        messages = client.list_emails(unread_only=unread_only, limit=limit, folder=folder, include_body=False)
        raw_list = [m.model_dump() for m in messages]
        return _save_short_ids(raw_list)
    except Exception as exc:
        return [_format_exception(exc)]


@mcp.tool()
def read_email(message_id: str) -> Dict[str, Any]:
    """
    Lee el contenido completo (en Markdown limpio) de un correo.
    Puedes pasar el 'short_id' (ej. 'E1', 'E2'), el número ('1', '2') o el 'id' completo.
    """
    try:
        client = get_mail_client()
        real_id = _resolve_message_id(client, message_id)
        msg = client.get_email(real_id)
        return {"status": "ok", **msg.model_dump()}
    except Exception as exc:
        return _format_exception(exc)


@mcp.tool()
def search_emails(query: str, limit: int = 5) -> List[Dict[str, Any]]:
    """
    Busca correos en Microsoft 365 por palabras clave, asunto o remitente.
    """
    try:
        client = get_mail_client()
        messages = client.search_emails(query=query, limit=limit)
        raw_list = [m.model_dump() for m in messages]
        return _save_short_ids(raw_list)
    except Exception as exc:
        return [_format_exception(exc)]


@mcp.tool()
def send_new_email(
    to_email: str,
    subject: str,
    body_markdown: str,
    cc_email: str = "",
) -> Dict[str, Any]:
    """
    Envía un correo electrónico NUEVO a uno o más destinatarios.
    - to_email: Dirección de correo del destinatario (o varias separadas por coma, ej. 'juan@empresa.com').
    - subject: Asunto del correo.
    - body_markdown: Contenido del mensaje en texto o Markdown.
    - cc_email: (Opcional) Direcciones en copia (CC) separadas por coma.
    """
    try:
        to_list = _parse_email_list(to_email)
        if not to_list:
            return {
                "status": "error",
                "error_message": f"Destinatario inválido: '{to_email}'. Debe ser un correo válido como usuario@dominio.com.",
            }
        cc_list = _parse_email_list(cc_email) if cc_email else None
        client = get_mail_client()
        client.send_new_email(
            to_recipients=to_list,
            subject=subject,
            body_markdown=body_markdown,
            cc_recipients=cc_list,
        )
        return {
            "status": "sent",
            "to": to_list,
            "cc": cc_list or [],
            "subject": subject,
        }
    except Exception as exc:
        return _format_exception(exc)


@mcp.tool()
def create_new_email_draft(
    to_email: str,
    subject: str,
    body_markdown: str,
    cc_email: str = "",
) -> Dict[str, Any]:
    """
    Crea un BORRADOR de un correo nuevo en Outlook sin enviarlo.
    - to_email: Dirección de correo del destinatario (o varias separadas por coma).
    - subject: Asunto del correo.
    - body_markdown: Contenido del borrador en texto o Markdown.
    - cc_email: (Opcional) Direcciones en copia (CC) separadas por coma.
    """
    try:
        to_list = _parse_email_list(to_email)
        if not to_list:
            return {
                "status": "error",
                "error_message": f"Destinatario inválido: '{to_email}'. Debe ser un correo válido como usuario@dominio.com.",
            }
        cc_list = _parse_email_list(cc_email) if cc_email else None
        client = get_mail_client()
        draft = client.create_new_draft(
            to_recipients=to_list,
            subject=subject,
            body_markdown=body_markdown,
            cc_recipients=cc_list,
        )
        return {
            "status": "draft_created",
            "draft_id": draft.get("id"),
            "to": to_list,
            "subject": draft.get("subject"),
            "web_link": draft.get("webLink"),
        }
    except Exception as exc:
        return _format_exception(exc)


@mcp.tool()
def create_reply_draft(message_id: str, reply_markdown: str, reply_all: bool = False) -> Dict[str, Any]:
    """
    Crea un borrador de RESPUESTA a un correo existente en Outlook sin enviarlo.
    Puedes usar el 'short_id' (ej. 'E1', 'E2') o el 'id' completo del correo.
    """
    try:
        client = get_mail_client()
        real_id = _resolve_message_id(client, message_id)
        draft = client.create_reply_draft(message_id=real_id, reply_markdown=reply_markdown, reply_all=reply_all)
        return {
            "status": "reply_draft_created",
            "draft_id": draft.get("id"),
            "subject": draft.get("subject"),
            "web_link": draft.get("webLink"),
        }
    except Exception as exc:
        return _format_exception(exc)


@mcp.tool()
def send_reply_email(message_id: str, reply_markdown: str, reply_all: bool = False) -> Dict[str, Any]:
    """
    Envía inmediatamente una RESPUESTA a un correo recibido existente.
    Puedes usar el 'short_id' (ej. 'E1', 'E2') o el 'id' completo del correo.
    NOTA: Si quieres enviar un correo nuevo a una dirección de email, usa 'send_new_email' en su lugar.
    """
    try:
        client = get_mail_client()
        real_id = _resolve_message_id(client, message_id)
        client.send_reply(message_id=real_id, reply_markdown=reply_markdown, reply_all=reply_all)
        return {
            "status": "reply_sent",
            "message_id": real_id,
        }
    except Exception as exc:
        return _format_exception(exc)


@mcp.tool()
def analyze_and_reply_with_local_ai(
    message_id: str,
    custom_instructions: Optional[str] = None,
    allow_auto_send: bool = True,
) -> Dict[str, Any]:
    """
    Envía un correo recibido (puedes usar 'E1', 'E2' o su 'id') a la IA local para clasificarlo,
    redactar la respuesta y aplicar las reglas de seguridad (rules.yaml).
    """
    try:
        mail_client = get_mail_client()
        llm_client = LocalLLMClient()
        rules = load_rules()
        db = AuditDB()

        real_id = _resolve_message_id(mail_client, message_id)
        email = mail_client.get_email(real_id)
        decision = llm_client.triage_and_draft_email(email, custom_instructions=custom_instructions)
        final_action, reason = decide_final_action(email, decision, rules)

        if not allow_auto_send and final_action == "auto_reply":
            final_action = "create_draft"
            reason = "Forzado a borrador por parámetro allow_auto_send=False."

        result_meta: Dict[str, Any] = {}
        if final_action == "auto_reply" and not settings.dry_run:
            mail_client.send_reply(message_id=email.id, reply_markdown=decision.borrador_respuesta_markdown)
            mail_client.mark_as_read(email.id)
            result_meta["executed"] = "sent_automatically"
        elif final_action == "create_draft" and not settings.dry_run:
            draft = mail_client.create_reply_draft(
                message_id=email.id,
                reply_markdown=decision.borrador_respuesta_markdown,
            )
            result_meta["executed"] = "draft_created"
            result_meta["draft_id"] = draft.get("id")
        else:
            result_meta["executed"] = "ignored_or_dry_run"

        db.record_action(
            message_id=email.id,
            sender_email=email.sender_email,
            subject=email.subject,
            category=decision.categoria,
            confidence=decision.confianza,
            final_action=final_action,
            reason=reason,
            draft_markdown=decision.borrador_respuesta_markdown,
        )

        return {
            "status": "ok",
            "message_id": email.id,
            "subject": email.subject,
            "sender": email.sender_email,
            "llm_decision": decision.model_dump(),
            "final_action": final_action,
            "reason": reason,
            **result_meta,
        }
    except Exception as exc:
        return _format_exception(exc)


@mcp.tool()
def get_day_schedule(date_yyyy_mm_dd: str = "") -> Dict[str, Any]:
    """
    Consulta la agenda del calendario y los bloques libres (mañana y tarde) para planificar el día.
    - date_yyyy_mm_dd: Fecha en formato 'YYYY-MM-DD' (ej. '2026-10-22'). Si se deja vacío, consulta el día de HOY.
    Devuelve la lista de reuniones del día, los huecos libres ('free_slots') entre 08:00 y 18:00
    y el indicador 'is_afternoon_mostly_free'.
    """
    try:
        client = get_mail_client()
        schedule = client.get_day_schedule_and_free_slots(
            target_date_str=date_yyyy_mm_dd if date_yyyy_mm_dd.strip() else None
        )
        return {"status": "ok", **schedule}
    except Exception as exc:
        return _format_exception(exc)


@mcp.tool()
def get_upcoming_calendar_events(days_ahead: int = 7) -> Dict[str, Any]:
    """
    Lista todas las reuniones y eventos del calendario para los próximos N días (por defecto 7 días).
    Ideal para revisar la semana o buscar compromisos próximos.
    """
    try:
        local_tz = datetime.now().astimezone().tzinfo
        now_local = datetime.now(local_tz)
        end_local = now_local + timedelta(days=max(1, min(days_ahead, 30)))

        client = get_mail_client()
        events = client.list_calendar_events(start_dt=now_local, end_dt=end_local, limit=50)
        return {
            "status": "ok",
            "from": now_local.strftime("%Y-%m-%d %H:%M"),
            "to": end_local.strftime("%Y-%m-%d %H:%M"),
            "total_events": len(events),
            "events": [e.model_dump() for e in events],
        }
    except Exception as exc:
        return _format_exception(exc)


@mcp.tool()
def create_calendar_event(
    subject: str,
    start_datetime: str,
    end_datetime: str,
    body_markdown: str = "",
    attendees_email: str = "",
    location: str = "",
) -> Dict[str, Any]:
    """
    Crea una reunión o bloque de trabajo en el calendario de Microsoft 365 (Outlook).
    - subject: Título del evento o bloque (ej. 'Revisar propuesta', 'Bloque de enfoque').
    - start_datetime: Fecha y hora de inicio en hora local 'YYYY-MM-DD HH:MM' (ej. '2026-10-22 15:00').
    - end_datetime: Fecha y hora de fin en hora local 'YYYY-MM-DD HH:MM' (ej. '2026-10-22 16:00').
    - body_markdown: (Opcional) Notas o agenda del evento en Markdown.
    - attendees_email: (Opcional) Correos de invitados separados por coma.
    - location: (Opcional) Sala o ubicación.
    """
    try:
        attendees_list = _parse_email_list(attendees_email) if attendees_email else None
        client = get_mail_client()
        created = client.create_calendar_event(
            subject=subject,
            start_local_str=start_datetime,
            end_local_str=end_datetime,
            body_markdown=body_markdown,
            attendees_emails=attendees_list,
            location=location,
        )
        return {
            "status": "event_created",
            "event_id": created.get("id"),
            "subject": created.get("subject"),
            "start": start_datetime,
            "end": end_datetime,
            "attendees": attendees_list or [],
            "web_link": created.get("webLink"),
        }
    except Exception as exc:
        return _format_exception(exc)


@mcp.tool()
def get_audit_history(limit: int = 20) -> List[Dict[str, Any]]:
    """
    Consulta el historial local (SQLite) de correos procesados por la IA local.
    """
    try:
        db = AuditDB()
        return db.get_recent_logs(limit=limit)
    except Exception as exc:
        return [_format_exception(exc)]


def main() -> None:
    mcp.run()


if __name__ == "__main__":
    main()
