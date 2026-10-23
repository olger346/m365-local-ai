from datetime import datetime, timedelta, timezone
import os
import sys
from pathlib import Path
from typing import Any, Dict, List, Optional
import httpx
import msal
import markdown
from bs4 import BeautifulSoup
from markdownify import markdownify as md
from pydantic import BaseModel

from src.config import PROJECT_ROOT, settings

GRAPH_BASE_URL = "https://graph.microsoft.com/v1.0"
TOKEN_CACHE_PATH = PROJECT_ROOT / ".m365_token_cache.bin"
SHORT_IDS_PATH = PROJECT_ROOT / ".m365_short_ids.json"


def get_configured_scopes() -> List[str]:
    """Devuelve la lista de permisos (scopes) configurados en .env."""
    raw = settings.m365_scopes or "User.Read,Mail.Read,Mail.ReadWrite,Mail.Send"
    return [s.strip() for s in raw.split(",") if s.strip()]


def clear_saved_session() -> bool:
    """Sobrescribe y elimina de forma segura el token OAuth2 y la caché de IDs en disco."""
    removed = False
    for path in (TOKEN_CACHE_PATH, SHORT_IDS_PATH):
        if path.exists():
            try:
                size = path.stat().st_size
                path.write_bytes(b"\x00" * size)
                path.unlink()
                removed = True
            except Exception:
                try:
                    path.unlink(missing_ok=True)
                    removed = True
                except Exception:
                    pass
    return removed


class EmailMessage(BaseModel):
    id: str
    conversation_id: Optional[str] = None
    subject: str
    sender_name: str
    sender_email: str
    to_recipients: List[str] = []
    cc_recipients: List[str] = []
    received_datetime: str
    is_read: bool
    body_preview: str
    body_markdown: str = ""
    web_link: Optional[str] = None


class CalendarEvent(BaseModel):
    id: str
    subject: str
    start_local: str
    end_local: str
    is_all_day: bool = False
    show_as: str = "busy"
    location: str = ""
    organizer_name: str = ""
    organizer_email: str = ""
    attendees: List[str] = []
    is_online_meeting: bool = False
    body_preview: str = ""
    web_link: Optional[str] = None


def _parse_graph_utc_to_local(dt_str: str) -> datetime:
    """Convierte una cadena ISO de Microsoft Graph (en UTC) a la zona horaria local de la PC."""
    local_tz = datetime.now().astimezone().tzinfo
    clean = dt_str.strip().rstrip("Z")
    if "." in clean:
        base, frac = clean.split(".", 1)
        clean = f"{base}.{frac[:6]}"
    dt_utc = datetime.fromisoformat(clean).replace(tzinfo=timezone.utc)
    return dt_utc.astimezone(local_tz)


def html_to_clean_markdown(html_content: str) -> str:
    """Convierte el HTML de Outlook a Markdown limpio para ahorrar tokens en el LLM local."""
    if not html_content:
        return ""
    soup = BeautifulSoup(html_content, "html.parser")
    for tag in soup(["style", "script", "head", "meta"]):
        tag.decompose()
    cleaned_html = str(soup)
    text_md = md(cleaned_html, heading_style="ATX", strip=["img"])
    lines = [line.rstrip() for line in text_md.splitlines()]
    # Reducir múltiples líneas en blanco consecutivas
    compact_lines: List[str] = []
    blank_count = 0
    for line in lines:
        if not line.strip():
            blank_count += 1
            if blank_count <= 2:
                compact_lines.append("")
        else:
            blank_count = 0
            compact_lines.append(line)
    return "\n".join(compact_lines).strip()


def markdown_to_html(md_text: str) -> str:
    """Convierte la respuesta Markdown del LLM a HTML compatible con Outlook."""
    body_html = markdown.markdown(md_text, extensions=["nl2br", "tables", "fenced_code"])
    return (
        "<div style='font-family: Calibri, Arial, sans-serif; font-size: 11pt;'>"
        f"{body_html}"
        "</div>"
    )


class GraphMailClient:
    """Cliente OAuth2 para Microsoft Graph API con caché persistente de sesión."""

    def __init__(
        self,
        client_id: Optional[str] = None,
        tenant_id: Optional[str] = None,
        auth_flow: Optional[str] = None,
    ) -> None:
        self.client_id = client_id or settings.m365_client_id
        self.tenant_id = tenant_id or settings.m365_tenant_id
        self.auth_flow = (auth_flow or settings.m365_auth_flow).lower()

        if not self.client_id or self.client_id == "tu-client-id-aqui":
            raise ValueError(
                "Falta configurar M365_CLIENT_ID en el archivo .env. "
                "Revisa el README.md o ejecuta python test_auth.py para ver los pasos."
            )

        self.authority = f"https://login.microsoftonline.com/{self.tenant_id}"
        self.cache = msal.SerializableTokenCache()
        if TOKEN_CACHE_PATH.exists():
            self.cache.deserialize(TOKEN_CACHE_PATH.read_text(encoding="utf-8"))

        self.app = msal.PublicClientApplication(
            client_id=self.client_id,
            authority=self.authority,
            token_cache=self.cache,
        )

    def _save_cache(self) -> None:
        if self.cache.has_state_changed:
            TOKEN_CACHE_PATH.write_text(self.cache.serialize(), encoding="utf-8")

    def get_access_token(self) -> str:
        """Obtiene un token válido desde la caché o inicia el flujo interactivo."""
        scopes = get_configured_scopes()
        accounts = self.app.get_accounts()
        result: Optional[Dict[str, Any]] = None

        if accounts:
            result = self.app.acquire_token_silent(scopes, account=accounts[0])

        if not result:
            if self.auth_flow == "device_code":
                flow = self.app.initiate_device_flow(scopes=scopes)
                if "user_code" not in flow:
                    raise RuntimeError(f"Error al iniciar Device Code flow: {flow}")
                print(flow["message"], file=sys.stderr, flush=True)
                result = self.app.acquire_token_by_device_flow(flow)
            else:
                result = self.app.acquire_token_interactive(
                    scopes=scopes,
                    prompt="select_account",
                )

        self._save_cache()

        if not result or "access_token" not in result:
            error_desc = (result or {}).get("error_description", "Error desconocido")
            raise RuntimeError(f"No se pudo obtener token de Microsoft 365: {error_desc}")

        return str(result["access_token"])

    def _headers(self) -> Dict[str, str]:
        return {
            "Authorization": f"Bearer {self.get_access_token()}",
            "Content-Type": "application/json",
        }

    def _parse_message(self, raw: Dict[str, Any], include_body: bool = False) -> EmailMessage:
        from_data = raw.get("from", {}).get("emailAddress", {})
        to_list = [
            r.get("emailAddress", {}).get("address", "")
            for r in raw.get("toRecipients", [])
            if r.get("emailAddress", {}).get("address")
        ]
        cc_list = [
            r.get("emailAddress", {}).get("address", "")
            for r in raw.get("ccRecipients", [])
            if r.get("emailAddress", {}).get("address")
        ]

        body_md = ""
        if include_body and "body" in raw:
            content = raw["body"].get("content", "")
            content_type = raw["body"].get("contentType", "text").lower()
            body_md = html_to_clean_markdown(content) if content_type == "html" else content

        return EmailMessage(
            id=raw["id"],
            conversation_id=raw.get("conversationId"),
            subject=raw.get("subject") or "(Sin asunto)",
            sender_name=from_data.get("name") or "",
            sender_email=from_data.get("address") or "",
            to_recipients=to_list,
            cc_recipients=cc_list,
            received_datetime=raw.get("receivedDateTime") or "",
            is_read=bool(raw.get("isRead", False)),
            body_preview=raw.get("bodyPreview") or "",
            body_markdown=body_md,
            web_link=raw.get("webLink"),
        )

    def get_my_profile(self) -> Dict[str, Any]:
        """Obtiene el perfil del usuario autenticado."""
        with httpx.Client(timeout=20.0) as client:
            resp = client.get(f"{GRAPH_BASE_URL}/me", headers=self._headers())
            resp.raise_for_status()
            return resp.json()

    def list_emails(
        self,
        unread_only: bool = True,
        limit: int = 10,
        folder: str = "inbox",
        include_body: bool = False,
    ) -> List[EmailMessage]:
        """Lista los correos más recientes de una carpeta (por defecto Inbox)."""
        select_fields = (
            "id,conversationId,subject,from,toRecipients,ccRecipients,"
            "receivedDateTime,isRead,bodyPreview,webLink"
        )
        if include_body:
            select_fields += ",body"

        params: Dict[str, Any] = {
            "$top": limit,
            "$select": select_fields,
            "$orderby": "receivedDateTime desc",
        }
        if unread_only:
            params["$filter"] = "isRead eq false"

        url = f"{GRAPH_BASE_URL}/me/mailFolders/{folder}/messages"
        with httpx.Client(timeout=25.0) as client:
            resp = client.get(url, headers=self._headers(), params=params)
            resp.raise_for_status()
            items = resp.json().get("value", [])

        return [self._parse_message(item, include_body=include_body) for item in items]

    def get_email(self, message_id: str) -> EmailMessage:
        """Obtiene el detalle completo de un correo específico en Markdown limpio."""
        url = f"{GRAPH_BASE_URL}/me/messages/{message_id}"
        with httpx.Client(timeout=25.0) as client:
            resp = client.get(url, headers=self._headers())
            resp.raise_for_status()
            return self._parse_message(resp.json(), include_body=True)

    def search_emails(self, query: str, limit: int = 10) -> List[EmailMessage]:
        """Busca correos usando Microsoft Search ($search)."""
        url = f"{GRAPH_BASE_URL}/me/messages"
        params = {
            "$search": f'"{query}"',
            "$top": limit,
            "$select": "id,conversationId,subject,from,toRecipients,ccRecipients,receivedDateTime,isRead,bodyPreview,webLink",
        }
        with httpx.Client(timeout=25.0) as client:
            resp = client.get(url, headers=self._headers(), params=params)
            resp.raise_for_status()
            items = resp.json().get("value", [])
        return [self._parse_message(item, include_body=False) for item in items]

    def create_reply_draft(self, message_id: str, reply_markdown: str, reply_all: bool = False) -> Dict[str, Any]:
        """Crea un borrador de respuesta en Outlook conservando el hilo original."""
        action = "createReplyAll" if reply_all else "createReply"
        url = f"{GRAPH_BASE_URL}/me/messages/{message_id}/{action}"
        payload = {
            "comment": markdown_to_html(reply_markdown),
        }
        with httpx.Client(timeout=25.0) as client:
            resp = client.post(url, headers=self._headers(), json=payload)
            resp.raise_for_status()
            return resp.json()

    def send_reply(self, message_id: str, reply_markdown: str, reply_all: bool = False) -> None:
        """Envía inmediatamente una respuesta al correo especificado."""
        action = "replyAll" if reply_all else "reply"
        url = f"{GRAPH_BASE_URL}/me/messages/{message_id}/{action}"
        payload = {
            "comment": markdown_to_html(reply_markdown),
        }
        with httpx.Client(timeout=25.0) as client:
            resp = client.post(url, headers=self._headers(), json=payload)
            resp.raise_for_status()

    @staticmethod
    def _format_recipients(addresses: List[str]) -> List[Dict[str, Any]]:
        cleaned = [addr.strip() for addr in addresses if addr and addr.strip()]
        return [{"emailAddress": {"address": addr}} for addr in cleaned]

    def send_new_email(
        self,
        to_recipients: List[str],
        subject: str,
        body_markdown: str,
        cc_recipients: Optional[List[str]] = None,
        save_to_sent_items: bool = True,
    ) -> None:
        """Envía un correo electrónico nuevo desde cero a uno o varios destinatarios."""
        url = f"{GRAPH_BASE_URL}/me/sendMail"
        message_obj: Dict[str, Any] = {
            "subject": subject,
            "body": {
                "contentType": "HTML",
                "content": markdown_to_html(body_markdown),
            },
            "toRecipients": self._format_recipients(to_recipients),
        }
        if cc_recipients:
            message_obj["ccRecipients"] = self._format_recipients(cc_recipients)

        payload = {
            "message": message_obj,
            "saveToSentItems": save_to_sent_items,
        }
        with httpx.Client(timeout=25.0) as client:
            resp = client.post(url, headers=self._headers(), json=payload)
            resp.raise_for_status()

    def create_new_draft(
        self,
        to_recipients: List[str],
        subject: str,
        body_markdown: str,
        cc_recipients: Optional[List[str]] = None,
    ) -> Dict[str, Any]:
        """Crea un borrador de un correo nuevo en la carpeta Borradores de Outlook."""
        url = f"{GRAPH_BASE_URL}/me/messages"
        payload: Dict[str, Any] = {
            "subject": subject,
            "body": {
                "contentType": "HTML",
                "content": markdown_to_html(body_markdown),
            },
            "toRecipients": self._format_recipients(to_recipients),
        }
        if cc_recipients:
            payload["ccRecipients"] = self._format_recipients(cc_recipients)

        with httpx.Client(timeout=25.0) as client:
            resp = client.post(url, headers=self._headers(), json=payload)
            resp.raise_for_status()
            return resp.json()

    def mark_as_read(self, message_id: str) -> None:
        """Marca un correo como leído."""
        url = f"{GRAPH_BASE_URL}/me/messages/{message_id}"
        with httpx.Client(timeout=20.0) as client:
            resp = client.patch(url, headers=self._headers(), json={"isRead": True})
            resp.raise_for_status()

    # =========================================================================
    # FUNCIONES DE CALENDARIO Y PLANIFICACIÓN DE AGENDA
    # =========================================================================

    def _parse_calendar_event(self, raw: Dict[str, Any]) -> CalendarEvent:
        start_raw = raw.get("start", {}).get("dateTime", "")
        end_raw = raw.get("end", {}).get("dateTime", "")

        start_local_str = start_raw
        end_local_str = end_raw
        try:
            if start_raw:
                start_local_str = _parse_graph_utc_to_local(start_raw).strftime("%Y-%m-%d %H:%M")
            if end_raw:
                end_local_str = _parse_graph_utc_to_local(end_raw).strftime("%Y-%m-%d %H:%M")
        except Exception:
            pass

        org_data = raw.get("organizer", {}).get("emailAddress", {})
        attendee_list = [
            a.get("emailAddress", {}).get("address", "")
            for a in raw.get("attendees", [])
            if a.get("emailAddress", {}).get("address")
        ]
        loc_name = raw.get("location", {}).get("displayName") or ""

        return CalendarEvent(
            id=raw.get("id", ""),
            subject=raw.get("subject") or "(Sin título)",
            start_local=start_local_str,
            end_local=end_local_str,
            is_all_day=bool(raw.get("isAllDay", False)),
            show_as=raw.get("showAs") or "busy",
            location=loc_name,
            organizer_name=org_data.get("name") or "",
            organizer_email=org_data.get("address") or "",
            attendees=attendee_list,
            is_online_meeting=bool(raw.get("isOnlineMeeting", False)),
            body_preview=raw.get("bodyPreview") or "",
            web_link=raw.get("webLink"),
        )

    def list_calendar_events(
        self,
        start_dt: datetime,
        end_dt: datetime,
        limit: int = 50,
    ) -> List[CalendarEvent]:
        """
        Obtiene los eventos del calendario entre start_dt y end_dt usando /me/calendarView
        (expande automáticamente reuniones recurrentes).
        """
        start_utc = start_dt.astimezone(timezone.utc).strftime("%Y-%m-%dT%H:%M:%SZ")
        end_utc = end_dt.astimezone(timezone.utc).strftime("%Y-%m-%dT%H:%M:%SZ")

        url = f"{GRAPH_BASE_URL}/me/calendarView"
        params: Dict[str, Any] = {
            "startDateTime": start_utc,
            "endDateTime": end_utc,
            "$top": limit,
            "$orderby": "start/dateTime asc",
            "$select": (
                "id,subject,start,end,isAllDay,showAs,location,"
                "organizer,attendees,isOnlineMeeting,bodyPreview,webLink,isCancelled"
            ),
        }
        with httpx.Client(timeout=25.0) as client:
            resp = client.get(url, headers=self._headers(), params=params)
            resp.raise_for_status()
            items = resp.json().get("value", [])

        active_items = [item for item in items if not item.get("isCancelled", False)]
        return [self._parse_calendar_event(item) for item in active_items]

    def get_day_schedule_and_free_slots(
        self,
        target_date_str: Optional[str] = None,
        work_start_hour: int = 8,
        work_end_hour: int = 18,
    ) -> Dict[str, Any]:
        """
        Analiza la agenda de un día específico (por defecto hoy) y calcula los bloques libres
        en la mañana y en la tarde para facilitar la planificación del día.
        """
        local_tz = datetime.now().astimezone().tzinfo
        now_local = datetime.now(local_tz)

        if target_date_str:
            base_date = datetime.strptime(target_date_str.strip()[:10], "%Y-%m-%d").date()
        else:
            base_date = now_local.date()

        day_start = datetime(base_date.year, base_date.month, base_date.day, 0, 0, 0, tzinfo=local_tz)
        day_end = datetime(base_date.year, base_date.month, base_date.day, 23, 59, 59, tzinfo=local_tz)

        events = self.list_calendar_events(day_start, day_end)

        work_start = datetime(base_date.year, base_date.month, base_date.day, work_start_hour, 0, tzinfo=local_tz)
        work_end = datetime(base_date.year, base_date.month, base_date.day, work_end_hour, 0, tzinfo=local_tz)

        busy_intervals: List[tuple[datetime, datetime]] = []
        for ev in events:
            if ev.show_as == "free":
                continue
            if ev.is_all_day:
                busy_intervals.append((work_start, work_end))
                continue
            try:
                ev_s = datetime.strptime(ev.start_local, "%Y-%m-%d %H:%M").replace(tzinfo=local_tz)
                ev_e = datetime.strptime(ev.end_local, "%Y-%m-%d %H:%M").replace(tzinfo=local_tz)
                s_clamped = max(work_start, ev_s)
                e_clamped = min(work_end, ev_e)
                if e_clamped > s_clamped:
                    busy_intervals.append((s_clamped, e_clamped))
            except Exception:
                pass

        busy_intervals.sort(key=lambda x: x[0])
        merged_busy: List[tuple[datetime, datetime]] = []
        for s, e in busy_intervals:
            if not merged_busy or s > merged_busy[-1][1]:
                merged_busy.append((s, e))
            else:
                merged_busy[-1] = (merged_busy[-1][0], max(merged_busy[-1][1], e))

        free_slots: List[Dict[str, Any]] = []
        cursor = work_start
        for b_start, b_end in merged_busy:
            if b_start > cursor:
                mins = int((b_start - cursor).total_seconds() // 60)
                if mins >= 15:
                    free_slots.append(
                        {
                            "start": cursor.strftime("%H:%M"),
                            "end": b_start.strftime("%H:%M"),
                            "duration_minutes": mins,
                            "period": "mañana" if cursor.hour < 12 else "tarde",
                        }
                    )
            cursor = max(cursor, b_end)

        if cursor < work_end:
            mins = int((work_end - cursor).total_seconds() // 60)
            if mins >= 15:
                free_slots.append(
                    {
                        "start": cursor.strftime("%H:%M"),
                        "end": work_end.strftime("%H:%M"),
                        "duration_minutes": mins,
                        "period": "mañana" if cursor.hour < 12 else "tarde",
                    }
                )

        afternoon_free_mins = sum(
            slot["duration_minutes"]
            for slot in free_slots
            if slot["period"] == "tarde" or int(slot["end"].split(":")[0]) > 12
        )

        return {
            "date": base_date.strftime("%Y-%m-%d"),
            "current_local_time": now_local.strftime("%Y-%m-%d %H:%M"),
            "work_hours_evaluated": f"{work_start_hour:02d}:00 - {work_end_hour:02d}:00",
            "total_events": len(events),
            "events": [e.model_dump() for e in events],
            "free_slots": free_slots,
            "is_afternoon_mostly_free": afternoon_free_mins >= 180,
        }

    def create_calendar_event(
        self,
        subject: str,
        start_local_str: str,
        end_local_str: str,
        body_markdown: str = "",
        attendees_emails: Optional[List[str]] = None,
        location: str = "",
    ) -> Dict[str, Any]:
        """
        Crea un evento o bloque de tiempo en el calendario de Microsoft 365.
        Acepta fechas en formato 'YYYY-MM-DD HH:MM' o 'YYYY-MM-DDTHH:MM' en hora local.
        """
        local_tz = datetime.now().astimezone().tzinfo

        def _to_utc_iso(raw_dt: str) -> str:
            normalized = raw_dt.strip().replace("T", " ")
            dt_local = datetime.strptime(normalized[:16], "%Y-%m-%d %H:%M").replace(tzinfo=local_tz)
            return dt_local.astimezone(timezone.utc).strftime("%Y-%m-%dT%H:%M:%S")

        payload: Dict[str, Any] = {
            "subject": subject,
            "start": {
                "dateTime": _to_utc_iso(start_local_str),
                "timeZone": "UTC",
            },
            "end": {
                "dateTime": _to_utc_iso(end_local_str),
                "timeZone": "UTC",
            },
        }
        if body_markdown:
            payload["body"] = {
                "contentType": "HTML",
                "content": markdown_to_html(body_markdown),
            }
        if location:
            payload["location"] = {"displayName": location}
        if attendees_emails:
            payload["attendees"] = [
                {
                    "emailAddress": {"address": addr.strip()},
                    "type": "required",
                }
                for addr in attendees_emails
                if addr and "@" in addr
            ]

        url = f"{GRAPH_BASE_URL}/me/events"
        with httpx.Client(timeout=25.0) as client:
            resp = client.post(url, headers=self._headers(), json=payload)
            resp.raise_for_status()
            return resp.json()
