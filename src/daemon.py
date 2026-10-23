import argparse
import sys
import time
from typing import Any
from rich.console import Console
from rich.table import Table

from src.config import load_rules, settings
from src.connectors import get_mail_client
from src.connectors.graph_client import clear_saved_session
from src.engine.db import AuditDB
from src.engine.processor import check_pre_ignore_rules, decide_final_action
from src.llm.local_provider import LocalLLMClient

if hasattr(sys.stdout, "reconfigure"):
    sys.stdout.reconfigure(encoding="utf-8", errors="replace")

console = Console()


def process_batch_once(mail_client: Any, llm_client: LocalLLMClient, db: AuditDB) -> int:
    """Procesa un lote de correos no leídos aplicando IA local + reglas de seguridad."""
    rules = load_rules()
    unread_list = mail_client.list_emails(
        unread_only=True,
        limit=settings.max_emails_per_batch,
        folder="inbox",
        include_body=True,
    )

    processed_count = 0
    for email in unread_list:
        if db.is_processed(email.id):
            continue

        should_ignore, pre_reason = check_pre_ignore_rules(email, rules)
        if should_ignore:
            console.print(f"[dim]Ignorando correo de {email.sender_email}: {pre_reason}[/dim]")
            db.record_action(
                message_id=email.id,
                sender_email=email.sender_email,
                subject=email.subject,
                category="pre_filtered",
                confidence=1.0,
                final_action="ignore",
                reason=pre_reason,
                draft_markdown="",
            )
            processed_count += 1
            continue

        console.print(f"[cyan]Analizando con IA local:[/cyan] '{email.subject}' de [bold]{email.sender_email}[/bold]")
        decision = llm_client.triage_and_draft_email(email)
        final_action, reason = decide_final_action(email, decision, rules)

        if final_action == "auto_reply":
            if settings.dry_run:
                console.print(f"  [yellow][DRY-RUN] Se habría enviado auto-respuesta:[/yellow] {reason}")
            else:
                mail_client.send_reply(email.id, decision.borrador_respuesta_markdown)
                mail_client.mark_as_read(email.id)
                console.print(f"  [green]✔ Respuesta enviada automáticamente:[/green] {reason}")
        elif final_action == "create_draft":
            if settings.dry_run:
                console.print(f"  [yellow][DRY-RUN] Se habría creado borrador:[/yellow] {reason}")
            else:
                mail_client.create_reply_draft(email.id, decision.borrador_respuesta_markdown)
                console.print(f"  [blue]📝 Borrador creado en Outlook:[/blue] {reason}")
        else:
            console.print(f"  [dim]⏭ Ignorado:[/dim] {reason}")

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
        processed_count += 1

    return processed_count


def main() -> None:
    parser = argparse.ArgumentParser(description="Demonio de auto-respuesta y borradores M365 + IA Local")
    parser.add_argument("--once", action="store_true", help="Ejecuta un solo ciclo y termina.")
    args = parser.parse_args()

    mail_client = get_mail_client()
    llm_client = LocalLLMClient()
    db = AuditDB()

    console.rule("[bold blue]M365 Local AI Worker[/bold blue]")
    console.print(f"Servidor LLM Local: [bold]{llm_client.base_url}[/bold]")
    console.print(f"Modo Dry-Run: [bold]{settings.dry_run}[/bold]")

    if args.once:
        count = process_batch_once(mail_client, llm_client, db)
        console.print(f"[green]Ciclo finalizado. Correos nuevos procesados: {count}[/green]")
        return

    console.print(f"Monitoreando bandeja de entrada cada {settings.poll_interval_seconds} segundos... (Ctrl+C para salir)")
    try:
        while True:
            count = process_batch_once(mail_client, llm_client, db)
            if count > 0:
                console.print(f"[green]Se procesaron {count} correos en este ciclo.[/green]")
            time.sleep(settings.poll_interval_seconds)
    except KeyboardInterrupt:
        console.print("\n[yellow]Servicio detenido por el usuario (Ctrl+C). Eliminando token local...[/yellow]")
        if clear_saved_session():
            console.print("[bold green]✔ Token (.m365_token_cache.bin) eliminado de forma segura.[/bold green]")


if __name__ == "__main__":
    main()
