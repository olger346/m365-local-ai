"""
Script de inicio de sesión y diagnóstico para:
1. Verificar conexión con LM Studio (puerto 1234) u Ollama (puerto 11434)
2. Autenticar con Microsoft 365 (Graph API) y mantener la sesión activa
3. Eliminar el token (.m365_token_cache.bin) de forma segura al presionar Ctrl+C
"""
import signal
import sys
import time
from rich.console import Console
from rich.panel import Panel

from src.config import settings
from src.connectors import get_mail_client
from src.connectors.graph_client import clear_saved_session, get_configured_scopes
from src.llm.local_provider import LocalLLMClient

if hasattr(sys.stdout, "reconfigure"):
    sys.stdout.reconfigure(encoding="utf-8", errors="replace")

console = Console()


def test_local_llm() -> None:
    console.rule("[bold cyan]1. Probando conexión con IA Local (LM Studio / Ollama)[/bold cyan]")
    try:
        llm = LocalLLMClient()
        models = llm.list_available_models()
        if models:
            console.print(f"[green]✔ Conectado a {llm.base_url}[/green]")
            console.print(f"Modelos disponibles: [bold]{', '.join(models)}[/bold]")
        else:
            console.print(
                f"[yellow]⚠ El servidor en {llm.base_url} respondió, pero no tiene ningún modelo cargado.[/yellow]"
            )
    except Exception as exc:
        console.print(
            f"[red]✘ No se pudo conectar a {settings.llm_base_url}:[/red] {exc}\n"
            "[dim]Asegúrate de tener LM Studio (Local Server activo en puerto 1234) u Ollama (puerto 11434) corriendo.[/dim]"
        )


def test_m365_auth() -> bool:
    backend = (settings.m365_backend or "graph").strip().lower()
    console.rule(f"[bold cyan]2. Probando conexión con Microsoft 365 (Modo: {backend})[/bold cyan]")
    active_scopes = get_configured_scopes()
    if backend == "graph":
        console.print(f"Permisos solicitados (M365_SCOPES): [bold yellow]{', '.join(active_scopes)}[/bold yellow]\n")
    try:
        client = get_mail_client()
        profile = client.get_my_profile()
        display_name = profile.get("displayName", "Usuario")
        mail = profile.get("mail") or profile.get("userPrincipalName", "")
        console.print(f"[green]✔ Conectado exitosamente como:[/green] [bold]{display_name}[/bold] ({mail})")
        if profile.get("available_accounts"):
            console.print(f"Cuentas detectadas en Outlook local: [dim]{', '.join(profile['available_accounts'])}[/dim]")

        if backend != "graph" or any(s.startswith("Mail.") for s in active_scopes):
            emails = client.list_emails(unread_only=False, limit=3)
            console.print(f"\nÚltimos {len(emails)} correos en tu bandeja de entrada:")
            for idx, msg in enumerate(emails, 1):
                status = "📩 No leído" if not msg.is_read else "✉️ Leído"
                console.print(f"  {idx}. [{status}] [bold]{msg.subject}[/bold] — de {msg.sender_email}")

        if backend != "graph" or any(s.startswith("Calendars.") for s in active_scopes):
            schedule = client.get_day_schedule_and_free_slots()
            console.print(
                f"\n📅 Agenda de hoy ({schedule['date']}): [bold]{schedule['total_events']} evento(s)[/bold] | "
                f"¿Tarde mayormente libre?: [bold]{'Sí' if schedule['is_afternoon_mostly_free'] else 'No'}[/bold]"
            )
        return True
    except Exception as exc:
        console.print(Panel(str(exc), title="[red]Pendiente o Error en Microsoft 365[/red]", border_style="red"))
        return False


def _handle_exit_signal(signum: int, frame: object) -> None:
    raise KeyboardInterrupt


if __name__ == "__main__":
    signal.signal(signal.SIGINT, _handle_exit_signal)
    if hasattr(signal, "SIGBREAK"):
        signal.signal(signal.SIGBREAK, _handle_exit_signal)

    test_local_llm()
    print()
    ok = test_m365_auth()
    if ok:
        console.print(
            Panel(
                "[bold green]Sesión activa y lista para usar en LM Studio / Antigravity.[/bold green]\n"
                "Deja esta ventana abierta mientras trabajas.\n"
                "Cuando termines, presiona [bold yellow]Ctrl + C[/bold yellow] para destruir y eliminar el token local ([code].m365_token_cache.bin[/code]).",
                title="🔒 Control de Sesión Segura",
                border_style="green",
            )
        )
        try:
            while True:
                time.sleep(1)
        except KeyboardInterrupt:
            console.print("\n[yellow]Cerrando sesión y destruyendo token local...[/yellow]")
            if clear_saved_session():
                console.print("[bold green]✔ Token (.m365_token_cache.bin) sobrescrito y eliminado de forma segura.[/bold green]")
            else:
                console.print("[dim]No había archivo de token pendiente por eliminar.[/dim]")
