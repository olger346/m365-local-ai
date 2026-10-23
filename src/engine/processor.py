from typing import Tuple
from src.config import RulesConfig, load_rules
from src.connectors.graph_client import EmailMessage
from src.llm.local_provider import TriageDecision


def check_pre_ignore_rules(email: EmailMessage, rules: RulesConfig) -> Tuple[bool, str]:
    """Verifica si un correo debe ignorarse antes siquiera de llamar al LLM (ej. no-reply)."""
    sender_lower = email.sender_email.lower()
    subject_lower = email.subject.lower()

    for blocked_sub in rules.ignore_rules.blocked_sender_substrings:
        if blocked_sub.lower() in sender_lower:
            return True, f"Remitente bloqueado por subcadena '{blocked_sub}'"

    for kw in rules.ignore_rules.blocked_subject_keywords:
        if kw.lower() in subject_lower:
            return True, f"Asunto bloqueado por palabra clave '{kw}'"

    return False, ""


def decide_final_action(
    email: EmailMessage,
    decision: TriageDecision,
    rules: RulesConfig,
) -> Tuple[str, str]:
    """
    Aplica los guardrails de seguridad sobre la sugerencia del LLM.
    Retorna: (final_action, reason) donde final_action es 'auto_reply' | 'create_draft' | 'ignore'.
    """
    if not decision.requiere_respuesta or decision.accion_sugerida == "ignore":
        return "ignore", f"No requiere respuesta ({decision.razonamiento})"

    if not decision.borrador_respuesta_markdown.strip():
        return "ignore", "El modelo no generó texto de respuesta."

    # Si contiene marcadores pendientes de rellenar por el usuario, forzar borrador
    if "[COMPLETAR" in decision.borrador_respuesta_markdown.upper():
        return "create_draft", "El borrador tiene campos pendientes ([COMPLETAR]) para revisión humana."

    # Si el envío automático está deshabilitado globalmente
    if not rules.autonomy.enable_auto_send:
        return "create_draft", "Auto-envío global desactivado en rules.yaml."

    # Si el LLM sugirió solo crear borrador
    if decision.accion_sugerida != "auto_reply":
        return "create_draft", f"IA sugirió borrador ({decision.razonamiento})"

    # Verificar número de destinatarios
    total_recipients = len(email.to_recipients) + len(email.cc_recipients)
    if total_recipients > rules.autonomy.max_recipients_for_auto_send:
        return (
            "create_draft",
            f"Demasiados destinatarios ({total_recipients} > {rules.autonomy.max_recipients_for_auto_send}); degradado a borrador.",
        )

    # Verificar umbral de confianza
    if decision.confianza < rules.autonomy.min_confidence_for_auto_send:
        return (
            "create_draft",
            f"Confianza insuficiente ({decision.confianza:.2f} < {rules.autonomy.min_confidence_for_auto_send}); degradado a borrador.",
        )

    # Verificar lista blanca de remitente o dominio
    sender_lower = email.sender_email.lower()
    sender_domain = sender_lower.split("@")[-1] if "@" in sender_lower else ""

    allowed_senders = [s.lower() for s in rules.whitelist.allowed_senders]
    allowed_domains = [d.lower() for d in rules.whitelist.allowed_domains]
    allowed_cats = [c.lower() for c in rules.whitelist.allowed_categories]

    sender_whitelisted = (sender_lower in allowed_senders) or (sender_domain in allowed_domains)
    category_whitelisted = (
        not allowed_cats or decision.categoria.lower() in allowed_cats
    )

    if sender_whitelisted and category_whitelisted:
        return (
            "auto_reply",
            f"Remitente/dominio autorizado y confianza alta ({decision.confianza:.2f}).",
        )

    return (
        "create_draft",
        f"Remitente '{sender_lower}' no está en whitelist de auto-envío; guardado en Borradores.",
    )
