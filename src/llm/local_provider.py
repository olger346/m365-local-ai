import json
import re
from typing import Any, Dict, List, Optional
from openai import OpenAI
from pydantic import BaseModel, Field

from src.config import load_prompts, settings
from src.connectors.graph_client import EmailMessage


class TriageDecision(BaseModel):
    categoria: str = Field(default="consulta_frecuente")
    requiere_respuesta: bool = Field(default=True)
    confianza: float = Field(default=0.5)
    accion_sugerida: str = Field(default="create_draft")
    razonamiento: str = Field(default="")
    borrador_respuesta_markdown: str = Field(default="")


def extract_json_object(raw_text: str) -> Dict[str, Any]:
    """Extrae un objeto JSON válido de la respuesta del LLM, incluso con bloques <think> o markdown."""
    # Eliminar bloques <think>...</think> de modelos razonadores (DeepSeek-R1, QwQ, etc.)
    cleaned = re.sub(r"<think>.*?</think>", "", raw_text, flags=re.DOTALL).strip()

    # Buscar bloques ```json ... ```
    fence_match = re.search(r"```(?:json)?\s*(\{.*?\})\s*```", cleaned, flags=re.DOTALL)
    if fence_match:
        cleaned = fence_match.group(1)
    else:
        # Buscar desde la primera llave '{' hasta la última '}'
        start = cleaned.find("{")
        end = cleaned.rfind("}")
        if start != -1 and end != -1 and end > start:
            cleaned = cleaned[start : end + 1]

    return json.loads(cleaned)


class LocalLLMClient:
    """Conector compatible con LM Studio (puerto 1234) y Ollama (puerto 11434)."""

    def __init__(
        self,
        base_url: Optional[str] = None,
        model: Optional[str] = None,
        api_key: Optional[str] = None,
    ) -> None:
        self.base_url = base_url or settings.llm_base_url
        self.model = model or settings.llm_model
        self.api_key = api_key or settings.llm_api_key
        self.client = OpenAI(base_url=self.base_url, api_key=self.api_key)

    def list_available_models(self) -> List[str]:
        """Lista los modelos activos en LM Studio u Ollama."""
        models = self.client.models.list()
        return [m.id for m in models.data]

    def resolve_model_name(self) -> str:
        """Si LLM_MODEL es 'local-model', detecta automáticamente el primer modelo cargado."""
        if self.model and self.model != "local-model":
            return self.model
        available = self.list_available_models()
        if not available:
            raise RuntimeError(
                f"No hay modelos cargados en {self.base_url}. "
                "Asegúrate de cargar un modelo en LM Studio u Ollama."
            )
        return available[0]

    def triage_and_draft_email(
        self,
        email: EmailMessage,
        custom_instructions: Optional[str] = None,
    ) -> TriageDecision:
        """Analiza un correo con el LLM local y genera clasificación + borrador."""
        prompts_cfg = load_prompts()
        profile = prompts_cfg.get("user_profile", {})
        raw_system = prompts_cfg.get("system_prompt_triage", "")

        system_prompt = raw_system.format(
            name=profile.get("name", "Usuario"),
            role=profile.get("role", "Profesional"),
            tone=profile.get("tone", "Profesional y claro"),
            language=profile.get("language", "Español"),
            signature=profile.get("signature", "Saludos,"),
        )

        body_text = email.body_markdown or email.body_preview
        user_content = (
            f"Remitente: {email.sender_name} <{email.sender_email}>\n"
            f"Para: {', '.join(email.to_recipients)}\n"
            f"CC: {', '.join(email.cc_recipients)}\n"
            f"Fecha: {email.received_datetime}\n"
            f"Asunto: {email.subject}\n\n"
            f"--- CONTENIDO DEL CORREO ---\n{body_text[:6000]}\n"
        )
        if custom_instructions:
            user_content += f"\n--- INSTRUCCIONES ADICIONALES PARA ESTA RESPUESTA ---\n{custom_instructions}\n"

        active_model = self.resolve_model_name()
        response = self.client.chat.completions.create(
            model=active_model,
            temperature=settings.llm_temperature,
            messages=[
                {"role": "system", "content": system_prompt},
                {"role": "user", "content": user_content},
            ],
        )

        raw_reply = response.choices[0].message.content or "{}"
        parsed = extract_json_object(raw_reply)
        return TriageDecision(**parsed)
