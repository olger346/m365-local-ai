from pathlib import Path
from typing import Any, Dict, List
import yaml
from pydantic import BaseModel, Field
from pydantic_settings import BaseSettings, SettingsConfigDict

PROJECT_ROOT = Path(__file__).resolve().parent.parent


class Settings(BaseSettings):
    """Configuración cargada desde variables de entorno o archivo .env."""

    model_config = SettingsConfigDict(
        env_file=str(PROJECT_ROOT / ".env"),
        env_file_encoding="utf-8",
        extra="ignore",
    )

    # Modo de conexión: "graph" (OAuth2 en la nube) o "outlook_local" (App de escritorio vía Windows COM)
    m365_backend: str = Field(default="outlook_local", alias="M365_BACKEND")
    m365_account_email: str = Field(default="", alias="M365_ACCOUNT_EMAIL")

    # Microsoft 365 Graph API
    m365_client_id: str = Field(default="", alias="M365_CLIENT_ID")
    m365_tenant_id: str = Field(default="common", alias="M365_TENANT_ID")
    m365_auth_flow: str = Field(default="interactive", alias="M365_AUTH_FLOW")
    m365_scopes: str = Field(
        default="User.Read,Mail.Read,Mail.ReadWrite,Mail.Send",
        alias="M365_SCOPES",
    )

    # IA Local (LM Studio / Ollama)
    llm_base_url: str = Field(default="http://localhost:1234/v1", alias="LLM_BASE_URL")
    llm_model: str = Field(default="local-model", alias="LLM_MODEL")
    llm_api_key: str = Field(default="lm-studio", alias="LLM_API_KEY")
    llm_temperature: float = Field(default=0.2, alias="LLM_TEMPERATURE")

    # Daemon
    poll_interval_seconds: int = Field(default=120, alias="POLL_INTERVAL_SECONDS")
    max_emails_per_batch: int = Field(default=10, alias="MAX_EMAILS_PER_BATCH")
    dry_run: bool = Field(default=False, alias="DRY_RUN")


class AutonomyConfig(BaseModel):
    enable_auto_send: bool = True
    min_confidence_for_auto_send: float = 0.85
    max_recipients_for_auto_send: int = 3


class WhitelistConfig(BaseModel):
    allowed_senders: List[str] = Field(default_factory=list)
    allowed_domains: List[str] = Field(default_factory=list)
    allowed_categories: List[str] = Field(default_factory=list)


class IgnoreRulesConfig(BaseModel):
    blocked_sender_substrings: List[str] = Field(default_factory=list)
    blocked_subject_keywords: List[str] = Field(default_factory=list)


class RulesConfig(BaseModel):
    autonomy: AutonomyConfig = Field(default_factory=AutonomyConfig)
    whitelist: WhitelistConfig = Field(default_factory=WhitelistConfig)
    ignore_rules: IgnoreRulesConfig = Field(default_factory=IgnoreRulesConfig)


def load_rules() -> RulesConfig:
    rules_path = PROJECT_ROOT / "config" / "rules.yaml"
    if not rules_path.exists():
        rules_path = PROJECT_ROOT / "config" / "rules.example.yaml"
    if not rules_path.exists():
        return RulesConfig()
    with open(rules_path, "r", encoding="utf-8") as f:
        data = yaml.safe_load(f) or {}
    return RulesConfig(**data)


def load_prompts() -> Dict[str, Any]:
    prompts_path = PROJECT_ROOT / "config" / "prompts.yaml"
    if not prompts_path.exists():
        prompts_path = PROJECT_ROOT / "config" / "prompts.example.yaml"
    if not prompts_path.exists():
        return {}
    with open(prompts_path, "r", encoding="utf-8") as f:
        return yaml.safe_load(f) or {}


settings = Settings()
