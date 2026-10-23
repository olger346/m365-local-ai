"""Selector automático de conector: Microsoft Graph API (nube) u Outlook Local (Windows COM)."""
from typing import Union
from src.config import settings
from src.connectors.graph_client import GraphMailClient
from src.connectors.outlook_com import OutlookLocalClient


def get_mail_client() -> Union[GraphMailClient, OutlookLocalClient]:
    """Devuelve el cliente de correo/calendario según M365_BACKEND en .env."""
    backend = (settings.m365_backend or "graph").strip().lower()
    if backend in ("outlook_local", "outlook", "com", "local"):
        return OutlookLocalClient()
    return GraphMailClient()
