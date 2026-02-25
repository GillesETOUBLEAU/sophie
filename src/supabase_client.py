"""Client Supabase singleton."""

import os

from src.logger import get_logger

log = get_logger()

_client = None


def get_supabase():
    """
    Retourne le client Supabase initialisé, ou None si pas configuré.
    Le client est créé une seule fois (singleton).
    """
    global _client

    if _client is not None:
        return _client

    url = os.environ.get("SUPABASE_URL")
    key = os.environ.get("SUPABASE_ANON_KEY")

    if not url or not key:
        log.warning("Supabase non configuré (SUPABASE_URL ou SUPABASE_ANON_KEY manquant)")
        return None

    try:
        from supabase import create_client
        _client = create_client(url, key)
        log.info(f"Client Supabase initialisé — {url}")
        return _client
    except Exception as e:
        log.error(f"Impossible de créer le client Supabase : {e}")
        return None
