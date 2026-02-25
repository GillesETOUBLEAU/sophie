"""Upload et téléchargement de vidéos via Supabase Storage."""

from __future__ import annotations

import os

from src.logger import get_logger
from src.supabase_client import get_supabase

log = get_logger()

BUCKET = os.environ.get("SUPABASE_BUCKET", "videos")


def _ensure_bucket() -> bool:
    """Vérifie que le bucket existe, le crée sinon. Retourne True si OK."""
    sb = get_supabase()
    if not sb:
        return False

    try:
        sb.storage.get_bucket(BUCKET)
        return True
    except Exception:
        try:
            sb.storage.create_bucket(
                BUCKET,
                options={
                    "public": False,
                    "allowed_mime_types": ["video/mp4"],
                    "file_size_limit": 104857600,  # 100 MB
                },
            )
            log.info(f"Bucket '{BUCKET}' créé dans Supabase Storage")
            return True
        except Exception as e:
            log.error(f"Impossible de créer le bucket '{BUCKET}' : {e}")
            return False


def upload_video(local_path: str, remote_name: str) -> str | None:
    """
    Upload un MP4 vers Supabase Storage.
    Retourne le chemin distant si réussi, None sinon.
    """
    sb = get_supabase()
    if not sb:
        log.warning("Supabase non disponible — vidéo conservée en local uniquement")
        return None

    if not _ensure_bucket():
        return None

    try:
        with open(local_path, "rb") as f:
            sb.storage.from_(BUCKET).upload(
                path=remote_name,
                file=f,
                file_options={
                    "content-type": "video/mp4",
                    "upsert": "true",
                },
            )
        log.info(f"Vidéo uploadée vers Supabase Storage : {remote_name}")
        return remote_name
    except Exception as e:
        log.error(f"Échec de l'upload vers Supabase : {e}")
        return None


def get_video_url(remote_name: str, expires_in: int = 3600) -> str | None:
    """
    Génère une URL signée pour télécharger la vidéo.
    Expire après `expires_in` secondes (défaut : 1h).
    """
    sb = get_supabase()
    if not sb:
        return None

    try:
        response = sb.storage.from_(BUCKET).create_signed_url(
            path=remote_name,
            expires_in=expires_in,
        )
        url = response.get("signedURL") or response.get("signedUrl")
        if url:
            log.info(f"URL signée générée pour {remote_name} (expire dans {expires_in}s)")
            return url
        log.warning(f"Pas d'URL signée dans la réponse : {response}")
        return None
    except Exception as e:
        log.error(f"Échec de la génération d'URL signée : {e}")
        return None
