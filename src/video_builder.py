"""Assemblage des images PNG en vidéo MP4 via ffmpeg."""

from __future__ import annotations

import os
import shutil
import subprocess

from src.logger import get_logger

log = get_logger()

FFMPEG_TIMEOUT = 120  # secondes


def _check_ffmpeg() -> None:
    """Vérifie que ffmpeg est disponible dans le PATH."""
    if not shutil.which("ffmpeg"):
        raise RuntimeError(
            "ffmpeg n'est pas installé ou n'est pas dans le PATH. "
            "Installez-le via : brew install ffmpeg (macOS) ou apt install ffmpeg (Linux)."
        )


def build_video(image_durations: list[tuple[str, int]], output_path: str):
    """
    Assemble une liste de (chemin_image, durée_secondes) en un fichier MP4.
    Utilise un fichier concat ffmpeg pour gérer les durées variables.
    """
    _check_ffmpeg()

    log.info(f"Assemblage de {len(image_durations)} images en vidéo…")

    # Créer le fichier concat pour ffmpeg
    concat_path = os.path.join(os.path.dirname(output_path), "concat.txt")

    with open(concat_path, "w") as f:
        for img_path, duration in image_durations:
            # ffmpeg concat demuxer syntax
            f.write(f"file '{img_path}'\n")
            f.write(f"duration {duration}\n")
        # Répéter la dernière image (nécessaire pour que la dernière durée soit respectée)
        if image_durations:
            f.write(f"file '{image_durations[-1][0]}'\n")

    os.makedirs(os.path.dirname(output_path), exist_ok=True)

    cmd = [
        "ffmpeg", "-y",
        "-f", "concat",
        "-safe", "0",
        "-i", concat_path,
        "-vf", "scale=1280:720",
        "-c:v", "libx264",
        "-preset", "medium",
        "-crf", "23",
        "-pix_fmt", "yuv420p",
        "-movflags", "+faststart",
        output_path,
    ]

    try:
        result = subprocess.run(
            cmd, capture_output=True, text=True, timeout=FFMPEG_TIMEOUT
        )
    except subprocess.TimeoutExpired:
        log.error(f"ffmpeg a dépassé le timeout de {FFMPEG_TIMEOUT}s")
        raise RuntimeError(
            f"La génération vidéo a pris trop de temps (>{FFMPEG_TIMEOUT}s). "
            "Essayez avec moins de slides."
        )
    finally:
        if os.path.exists(concat_path):
            os.remove(concat_path)

    if result.returncode != 0:
        log.error(f"ffmpeg a échoué (code {result.returncode}) :\n{result.stderr[-500:]}")
        raise RuntimeError("La génération vidéo a échoué (ffmpeg). Voir les logs pour le détail.")

    file_size = os.path.getsize(output_path)
    log.info(f"Vidéo générée : {output_path} ({file_size / 1024 / 1024:.1f} MB)")
