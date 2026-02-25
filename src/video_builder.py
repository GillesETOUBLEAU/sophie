"""Assemblage des images PNG en vidéo MP4 via ffmpeg."""

import os
import subprocess
import tempfile


def build_video(image_durations: list[tuple[str, int]], output_path: str):
    """
    Assemble une liste de (chemin_image, durée_secondes) en un fichier MP4.
    Utilise un fichier concat ffmpeg pour gérer les durées variables.
    """
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

    result = subprocess.run(cmd, capture_output=True, text=True)

    # Nettoyage
    os.remove(concat_path)

    if result.returncode != 0:
        raise RuntimeError(f"ffmpeg a échoué :\n{result.stderr}")

    file_size = os.path.getsize(output_path)
    print(f"Vidéo générée : {output_path} ({file_size / 1024 / 1024:.1f} MB)")
