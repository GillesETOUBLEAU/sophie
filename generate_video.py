#!/usr/bin/env python3
"""
Génération automatique de la vidéo Statistiques Écrans.

Usage:
    python generate_video.py [--data-dir CHEMIN] [--date YYYY-MM-DD]
    python generate_video.py --gagné fichier_gagné.xlsx --pipe fichier_pipe.xlsx

Par défaut :
    --data-dir : ~/Library/CloudStorage/OneDrive-WMHPROJECT/Bureau/TableauSophie
    --date     : aujourd'hui (pour déterminer la semaine courante)
"""

import argparse
import datetime
import os
import tempfile

from src.data_loader import (
    find_excel_files,
    load_projects,
    filter_events_this_week,
    filter_recos_this_week,
    format_event_card,
    format_reco_card,
    get_week_number,
    get_week_bounds,
)
from src.slide_renderer import build_slide_sequence, render_slides_to_images
from src.video_builder import build_video


DEFAULT_DATA_DIR = os.path.expanduser(
    "~/Library/CloudStorage/OneDrive-WMHPROJECT/Bureau/TableauSophie"
)
PROJECT_DIR = os.path.dirname(os.path.abspath(__file__))
OUTPUT_DIR = os.path.join(PROJECT_DIR, "output")


def main():
    parser = argparse.ArgumentParser(description="Génère la vidéo Statistiques Écrans")
    parser.add_argument(
        "--data-dir",
        default=DEFAULT_DATA_DIR,
        help="Dossier contenant les fichiers Excel (Gagné + Pipe)",
    )
    parser.add_argument("--gagné", default=None, help="Chemin direct vers le fichier Gagné .xlsx")
    parser.add_argument("--pipe", default=None, help="Chemin direct vers le fichier Pipe .xlsx")
    parser.add_argument(
        "--date",
        default=None,
        help="Date de référence (YYYY-MM-DD) pour la semaine. Défaut: aujourd'hui.",
    )
    args = parser.parse_args()

    # Date de référence
    if args.date:
        ref_date = datetime.date.fromisoformat(args.date)
    else:
        ref_date = datetime.date.today()

    week_num = get_week_number(ref_date)
    monday, sunday = get_week_bounds(ref_date)
    week_dates = f"{monday.strftime('%d/%m')} — {sunday.strftime('%d/%m/%Y')}"

    print(f"Semaine {week_num:02d} : {week_dates}")

    # 1. Charger les données depuis 2 fichiers séparés
    if args.gagné and args.pipe:
        gagné_path = args.gagné
        pipe_path = args.pipe
    elif args.gagné or args.pipe:
        # Un seul fichier fourni → on l'utilise pour les deux
        single = args.gagné or args.pipe
        gagné_path = single
        pipe_path = single
    else:
        files = find_excel_files(args.data_dir)
        gagné_path = files.get("gagné")
        pipe_path = files.get("pipe")
        print(f"Dossier données : {args.data_dir}")

    # Charger les événements depuis le fichier Gagné
    events = []
    if gagné_path:
        print(f"Fichier Gagné : {os.path.basename(gagné_path)}")
        df_gagné = load_projects(gagné_path)
        events_df = filter_events_this_week(df_gagné, ref_date)
        events = [format_event_card(row) for _, row in events_df.iterrows()]

    # Charger les recos depuis le fichier Pipe
    recos = []
    if pipe_path:
        print(f"Fichier Pipe  : {os.path.basename(pipe_path)}")
        df_pipe = load_projects(pipe_path)
        recos_df = filter_recos_this_week(df_pipe, ref_date)
        recos = [format_reco_card(row) for _, row in recos_df.iterrows()]

    print(f"Événements cette semaine : {len(events)}")
    print(f"Recos à rendre : {len(recos)}")

    if not events and not recos:
        print("Aucun événement ni reco pour cette semaine. Aucune vidéo générée.")
        return

    # 3. Construire la séquence de slides
    slides = build_slide_sequence(events, recos, week_num, week_dates)
    print(f"Slides à générer : {len(slides)}")

    # 4. Rendre les slides en images
    tmp_dir = tempfile.mkdtemp(prefix="sophie_slides_")
    print(f"Rendu des slides dans : {tmp_dir}")

    image_durations = render_slides_to_images(slides, tmp_dir)

    # 5. Assembler la vidéo
    output_path = os.path.join(
        OUTPUT_DIR,
        f"Semaine_{week_num:02d}_Statistiques_ecrans.mp4",
    )
    build_video(image_durations, output_path)

    print(f"\nTerminé ! Vidéo : {output_path}")


if __name__ == "__main__":
    main()
