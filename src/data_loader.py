"""Lecture des fichiers Excel et filtrage par semaine."""

from __future__ import annotations

import datetime
import glob
import os

import pandas as pd
from openpyxl import load_workbook

from src.logger import get_logger

log = get_logger()

# Mapping des noms de colonnes : nouveau format → nom interne utilisé par l'app
COLUMN_MAP = {
    "N° Toolbox": "N° TOOLBOX",
    "Nom du client": "NOM DU CLIENT",
    "Nom du dossier": "NOM DU DOSSIER",
    "Statut": "STATUT",
    "Métier principal": "MÉTIER PRINCIPAL",
    "Directeur Conseil": "LEAD CONSEIL",
    "Leader Squad": "LEAD PROJET",
    "Début": "DATE DÉBUT",
    "Fin": "DATE FIN",
    "Rendu": "DATE DE RENDU DU DOSSIER",
    "Ville": "VILLE EXPLOITATION",
    "Pax": "NOMBRE DE GUESTS",
}

# Statuts à exclure des recos (annulés / perdus)
EXCLUDED_STATUTS = {"Annulé", "Annule", "Perdu"}


def find_excel_files(data_dir: str) -> dict:
    """
    Trouve les fichiers Excel Gagné et Pipe dans le dossier.
    Retourne {"gagné": path, "pipe": path}.
    """
    result = {}

    # Chercher le fichier Gagné
    gagné_files = sorted(
        glob.glob(os.path.join(data_dir, "*Gagn*s*.xlsx")),
        key=os.path.getmtime, reverse=True,
    )
    if gagné_files:
        result["gagné"] = gagné_files[0]

    # Chercher le fichier Pipe
    pipe_files = sorted(
        glob.glob(os.path.join(data_dir, "*Pipe*.xlsx")),
        key=os.path.getmtime, reverse=True,
    )
    if pipe_files:
        result["pipe"] = pipe_files[0]

    if not result:
        raise FileNotFoundError(f"Aucun fichier Excel (Gagné/Pipe) trouvé dans {data_dir}")

    return result


def _find_data_sheet(excel_path: str, label: str) -> str | int:
    """
    Trouve la feuille de données dans le fichier Excel.
    Cherche "Liste des dossiers" par nom, sinon retombe sur l'index 1.
    """
    try:
        wb = load_workbook(excel_path, read_only=True)
        sheet_names = wb.sheetnames
        wb.close()
    except Exception:
        return 1

    if "Liste des dossiers" in sheet_names:
        return "Liste des dossiers"
    return 1


def validate_excel(excel_path: str, label: str = "Excel") -> None:
    """
    Valide qu'un fichier Excel est exploitable.
    Lève ValueError avec un message clair si ce n'est pas le cas.
    """
    if not os.path.exists(excel_path):
        raise FileNotFoundError(f"Fichier {label} introuvable : {excel_path}")

    try:
        wb = load_workbook(excel_path, read_only=True)
    except Exception:
        raise ValueError(
            f"Le fichier {label} n'est pas un fichier Excel valide. "
            "Vérifiez qu'il s'agit bien d'un .xlsx."
        )

    sheet_names = wb.sheetnames
    wb.close()

    if len(sheet_names) < 2:
        raise ValueError(
            f"Le fichier {label} n'a qu'une seule feuille ({sheet_names[0]}). "
            "Il en faut au moins 2."
        )

    log.info(f"Fichier {label} validé : {len(sheet_names)} feuilles — {', '.join(sheet_names)}")


def _apply_column_mapping(df: pd.DataFrame) -> pd.DataFrame:
    """Renomme les colonnes du nouveau format vers le format interne."""
    rename = {k: v for k, v in COLUMN_MAP.items() if k in df.columns}
    if rename:
        log.info(f"Mapping colonnes appliqué : {list(rename.keys())}")
        df = df.rename(columns=rename)
    return df


def load_projects(excel_path: str, label: str = "Excel") -> pd.DataFrame:
    """Charge la feuille de données détaillées."""
    validate_excel(excel_path, label)

    sheet = _find_data_sheet(excel_path, label)
    sheet_desc = f"'{sheet}'" if isinstance(sheet, str) else f"index {sheet}"

    try:
        df = pd.read_excel(excel_path, sheet_name=sheet)
    except Exception as e:
        raise ValueError(
            f"Impossible de lire la feuille {sheet_desc} du fichier {label} : {e}"
        )

    # Appliquer le mapping des colonnes (nouveau format → format interne)
    df = _apply_column_mapping(df)

    # Colonnes d'intérêt
    cols_needed = [
        "N° TOOLBOX", "NOM DU CLIENT", "NOM DU DOSSIER", "STATUT",
        "MÉTIER PRINCIPAL", "LEAD CONSEIL", "LEAD PROJET",
        "LEAD PRODUCTION", "LEAD LOGISTIQUE",
        "DATE DÉBUT", "DATE FIN", "DATE DE RENDU DU DOSSIER",
        "VILLE EXPLOITATION", "NOMBRE DE GUESTS",
    ]

    # Vérifier la colonne obligatoire
    if "STATUT" not in df.columns:
        raise ValueError(
            f"Colonne 'STATUT' (ou 'Statut') manquante dans le fichier {label}. "
            f"Colonnes trouvées : {', '.join(df.columns[:10])}..."
        )

    existing = [c for c in cols_needed if c in df.columns]
    missing = [c for c in cols_needed if c not in df.columns]
    if missing:
        log.warning(f"Fichier {label} — colonnes absentes (ignorées) : {', '.join(missing)}")

    df = df[existing].copy()

    # Normaliser les dates : convertir time(0,0) en NaT
    for col in ["DATE DÉBUT", "DATE FIN", "DATE DE RENDU DU DOSSIER"]:
        if col in df.columns:
            df[col] = df[col].apply(_normalize_date)

    # Filtrer les lignes vides (STATUT == 0 ou NaN)
    df = df[df["STATUT"].apply(lambda x: isinstance(x, str) and x.strip() != "")]

    log.info(f"Fichier {label} chargé ({sheet_desc}) : {len(df)} lignes avec un statut valide")

    return df


def _normalize_date(val):
    """Convertit les valeurs de date Excel en datetime ou NaT."""
    if isinstance(val, (pd.Timestamp, datetime.datetime)):
        return pd.Timestamp(val)
    return pd.NaT


def get_week_bounds(reference_date: datetime.date = None):
    """Retourne (lundi, dimanche) de la semaine contenant reference_date."""
    if reference_date is None:
        reference_date = datetime.date.today()
    monday = reference_date - datetime.timedelta(days=reference_date.weekday())
    sunday = monday + datetime.timedelta(days=6)
    return (
        pd.Timestamp(monday),
        pd.Timestamp(sunday, hour=23, minute=59, second=59),
    )


def get_week_number(reference_date: datetime.date = None) -> int:
    """Retourne le numéro de semaine ISO."""
    if reference_date is None:
        reference_date = datetime.date.today()
    return reference_date.isocalendar()[1]


def filter_events_this_week(df: pd.DataFrame, reference_date: datetime.date = None) -> pd.DataFrame:
    """
    Événements qui jouent cette semaine :
    STATUT in ('Gagné', 'Gagne', 'Signé', 'Signe') ET l'événement chevauche la semaine courante.
    Un événement chevauche si DATE DÉBUT <= dimanche ET DATE FIN >= lundi.
    """
    monday, sunday = get_week_bounds(reference_date)
    gagné = df[df["STATUT"].isin(["Gagné", "Gagne", "Signé", "Signe"])].copy()

    if "DATE DÉBUT" not in gagné.columns:
        return gagné.iloc[0:0]

    has_start = gagné["DATE DÉBUT"].notna()
    mask = has_start.copy()

    # Événement commence avant ou pendant la semaine
    mask &= gagné["DATE DÉBUT"] <= sunday

    # Événement finit pendant ou après le début de la semaine
    # Si DATE FIN est vide, on considère que c'est un événement d'un jour
    if "DATE FIN" in gagné.columns:
        has_end = gagné["DATE FIN"].notna()
        end_ok = (has_end & (gagné["DATE FIN"] >= monday)) | (~has_end & (gagné["DATE DÉBUT"] >= monday))
    else:
        end_ok = gagné["DATE DÉBUT"] >= monday
    mask &= end_ok

    result = gagné[mask].sort_values("DATE DÉBUT")
    return result


def filter_recos_this_week(df: pd.DataFrame, reference_date: datetime.date = None) -> pd.DataFrame:
    """
    Recos à rendre cette semaine :
    Tous les statuts sauf Annule/Perdu, avec DATE DE RENDU DU DOSSIER dans la semaine.
    """
    monday, sunday = get_week_bounds(reference_date)

    if "DATE DE RENDU DU DOSSIER" not in df.columns:
        return df.iloc[0:0]

    # Exclure les dossiers annulés et perdus
    active = df[~df["STATUT"].isin(EXCLUDED_STATUTS)].copy()

    has_rendu = active["DATE DE RENDU DU DOSSIER"].notna()
    mask = has_rendu & (active["DATE DE RENDU DU DOSSIER"] >= monday) & (active["DATE DE RENDU DU DOSSIER"] <= sunday)

    result = active[mask].sort_values("DATE DE RENDU DU DOSSIER")
    return result


def format_event_card(row: pd.Series) -> dict:
    """Formate une ligne en dictionnaire pour le template de carte."""
    def fmt_date(d):
        if pd.isna(d):
            return ""
        return d.strftime("%d/%m/%Y")

    def clean(val):
        if pd.isna(val) or val == 0 or val == "0":
            return ""
        return str(val).strip()

    start = fmt_date(row.get("DATE DÉBUT"))
    end = fmt_date(row.get("DATE FIN"))
    if start and end and start != end:
        dates = f"{start} → {end}"
    elif start:
        dates = start
    else:
        dates = ""

    return {
        "toolbox": clean(row.get("N° TOOLBOX", "")),
        "client": clean(row.get("NOM DU CLIENT", "")),
        "dossier": clean(row.get("NOM DU DOSSIER", "")),
        "dates": dates,
        "lead_conseil": clean(row.get("LEAD CONSEIL", "")),
        "lead_projet": clean(row.get("LEAD PROJET", "")),
        "lieu": clean(row.get("VILLE EXPLOITATION", "")),
        "guests": clean(row.get("NOMBRE DE GUESTS", "")),
        "metier": clean(row.get("MÉTIER PRINCIPAL", "")),
    }


def format_reco_card(row: pd.Series) -> dict:
    """Formate une ligne reco en dictionnaire pour le template de carte."""
    def fmt_date(d):
        if pd.isna(d):
            return ""
        return d.strftime("%d/%m/%Y")

    def clean(val):
        if pd.isna(val) or val == 0 or val == "0":
            return ""
        return str(val).strip()

    return {
        "toolbox": clean(row.get("N° TOOLBOX", "")),
        "client": clean(row.get("NOM DU CLIENT", "")),
        "dossier": clean(row.get("NOM DU DOSSIER", "")),
        "date_rendu": fmt_date(row.get("DATE DE RENDU DU DOSSIER")),
        "lead_conseil": clean(row.get("LEAD CONSEIL", "")),
        "lead_projet": clean(row.get("LEAD PROJET", "")),
        "lieu": clean(row.get("VILLE EXPLOITATION", "")),
        "guests": clean(row.get("NOMBRE DE GUESTS", "")),
        "metier": clean(row.get("MÉTIER PRINCIPAL", "")),
    }
