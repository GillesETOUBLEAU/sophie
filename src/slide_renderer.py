"""Génération des slides HTML et capture en PNG via Playwright."""

import datetime
import os

from jinja2 import Environment, FileSystemLoader


TEMPLATES_DIR = os.path.join(os.path.dirname(os.path.dirname(__file__)), "templates")
ASSETS_DIR = os.path.join(os.path.dirname(os.path.dirname(__file__)), "assets")

CARDS_PER_SLIDE = 4
VIEWPORT = {"width": 1280, "height": 720}

# Logos WMH (chemins absolus pour file:// dans Playwright)
LOGO_BLANC = os.path.join(ASSETS_DIR, "WMH_Project_horizontal_blanc.png")
LOGO_NOIR = os.path.join(ASSETS_DIR, "WMH_Project_horizontal_noir.png")


def _common_context() -> dict:
    """Variables communes injectées dans tous les templates."""
    return {
        "year": str(datetime.date.today().year),
    }


def build_slide_sequence(events: list[dict], recos: list[dict], week_number: int, week_dates: str) -> list[dict]:
    """
    Construit la séquence ordonnée des slides avec leurs paramètres.
    Retourne une liste de dicts : {template, context, duration_seconds}
    """
    common = _common_context()
    slides = []

    # 1. Slide d'ouverture
    opening_img = os.path.join(ASSETS_DIR, "opening.png")
    if os.path.exists(opening_img):
        slides.append({
            "template": "slide_title.html",
            "context": {**common, "image_path": opening_img, "alt_text": "WMH Project"},
            "duration": 3,
        })
    else:
        slides.append({
            "template": "slide_title.html",
            "context": {
                **common,
                "logo_path": LOGO_BLANC,
                "subtitle": f"Statistiques Écrans — Semaine {week_number:02d}",
            },
            "duration": 3,
        })

    # 2. Header événements
    slides.append({
        "template": "slide_header.html",
        "context": {
            **common,
            "logo_path": LOGO_BLANC,
            "title": "Les événements qui jouent cette semaine",
            "subtitle": "Corporate Event & Institutionnel",
            "week_number": f"{week_number:02d}",
            "week_dates": week_dates,
        },
        "duration": 5,
    })

    # 3. Slides de cartes événements (4 par page)
    event_pages = _paginate(events, CARDS_PER_SLIDE)
    for i, page in enumerate(event_pages):
        slides.append({
            "template": "slide_event_cards.html",
            "context": {
                **common,
                "logo_path": LOGO_NOIR,
                "section_title": "Événements de la semaine",
                "cards": page,
                "page_current": i + 1,
                "page_total": len(event_pages),
            },
            "duration": 10,
        })

    # 4. Header recos
    slides.append({
        "template": "slide_header.html",
        "context": {
            **common,
            "logo_path": LOGO_BLANC,
            "title": "Les recos à rendre cette semaine",
            "subtitle": "Corporate Event & Institutionnel",
            "week_number": f"{week_number:02d}",
            "week_dates": week_dates,
        },
        "duration": 5,
    })

    # 5. Slides de cartes recos (4 par page)
    reco_pages = _paginate(recos, CARDS_PER_SLIDE)
    for i, page in enumerate(reco_pages):
        slides.append({
            "template": "slide_event_cards.html",
            "context": {
                **common,
                "logo_path": LOGO_NOIR,
                "section_title": "Recos à rendre",
                "cards": page,
                "page_current": i + 1,
                "page_total": len(reco_pages),
            },
            "duration": 10,
        })

    # 6. Slide de fermeture
    closing_img = os.path.join(ASSETS_DIR, "closing.png")
    if os.path.exists(closing_img):
        slides.append({
            "template": "slide_title.html",
            "context": {**common, "image_path": closing_img, "alt_text": "WMH Project"},
            "duration": 3,
        })
    else:
        slides.append({
            "template": "slide_title.html",
            "context": {
                **common,
                "logo_path": LOGO_BLANC,
                "subtitle": "We Make It Happen",
            },
            "duration": 3,
        })

    return slides


def _load_css() -> str:
    """Charge le CSS depuis le fichier style.css."""
    css_path = os.path.join(TEMPLATES_DIR, "style.css")
    with open(css_path, "r", encoding="utf-8") as f:
        return f.read()


def render_slides_to_images(slides: list[dict], output_dir: str) -> list[tuple[str, int]]:
    """
    Rend chaque slide en HTML puis capture un screenshot PNG.
    Retourne une liste de (chemin_image, durée_secondes).
    """
    from playwright.sync_api import sync_playwright

    env = Environment(loader=FileSystemLoader(TEMPLATES_DIR))
    os.makedirs(output_dir, exist_ok=True)

    css_content = _load_css()
    image_durations = []

    with sync_playwright() as p:
        browser = p.chromium.launch(headless=True)
        page = browser.new_page(viewport=VIEWPORT)

        for idx, slide in enumerate(slides):
            template = env.get_template(slide["template"])
            context = {**slide["context"], "css_content": css_content}
            html_content = template.render(**context)

            html_path = os.path.join(output_dir, f"slide_{idx:03d}.html")
            with open(html_path, "w", encoding="utf-8") as f:
                f.write(html_content)

            page.goto(f"file://{html_path}")
            page.wait_for_load_state("networkidle")

            img_path = os.path.join(output_dir, f"slide_{idx:03d}.png")
            page.screenshot(path=img_path, full_page=False)

            image_durations.append((img_path, slide["duration"]))

        browser.close()

    return image_durations


def _paginate(items: list, per_page: int) -> list[list]:
    """Découpe une liste en pages de `per_page` éléments."""
    if not items:
        return []
    return [items[i:i + per_page] for i in range(0, len(items), per_page)]
