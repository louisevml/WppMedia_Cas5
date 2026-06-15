"""
Recree ``assets/templates/default.pptx`` a partir de ``template_slide.pptx``.

Le fichier genere conserve le meme theme, les memes masques et toutes les
mises en page (y compris les nouveaux modeles 2 / 3 / 4 graphiques), sans
aucune diapo document.

Usage (depuis le dossier WPPmedia, venv active) :

    .venv\\Scripts\\python scripts\\rebuild_default_from_template.py

    .venv\\Scripts\\python scripts\\rebuild_default_from_template.py ^
        --input assets\\templates\\template_slide.pptx ^
        --output assets\\templates\\default.pptx
"""

from __future__ import annotations

import argparse
import sys
from pathlib import Path

# Permet ``python scripts/rebuild_default_from_template.py`` sans PYTHONPATH
_ROOT = Path(__file__).resolve().parent.parent
if str(_ROOT) not in sys.path:
    sys.path.insert(0, str(_ROOT))

from app.slide_templates import rebuild_default_pptx_from_template  # noqa: E402


def main() -> None:
    root = Path(__file__).resolve().parent.parent
    ap = argparse.ArgumentParser(
        description="Recree default.pptx (theme + masques, 0 diapo) depuis template_slide.pptx."
    )
    ap.add_argument(
        "--input",
        type=Path,
        default=root / "assets" / "templates" / "template_slide.pptx",
        help="PPTX source (gabarit banque de diapos)",
    )
    ap.add_argument(
        "--output",
        type=Path,
        default=root / "assets" / "templates" / "default.pptx",
        help="Fichier default.pptx a ecrire",
    )
    args = ap.parse_args()

    rebuild_default_pptx_from_template(
        template_slide_path=args.input,
        default_path=args.output,
    )


if __name__ == "__main__":
    main()
