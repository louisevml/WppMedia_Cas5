"""
Point d'entrée — génère le PowerPoint WPPmedia à partir :

- du fichier Excel de résultats (onglet ``mapping`` + onglets ``slide_N``),
- des deux gabarits PowerPoint (``assets/templates/``).

Exemple d'exécution :

    python main.py
"""

from __future__ import annotations

from datetime import date
from pathlib import Path

from app.app_ppt_WPP import AppWPPgabarit
from app.chart_styles import apply_chart_style


ROOT = Path(__file__).resolve().parent


def _resolve_asset(filename: str) -> Path:
    """Cherche un gabarit dans ``assets/templates/`` puis dans ``assets/``."""
    for candidate in (
        ROOT / "assets" / "templates" / filename,
        ROOT / "assets" / filename,
    ):
        if candidate.is_file():
            return candidate
    raise SystemExit(
        f"Gabarit introuvable : placez « {filename} » dans "
        f"{ROOT / 'assets' / 'templates'}."
    )


def main() -> None:
    xlsx = ROOT / "data" / "resultats_WPPmedia_test.xlsx"
    if not xlsx.is_file():
        raise SystemExit(f"Fichier Excel introuvable : {xlsx}")

    default_ppt = _resolve_asset("default.pptx")
    template_slide = _resolve_asset("template_slide.pptx")
    output = ROOT / "output" / f"{xlsx.stem}_{date.today():%Y-%m-%d}.pptx"

    app = AppWPPgabarit(
        default_ppt_path=default_ppt,
        template_ppt_path=template_slide,
        excel_path=xlsx,
        output_gabarit_path=output,
        style_callback=apply_chart_style,
    )
    app.run()


if __name__ == "__main__":
    main()
