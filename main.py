"""Point d'entrée du projet — pipeline complet WPPmedia DTA.

Orchestre les deux étapes du projet :

    Étape 1  (split_excel_par_slides)  Excel brut  → Excel nettoyé (.data/)
    Étape 2  (app_powerpoint)          Excel nettoyé → PowerPoint (output_ppt/)

Exécution :

    python main.py              # pipeline complet (étape 1 puis étape 2)
    python main.py --split-only # uniquement la préparation des données
    python main.py --ppt-only   # uniquement la génération du PowerPoint
"""

import argparse
import sys

import app_powerpoint
import split_excel_par_slides


def run_split():
    """Étape 1 : prépare l'Excel nettoyé. Renvoie True si l'étape réussit."""
    print("=" * 70)
    print("ÉTAPE 1 — Préparation des données (Excel brut → Excel nettoyé)")
    print("=" * 70)
    success = split_excel_par_slides.run()
    if not success:
        print("\n❌ Étape 1 échouée : arrêt du pipeline.")
    return success


def run_ppt():
    """Étape 2 : génère le PowerPoint. Renvoie le chemin du fichier produit."""
    print("\n" + "=" * 70)
    print("ÉTAPE 2 — Génération du PowerPoint (Excel nettoyé → PowerPoint)")
    print("=" * 70)
    app_powerpoint.main()


def parse_args(argv=None):
    parser = argparse.ArgumentParser(
        description="Pipeline de génération du PowerPoint WPPmedia DTA.",
    )
    group = parser.add_mutually_exclusive_group()
    group.add_argument(
        "--split-only",
        action="store_true",
        help="Exécute uniquement l'étape 1 (préparation des données).",
    )
    group.add_argument(
        "--ppt-only",
        action="store_true",
        help="Exécute uniquement l'étape 2 (génération du PowerPoint).",
    )
    return parser.parse_args(argv)


def main(argv=None):
    args = parse_args(argv)

    if args.ppt_only:
        run_ppt()
        return 0

    if not run_split():
        return 1

    if args.split_only:
        return 0

    run_ppt()
    return 0


if __name__ == "__main__":
    sys.exit(main())
