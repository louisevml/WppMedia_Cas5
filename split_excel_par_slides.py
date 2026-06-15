"""Étape 1 — Préparation des données.

Lit le fichier Excel brut (une seule feuille plate) et découpe son contenu
en blocs séparés par des lignes vides. Chaque bloc correspond à un tableau
de résultats pour une slide donnée. Le résultat est écrit dans un Excel nettoyé
(un onglet par slide) dans le dossier .data/.
"""

import re
from pathlib import Path

import pandas as pd
from openpyxl import Workbook

_BASE_DIR = Path(__file__).resolve().parent
_DATA_DIR = _BASE_DIR / ".data"
_DATA_DIR.mkdir(parents=True, exist_ok=True)

INPUT_FILE = _DATA_DIR / "Resultats Etude WPPmedia DTA v2.xlsx"
OUTPUT_FILE = _DATA_DIR / "Resultats_Etude_WPPmedia_DTA_clean.xlsx"

EMPTY_ROW_THRESHOLD = 1
END_OF_FILE_THRESHOLD = 3

SLIDE_RANGE_PATTERN = re.compile(
    r"^[Ss]lide\s+(\d+)\s*[-–—&/etETàÀauU]+\s*(\d+)$",
)
SLIDE_SINGLE_PATTERN = re.compile(r"^[Ss]lide\s+(\d+)$")


def is_empty_row(row):
    """Renvoie True si toutes les cellules de la ligne sont vides ou NaN."""
    return all(str(cell).strip() in ("", "nan") for cell in row)


def get_sheet_name(rows, index):
    """Détecte le nom de la slide depuis la première cellule non vide du bloc.

    Cherche un motif "Slide X" ou "Slide X-Y" dans les premières cellules.
    Retourne "Bloc_N" si aucune référence de slide n'est trouvée.
    """
    for row in rows:
        for cell in row:
            val = str(cell).strip()
            if val and val != "nan":
                match_range = re.search(r'[Ss]lide\s+(\d+)\s*[&et\-]+\s*(\d+)', val)
                match_single = re.search(r'[Ss]lide\s+(\d+)', val)
                if match_range:
                    return f"Slide {match_range.group(1)}-{match_range.group(2)}"
                elif match_single:
                    return f"Slide {match_single.group(1)}"
                else:
                    return f"Bloc_{index + 1}"
    return f"Bloc_{index + 1}"


def expand_sheet_names(sheet_name):
    """Décompose un nom d'onglet en un onglet par numéro de slide cible.

    Exemples :
    - 'Slide 7'     → ['Slide 7']
    - 'Slide 16-17' → ['Slide 16', 'Slide 17']
    - 'Bloc_3'      → ['Bloc_3']  (inchangé si pas de référence slide)
    """
    name = sheet_name.strip()

    range_match = SLIDE_RANGE_PATTERN.match(name)
    if range_match:
        start, end = int(range_match.group(1)), int(range_match.group(2))
        if start > end:
            start, end = end, start
        return [f"Slide {n}" for n in range(start, end + 1)]

    single_match = SLIDE_SINGLE_PATTERN.match(name)
    if single_match:
        return [f"Slide {single_match.group(1)}"]

    return [name]


def split_into_blocks(rows):
    """Découpe la liste de lignes en blocs séparés par des lignes vides.

    Un seul saut de ligne vide délimite deux blocs.
    Trois lignes vides consécutives signalent la fin du fichier.
    """
    blocks = []
    current_block = []
    consecutive_empty = 0

    for row in rows:
        if is_empty_row(row):
            consecutive_empty += 1
            if consecutive_empty >= END_OF_FILE_THRESHOLD:
                if current_block:
                    blocks.append(current_block)
                    current_block = []
                break
            elif consecutive_empty == EMPTY_ROW_THRESHOLD:
                if current_block:
                    blocks.append(current_block)
                    current_block = []
        else:
            consecutive_empty = 0
            current_block.append(row)

    if current_block:
        blocks.append(current_block)

    return blocks


def write_blocks_to_excel(blocks, output_file):
    """Écrit chaque bloc dans un ou plusieurs onglets (un onglet = une slide).

    Si un bloc couvre plusieurs slides (ex. « Slide 16-17 »), le même tableau
    est dupliqué dans un onglet distinct par numéro de slide.
    """
    wb = Workbook()
    used_names = {}
    first_sheet = True

    for i, block in enumerate(blocks):
        base_name = get_sheet_name(block, i)
        sheet_names = expand_sheet_names(base_name)

        for sheet_name in sheet_names:
            final_name = sheet_name
            if final_name in used_names:
                used_names[final_name] += 1
                final_name = f"{sheet_name}_{used_names[final_name]}"
            else:
                used_names[final_name] = 1

            if first_sheet:
                ws = wb.active
                ws.title = final_name
                first_sheet = False
            else:
                ws = wb.create_sheet(title=final_name)

            for row in block:
                ws.append(row)

    wb.save(output_file)


def run(input_file=INPUT_FILE, output_file=OUTPUT_FILE):
    """Lit le fichier Excel brut, découpe en blocs et écrit le fichier nettoyé."""
    try:
        df = pd.read_excel(
            input_file,
            header=None,
            dtype=str,
            sheet_name=0,
            engine="openpyxl",
        )
        print(f"✅ Lu : {df.shape[0]} lignes x {df.shape[1]} colonnes")
    except Exception as e:
        print(f"❌ Echec de lecture : {e}")
        return False

    if df is None or df.empty:
        print("⚠️  Fichier vide ou illisible.")
        return False

    rows = df.values.tolist()
    blocks = split_into_blocks(rows)

    print(f"\n📦 {len(blocks)} bloc(s) détecté(s) :")
    total_sheets = 0
    for i, block in enumerate(blocks):
        base_name = get_sheet_name(block, i)
        sheet_names = expand_sheet_names(base_name)
        total_sheets += len(sheet_names)
        if len(sheet_names) > 1:
            print(
                f"   {i + 1}. {base_name} ({len(block)} lignes)"
                f" → {len(sheet_names)} onglets : {', '.join(sheet_names)}"
            )
        else:
            print(f"   {i + 1}. {sheet_names[0]} ({len(block)} lignes)")

    if not blocks:
        print("⚠️  Aucun bloc trouvé. Le fichier est peut-être mal formaté.")
        return False

    write_blocks_to_excel(blocks, output_file)
    print(f"\n✅ Fichier généré : {output_file} ({total_sheets} onglet(s))")
    return True


def main():
    run()


if __name__ == "__main__":
    main()
