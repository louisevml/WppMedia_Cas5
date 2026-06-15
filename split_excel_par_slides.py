#%%
import pandas as pd
import re
from pathlib import Path
from openpyxl import Workbook
#%%
_BASE_DIR = Path(__file__).resolve().parent
_DATA_DIR = _BASE_DIR / ".data"
_DATA_DIR.mkdir(parents=True, exist_ok=True)
INPUT_FILE = _BASE_DIR / "Resultats Etude WPPmedia DTA v2.xlsx"
OUTPUT_FILE = _DATA_DIR / "Resultats_Etude_WPPmedia_DTA_clean.xlsx"
EMPTY_ROW_THRESHOLD = 1
END_OF_FILE_THRESHOLD = 3

#%%
def is_empty_row(row):
 return all(str(cell).strip() in ("", "nan") for cell in row)

#%%
def get_sheet_name(rows, index):
 for row in rows:
    for cell in row:
        val = str(cell).strip()
        if val and val != "nan":
            match_range = re.search(r'[Ss]lide\s+(\d+)\s*[&et]+\s*(\d+)', val)
            match_single = re.search(r'[Ss]lide\s+(\d+)', val)
            if match_range:
                return f"Slide {match_range.group(1)}-{match_range.group(2)}"
            elif match_single:
                return f"Slide {match_single.group(1)}"
            else:
                return f"Bloc_{index + 1}"
 return f"Bloc_{index + 1}"

#%%
# --- Lecture du fichier Excel ---
try:
    df = pd.read_excel(
        INPUT_FILE,
        header=None,
        dtype=str,
        sheet_name=0,
        engine="openpyxl",
    )
    print(f"✅ Lu : {df.shape[0]} lignes x {df.shape[1]} colonnes")
except Exception as e:
    print(f"❌ Echec de lecture : {e}")
    df = None

if df is None or df.empty:
    print("⚠️  Impossible de lire le fichier. Vérifiez le nom et l'emplacement.")
    exit()

rows = df.values.tolist()
#%%
df.head()
#%%
# --- Découpage en blocs ---
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

# Ajoute le dernier bloc si le fichier ne finit pas par 3 lignes vides
if current_block:
    blocks.append(current_block)

print(f"\n📦 {len(blocks)} bloc(s) détecté(s) :")
for i, block in enumerate(blocks):
    name = get_sheet_name(block, i)
    print(f"   {i+1}. {name} ({len(block)} lignes)")

if len(blocks) == 0:
    print("⚠️  Aucun bloc trouvé. Le fichier est peut-être mal formaté.")
else:
    wb = Workbook()
    used_names = {}

    for i, block in enumerate(blocks):
        sheet_name = get_sheet_name(block, i)
        if sheet_name in used_names:
            used_names[sheet_name] += 1
            sheet_name = f"{sheet_name}_{used_names[sheet_name]}"
        else:
            used_names[sheet_name] = 1
        if i == 0:
            ws = wb.active
            ws.title = sheet_name
        else:
            ws = wb.create_sheet(title=sheet_name)
        for row in block:
            ws.append(row)

    wb.save(OUTPUT_FILE)
    print(f"\n✅ Fichier généré : {OUTPUT_FILE}")

# %%
