import streamlit as st
import pandas as pd
import re
import tempfile
import unicodedata
from datetime import datetime

st.set_page_config(page_title="Lecture journaux XLS → FEC", layout="wide")
st.title("Lecture complète d’un journal comptable (XLS) + Export FEC (TAB)")

# -------------------------
# Utils
# -------------------------
def clean_str(x) -> str:
    if x is None:
        return ""
    if isinstance(x, float) and pd.isna(x):
        return ""
    s = str(x)
    if s.lower() == "nan":
        return ""
    return s.replace("\u00a0", " ").strip()

def norm(s: str) -> str:
    s = clean_str(s).lower()
    s = unicodedata.normalize("NFD", s)
    s = "".join(ch for ch in s if unicodedata.category(ch) != "Mn")  # remove accents
    s = re.sub(r"\s+", " ", s).strip()
    return s

def to_float_fr(x) -> float:
    s = clean_str(x)
    if s == "":
        return 0.0
    s = s.replace(" ", "").replace("\u00a0", "").replace(",", ".")
    s = re.sub(r"[^0-9\.-]", "", s)
    try:
        return float(s)
    except:
        return 0.0

def looks_like_amount(s: str) -> bool:
    s0 = clean_str(s)
    if s0 == "":
        return False
    return bool(re.fullmatch(r"-?\d{1,3}(\.\d{3})*(,\d{2})|-?\d+(,\d{2})|-?\d+", s0))

def find_header_row(df: pd.DataFrame):
    for i in range(min(300, len(df))):
        row = " ".join(norm(c) for c in df.iloc[i].tolist())
        if ("ecr" in row) and ("jour" in row) and ("piece" in row) and ("compte" in row) and ("debit" in row) and ("credit" in row):
            return i
    return None

def detect_meta(df: pd.DataFrame):
    """
    Détection robuste des métadonnées :
    - Journal : code (3 chiffres) + libellé
    - Période : MM/YYYY
    """
    journal_code, journal_lib, period = "", "", ""

    # 1) Détection "Période"
    for i in range(min(250, len(df))):
        line = " ".join(clean_str(c) for c in df.iloc[i].tolist())
        nline = norm(line)
        if period == "" and "periode" in nline:
            m = re.search(r"([0-1]?\d\/20\d{2})", line)
            if m:
                period = m.group(1).strip()

    # 2) Détection "Journal" via scan cellule par cellule
    # On cherche une cellule contenant exactement "Journal" (ou proche),
    # puis on prend les cellules non vides suivantes sur la même ligne.
    for i in range(min(250, len(df))):
        row = [clean_str(x) for x in df.iloc[i].tolist()]
        row_n = [norm(x) for x in row]
        # index de cellule "journal"
        idxs = [k for k, v in enumerate(row_n) if v == "journal" or v.startswith("journal ")]
        if not idxs:
            continue
        k = idxs[0]
        # Récupérer les cellules non vides suivantes (code, libellé)
        tail = [row[j] for j in range(k + 1, len(row)) if clean_str(row[j]) != ""]
        if tail:
            # code = premier token numérique 3 chiffres trouvé
            m = re.search(r"\b(\d{3})\b", " ".join(tail))
            if m:
                journal_code = m.group(1)
                # libellé = texte après le code si possible
                # ex: "004 Assurance"
                after = " ".join(tail)
                after = re.sub(r"\s+", " ", after).strip()
                # retire le code trouvé
                journal_lib = re.sub(rf"\b{re.escape(journal_code)}\b", "", after, count=1).strip(" -\t")
                # si libellé vide, on laisse ""
                break

    # 3) Fallback : pattern "Journal 004 Assurance"
    if journal_code == "":
        for i in range(min(250, len(df))):
            line = " ".join(clean_str(c) for c in df.iloc[i].tolist())
            m = re.search(r"\bjournal\b\s+(\d{3})\s+(.+)", line, flags=re.IGNORECASE)
            if m:
                journal_code = m.group(1).strip()
                journal_lib = m.group(2).strip()
                break

    return journal_code, journal_lib, period

def period_jour_to_fec_date(period: str, jour: str) -> str:
    """
    period: '10/2025'  jour: '09' -> '20251009' (YYYYMMDD)
    """
    p = clean_str(period)
    j = clean_str(jour)
    m = re.fullmatch(r"([0-1]?\d)\/(20\d{2})", p)
    if not m or not re.fullmatch(r"\d{2}", j):
        return ""
    mm = int(m.group(1))
    yyyy = int(m.group(2))
    dd = int(j)
    try:
        dt = datetime(yyyy, mm, dd)
        return dt.strftime("%Y%m%d")
    except:
        return ""

def parse_row_as_entry(cells):
    """
    Parse une ligne en mode "heuristique robuste".
    On récupère :
    - Ecr / Jour / Piece
    - Compte (5/6 digits) après la pièce (ignore 'C')
    - Débit / Crédit (heuristique)
    - Libellé (texte concat)
    """
    vals = [clean_str(x) for x in cells]

    # Tokens
    tokens = []
    for v in vals:
        if v != "":
            parts = re.split(r"\s+", v)
            tokens.extend([p for p in parts if p != ""])

    # Find Ecr / Jour / Piece
    ecr = jour = piece = None
    start_idx = None
    for i in range(len(tokens) - 2):
        if re.fullmatch(r"\d{1,4}", tokens[i]) and re.fullmatch(r"\d{2}", tokens[i+1]) and re.fullmatch(r"\d{6}", tokens[i+2]):
            ecr = int(tokens[i])
            jour = tokens[i+1]
            piece = tokens[i+2]
            start_idx = i + 3
            break
    if piece is None:
        return None

    # Compte : premier 5/6 digits après pièce (ignore 'C')
    compte = None
    for t in tokens[start_idx:]:
        if t.upper() == "C":
            continue
        if re.fullmatch(r"\d{5,6}", t):
            compte = t
            break
    if compte is None:
        return None

    # Montants
    pos_amounts = []
    for idx, v in enumerate(vals):
        if looks_like_amount(v):
            pos_amounts.append((idx, to_float_fr(v)))
    pos_amounts_nz = [(i, a) for (i, a) in pos_amounts if abs(a) > 1e-9]

    debit = 0.0
    credit = 0.0

    if len(pos_amounts_nz) == 1:
        v = pos_amounts_nz[0][1]
        # Heuristique : si compte client (41/42), plutôt débit, sinon crédit
        if compte.startswith(("41", "42")):
            debit = v
        else:
            credit = v
    elif len(pos_amounts_nz) >= 2:
        biggest = max(pos_amounts_nz, key=lambda x: abs(x[1]))[1]
        # si compte client, on met le plus gros en débit
        if compte.startswith(("41", "42")):
            debit = biggest
        else:
            # sinon, la valeur la plus à droite est généralement le crédit
            rightmost = max(pos_amounts_nz, key=lambda x: x[0])[1]
            credit = rightmost

    # Libellé : concat texte non-numérique
    text_parts = []
    for v in vals:
        v2 = clean_str(v)
        if v2 == "":
            continue
        if looks_like_amount(v2):
            continue
        if re.fullmatch(r"\d{1,6}", v2):
            continue
        if v2.upper() == "C":
            continue
        text_parts.append(v2)

    libelle = re.sub(r"\s+", " ", " ".join(text_parts)).strip()

    return {
        "Ecr": ecr,
        "Jour": jour,
        "Piece": piece,
        "Compte": compte,
        "Libelle": libelle,
        "Debit": float(debit),
        "Credit": float(credit),
    }

def make_fec(df_lines: pd.DataFrame, journal_code: str, journal_lib: str, period: str) -> pd.DataFrame:
    """
    Crée un DataFrame FEC standard.
    Colonnes (ordre standard) :
    JournalCode, JournalLib, EcritureNum, EcritureDate, CompteNum, CompteLib,
    CompAuxNum, CompAuxLib, PieceRef, PieceDate, EcritureLib, Debit, Credit,
    EcritureLet, DateLet, ValidDate, Montantdevise, Idevise
    """
    df = df_lines.copy()

    df["JournalCode"] = journal_code
    df["JournalLib"] = journal_lib

    # Num écriture : unique + lisible
    df["EcritureNum"] = df.apply(lambda r: f"{journal_code}-{r['Piece']}-{r['Ecr']}", axis=1)

    # Date écriture FEC
    df["EcritureDate"] = df["Jour"].apply(lambda j: period_jour_to_fec_date(period, j))

    df["CompteNum"] = df["Compte"]
    df["CompteLib"] = ""  # non fourni de façon fiable dans l’export
    df["CompAuxNum"] = ""
    df["CompAuxLib"] = ""

    df["PieceRef"] = df["Piece"]
    df["PieceDate"] = df["EcritureDate"]

    df["EcritureLib"] = df["Libelle"]

    # FEC attend souvent le point en séparateur décimal dans le fichier
    df["Debit"] = df["Debit"].round(2)
    df["Credit"] = df["Credit"].round(2)

    df["EcritureLet"] = ""
    df["DateLet"] = ""
    df["ValidDate"] = df["EcritureDate"]
    df["Montantdevise"] = ""
    df["Idevise"] = ""

    cols = [
        "JournalCode", "JournalLib", "EcritureNum", "EcritureDate",
        "CompteNum", "CompteLib", "CompAuxNum", "CompAuxLib",
        "PieceRef", "PieceDate", "EcritureLib",
        "Debit", "Credit",
        "EcritureLet", "DateLet", "ValidDate",
        "Montantdevise", "Idevise"
    ]
    return df[cols]

# -------------------------
# Upload
# -------------------------
uploaded = st.file_uploader("Dépose ton fichier .xls / .xlsx", type=["xls", "xlsx"])
if not uploaded:
    st.stop()

with tempfile.NamedTemporaryFile(delete=False, suffix=".xls") as tmp:
    tmp.write(uploaded.getbuffer())
    path = tmp.name

# -------------------------
# Lecture : Excel ou HTML déguisé
# -------------------------
df_raw = None
read_mode = None

try:
    df_raw = pd.read_excel(path, header=None)
    read_mode = "excel"
except Exception:
    try:
        tables = pd.read_html(path)
        df_raw = tables[0]
        read_mode = "html"
    except Exception as e:
        st.error("Impossible de lire le fichier (ni Excel, ni HTML).")
        st.exception(e)
        st.stop()

st.caption(f"Mode lecture : {read_mode}")

# -------------------------
# Métadonnées
# -------------------------
journal_code, journal_lib, period = detect_meta(df_raw)

# Fallback manuel (colonne de remplissage)
with st.sidebar:
    st.header("Paramètres (fallback)")
    jc = st.text_input("JournalCode (si non détecté)", value=journal_code or "")
    jl = st.text_input("JournalLib (si non détecté)", value=journal_lib or "")
    pr = st.text_input("Période MM/YYYY (si non détectée)", value=period or "")

# on remplace si l’utilisateur remplit
journal_code = jc.strip()
journal_lib = jl.strip()
period = pr.strip()

c1, c2, c3 = st.columns(3)
c1.metric("Journal", journal_code if journal_code else "—")
c2.metric("Libellé", journal_lib if journal_lib else "—")
c3.metric("Période", period if period else "—")

# -------------------------
# Find header + parse
# -------------------------
header_row = find_header_row(df_raw)
if header_row is None:
    st.error("Je ne trouve pas la ligne d’entête (Ecr / Jour / Pièce / Compte / Débit / Crédit).")
    st.write("Aperçu des 40 premières lignes pour diagnostic :")
    st.dataframe(df_raw.head(40), use_container_width=True)
    st.stop()

st.success(f"Entête détectée ligne : {header_row+1}")

data_rows = df_raw.iloc[header_row + 1:].copy()

entries = []
for _, row in data_rows.iterrows():
    entry = parse_row_as_entry(row.tolist())
    if entry:
        entries.append(entry)

df_lines = pd.DataFrame(entries)

if df_lines.empty:
    st.error("Aucune écriture détectée (0 ligne).")
    st.write("Aperçu des 60 lignes après entête pour diagnostic :")
    st.dataframe(data_rows.head(60), use_container_width=True)
    st.stop()

# Ajout colonnes journal sur les lignes
df_lines["JournalCode"] = journal_code
df_lines["JournalLib"] = journal_lib

st.subheader("Écritures détectées (lignes)")
st.dataframe(df_lines, use_container_width=True, height=520)

# Contrôle par pièce
st.subheader("Contrôle Débit / Crédit par pièce")
control = (
    df_lines.groupby("Piece")[["Debit", "Credit"]]
      .sum()
      .assign(Ecart=lambda x: (x["Debit"] - x["Credit"]).round(2))
      .reset_index()
      .sort_values("Piece")
)
st.dataframe(control, use_container_width=True, height=420)

st.subheader("Contrôle global")
tot_deb = float(df_lines["Debit"].sum())
tot_cre = float(df_lines["Credit"].sum())
st.write({
    "Total Débit": round(tot_deb, 2),
    "Total Crédit": round(tot_cre, 2),
    "Écart": round(tot_deb - tot_cre, 2),
    "Nb lignes": int(len(df_lines)),
    "Nb pièces": int(df_lines["Piece"].nunique()),
})

# -------------------------
# Export FEC
# -------------------------
st.subheader("Export FEC (TAB)")

fec_df = make_fec(df_lines, journal_code, journal_lib, period)

st.caption("Aperçu FEC")
st.dataframe(fec_df, use_container_width=True, height=420)

# FEC en tabulation
# - separateur = \t
# - décimales avec point (pandas le fait), pas de milliers
# - en-tête inclus
fec_txt = fec_df.to_csv(sep="\t", index=False, encoding="utf-8", lineterminator="\n")

fname = "FEC_export.txt"
if journal_code:
    fname = f"FEC_{journal_code}_{period.replace('/','-') if period else 'periode'}_{datetime.now().strftime('%Y%m%d_%H%M%S')}.txt"

st.download_button(
    "Télécharger le FEC (TAB .txt)",
    data=fec_txt.encode("utf-8"),
    file_name=fname,
    mime="text/plain"
)
