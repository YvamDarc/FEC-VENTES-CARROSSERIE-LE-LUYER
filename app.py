import streamlit as st
import pandas as pd
import re
import tempfile
import unicodedata

st.set_page_config(page_title="Lecture journaux XLS", layout="wide")
st.title("Lecture complète d’un journal comptable (XLS édition provisoire)")

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
    """lower + sans accents + espaces normalisés"""
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
    # ex: 171,43 / 0 / 825,40
    return bool(re.fullmatch(r"-?\d{1,3}(\.\d{3})*(,\d{2})|-?\d+(,\d{2})|-?\d+", s0))

def find_header_row(df: pd.DataFrame):
    # On cherche la ligne avec Ecr + Jour + Pièce + Compte + Débit + Crédit (avec ou sans accents)
    for i in range(min(250, len(df))):
        row = " ".join(norm(c) for c in df.iloc[i].tolist())
        if ("ecr" in row) and ("jour" in row) and ("piece" in row) and ("compte" in row) and ("debit" in row) and ("credit" in row):
            return i
    return None

def detect_meta(df: pd.DataFrame):
    journal_code, journal_lib, period = "", "", ""
    for i in range(min(200, len(df))):
        line = " ".join(clean_str(c) for c in df.iloc[i].tolist())
        nline = norm(line)

        if journal_code == "" and "journal" in nline:
            # tente : "Journal 004 Assurance"
            m = re.search(r"\bjournal\b\s+(\d{3})\s+(.+)", line, flags=re.IGNORECASE)
            if m:
                journal_code = m.group(1).strip()
                journal_lib = m.group(2).strip()

        if period == "" and "periode" in nline:
            m = re.search(r"([0-1]?\d\/20\d{2})", line)
            if m:
                period = m.group(1)

    return journal_code, journal_lib, period

def parse_row_as_entry(cells):
    """
    Parse une ligne de tableau en cherchant :
    Ecr = entier (début)
    Jour = 2 chiffres
    Pièce = 6 chiffres
    Compte = 5 ou 6 chiffres après la pièce (ignore 'C')
    Montants = dans toutes les cellules (on garde débit/crédit)
    """
    vals = [clean_str(x) for x in cells]
    joined = " ".join(vals)

    # 1) Ecr / Jour / Pièce
    # on cherche une séquence : Ecr (1-4 chiffres) + jour (2 chiffres) + piece (6 chiffres)
    # dans l’ordre, en scannant les tokens des cellules
    tokens = []
    for v in vals:
        if v != "":
            # split doux, mais garde aussi le token entier si c'est un code
            parts = re.split(r"\s+", v)
            tokens.extend([p for p in parts if p != ""])

    # Find ecr/jour/piece in token stream
    ecr = jour = piece = None
    for i in range(len(tokens) - 2):
        if re.fullmatch(r"\d{1,4}", tokens[i]) and re.fullmatch(r"\d{2}", tokens[i+1]) and re.fullmatch(r"\d{6}", tokens[i+2]):
            ecr = int(tokens[i])
            jour = tokens[i+1]
            piece = tokens[i+2]
            start_idx = i + 3
            break
    if piece is None:
        return None  # pas une ligne d'écriture

    # 2) Compte : premier 5/6 chiffres après la pièce (en ignorant 'C')
    compte = None
    for t in tokens[start_idx:]:
        if t.upper() == "C":
            continue
        if re.fullmatch(r"\d{5,6}", t):
            compte = t
            break

    if compte is None:
        # si pas de compte => on ignore
        return None

    # 3) Montants : récupérer tous les montants présents dans la ligne
    amounts = []
    for v in vals:
        if looks_like_amount(v):
            amounts.append(to_float_fr(v))

    # Heuristique débit/crédit :
    # - sur ce type de journal, on a souvent plusieurs crédits et un débit (client) OU l'inverse selon journal.
    # - Ici on prend : debit = le montant le plus “à gauche” si on arrive à repérer,
    #   sinon : debit = max(amounts) si on voit une seule grande valeur côté client.
    debit = 0.0
    credit = 0.0

    # Try "position" : on regarde l'index de cellule où est le montant max, et la présence d’autres montants
    if amounts:
        # reconstruire (idx_cell, value)
        pos_amounts = []
        for idx, v in enumerate(vals):
            if looks_like_amount(v):
                pos_amounts.append((idx, to_float_fr(v)))

        # On filtre les 0.00 parasites
        pos_amounts_nz = [(i, a) for (i, a) in pos_amounts if abs(a) > 1e-9]

        if len(pos_amounts_nz) == 0:
            debit = 0.0
            credit = 0.0
        elif len(pos_amounts_nz) == 1:
            # Une seule valeur -> on la met en crédit par défaut, sauf si libellé client (compte 41xxx)
            v = pos_amounts_nz[0][1]
            if compte.startswith(("41", "42")):
                debit = v
            else:
                credit = v
        else:
            # Plusieurs valeurs : en vente, généralement crédits multiples + un débit client (plus gros)
            # => le plus gros en débit si compte client 41/42, sinon le plus gros en crédit
            biggest = max(pos_amounts_nz, key=lambda x: abs(x[1]))[1]
            if compte.startswith(("41", "42")):
                debit = biggest
            else:
                # ici on met la valeur de la cellule la plus à droite en crédit (souvent colonne Crédit)
                rightmost = max(pos_amounts_nz, key=lambda x: x[0])[1]
                credit = rightmost
                # et si on n’a que des montants "crédit" (cas standard), OK.

    # 4) Libellé : on garde les morceaux de texte “utiles”
    # (on enlève les tokens purement numériques et 'C')
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

    libelle = " ".join(text_parts)
    libelle = re.sub(r"\s+", " ", libelle).strip()

    return {
        "Ecr": ecr,
        "Jour": jour,
        "Piece": piece,
        "Compte": compte,
        "Libelle": libelle,
        "Debit": float(debit),
        "Credit": float(credit),
    }

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
col1, col2, col3 = st.columns(3)
col1.metric("Journal", journal_code if journal_code else "—")
col2.metric("Libellé", journal_lib if journal_lib else "—")
col3.metric("Période", period if period else "—")

# -------------------------
# Find header + parse all following rows
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

df = pd.DataFrame(entries)

# -------------------------
# Résultats
# -------------------------
if df.empty:
    st.error("Aucune écriture détectée (0 ligne).")
    st.write("Aperçu des 60 lignes après entête pour diagnostic :")
    st.dataframe(data_rows.head(60), use_container_width=True)
    st.stop()

st.subheader("Écritures détectées (lignes)")
st.dataframe(df, use_container_width=True, height=520)

# Contrôles
st.subheader("Contrôle Débit / Crédit par pièce")
control = (
    df.groupby("Piece")[["Debit", "Credit"]]
      .sum()
      .assign(Ecart=lambda x: (x["Debit"] - x["Credit"]).round(2))
      .reset_index()
      .sort_values("Piece")
)
st.dataframe(control, use_container_width=True, height=420)

st.subheader("Contrôle global")
tot_deb = float(df["Debit"].sum())
tot_cre = float(df["Credit"].sum())
st.write({
    "Total Débit": round(tot_deb, 2),
    "Total Crédit": round(tot_cre, 2),
    "Écart": round(tot_deb - tot_cre, 2),
    "Nb lignes": int(len(df)),
    "Nb pièces": int(df["Piece"].nunique()),
})

# Export CSV
st.subheader("Export")
csv = df.to_csv(index=False).encode("utf-8-sig")
st.download_button("Télécharger les lignes (CSV)", data=csv, file_name="journal_lignes.csv", mime="text/csv")
