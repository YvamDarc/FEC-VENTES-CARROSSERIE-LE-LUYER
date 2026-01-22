import streamlit as st
import pandas as pd
import re
import tempfile
from datetime import date

# =============================
# Sécurité anti page blanche
# =============================
st.set_page_config(page_title="Lecture journal XLS", layout="wide")
st.title("Lecture complète d’un journal comptable (édition provisoire)")

st.success("✅ L’application démarre correctement")

# =============================
# Fonctions utilitaires
# =============================

def clean_str(x):
    if x is None:
        return ""
    if isinstance(x, float) and pd.isna(x):
        return ""
    s = str(x)
    if s.lower() == "nan":
        return ""
    return s.strip()

def to_float_fr(x):
    s = clean_str(x)
    if s == "":
        return 0.0
    s = s.replace("\u00a0", "").replace(" ", "").replace(",", ".")
    s = re.sub(r"[^0-9\.-]", "", s)
    try:
        return float(s)
    except:
        return 0.0

def detect_header_row(df):
    """
    On cherche la ligne contenant simultanément :
    Ecr / Jour / Pièce / Compte / Débit / Crédit
    """
    for i in range(min(200, len(df))):
        row = " ".join([clean_str(c).lower() for c in df.iloc[i]])
        if all(k in row for k in ["ecr", "jour", "pi", "compte", "débit", "crédit"]):
            return i
    return None

def detect_period_and_journal(df):
    journal_code, journal_lib, period = "", "", ""
    for i in range(150):
        row = " ".join([clean_str(c) for c in df.iloc[i]])
        if "Journal" in row and journal_code == "":
            m = re.search(r"Journal\s+([0-9]{3})\s+(.+)", row)
            if m:
                journal_code = m.group(1)
                journal_lib = m.group(2).strip()
        if "Période" in row and period == "":
            m = re.search(r"([0-1]?\d\/20\d{2})", row)
            if m:
                period = m.group(1)
    return journal_code, journal_lib, period

# =============================
# Upload fichier
# =============================
uploaded = st.file_uploader("Dépose ton fichier XLS", type=["xls", "xlsx"])

if not uploaded:
    st.stop()

# Sauvegarde temporaire (important pour read_html)
with tempfile.NamedTemporaryFile(delete=False, suffix=".xls") as tmp:
    tmp.write(uploaded.getbuffer())
    path = tmp.name

# =============================
# Lecture XLS (Excel ou HTML)
# =============================
try:
    try:
        df_raw = pd.read_excel(path, header=None)
        read_mode = "excel"
    except:
        tables = pd.read_html(path)
        df_raw = tables[0]
        read_mode = "html"

    st.info(f"Mode de lecture utilisé : {read_mode}")

except Exception as e:
    st.error("Impossible de lire le fichier")
    st.exception(e)
    st.stop()

# =============================
# Détection entête
# =============================
header_row = detect_header_row(df_raw)
if header_row is None:
    st.error("Impossible de détecter l’entête du tableau")
    st.stop()

headers = [clean_str(c) for c in df_raw.iloc[header_row]]
data = df_raw.iloc[header_row + 1:].copy()
data.columns = headers
data = data.dropna(how="all")

# =============================
# Détection colonnes par position
# =============================
cols = [c.lower() for c in headers]

def col_index(keyword):
    for i, c in enumerate(cols):
        if keyword in c:
            return i
    return None

idx_ecr = col_index("ecr")
idx_jour = col_index("jour")
idx_piece = col_index("pi")
idx_compte = col_index("compte")
idx_debit = col_index("débit")
idx_credit = col_index("crédit")
idx_lib = col_index("libell")

# =============================
# Métadonnées
# =============================
journal_code, journal_lib, period = detect_period_and_journal(df_raw)
st.subheader("Métadonnées détectées")
st.write({
    "Journal": journal_code,
    "Libellé": journal_lib,
    "Période": period
})

# =============================
# Lecture des lignes comptables
# =============================
rows = []

for _, r in data.iterrows():
    piece = clean_str(r.iloc[idx_piece])
    compte = clean_str(r.iloc[idx_compte])

    if piece == "" or compte == "":
        continue

    rows.append({
        "Ecriture": piece,
        "Jour": clean_str(r.iloc[idx_jour]),
        "Compte": compte,
        "Libellé": clean_str(r.iloc[idx_lib]) if idx_lib is not None else "",
        "Débit": to_float_fr(r.iloc[idx_debit]),
        "Crédit": to_float_fr(r.iloc[idx_credit]),
    })

df = pd.DataFrame(rows)

# =============================
# Contrôles comptables
# =============================
control = (
    df.groupby("Ecriture")[["Débit", "Crédit"]]
      .sum()
      .assign(Écart=lambda x: (x["Débit"] - x["Crédit"]).round(2))
      .reset_index()
)

# =============================
# AFFICHAGE
# =============================
st.subheader("Lignes comptables lues")
st.dataframe(df, use_container_width=True)

st.subheader("Contrôle Débit / Crédit par écriture")
st.dataframe(control, use_container_width=True)

st.subheader("Contrôle global")
st.write({
    "Total débit": df["Débit"].sum(),
    "Total crédit": df["Crédit"].sum(),
    "Écart": round(df["Débit"].sum() - df["Crédit"].sum(), 2)
})
