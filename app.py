# app.py
# Streamlit: Multi fichiers éditions journaux (xls/xlsx) -> 1 FEC unique
# Robustesse: repérage des colonnes par index (pas par noms) pour gérer entêtes dupliquées / fusionnées
#
# requirements.txt:
#   streamlit>=1.30
#   pandas>=2.0
#   openpyxl>=3.1
#   xlrd>=2.0.1
#
# (Streamlit Cloud conseillé) runtime.txt:
#   python-3.11

import io
import re
from datetime import date
from typing import Dict, List, Optional, Tuple

import pandas as pd
import streamlit as st


FEC_COLUMNS = [
    "JournalCode", "JournalLib", "EcritureNum", "EcritureDate",
    "CompteNum", "CompteLib", "CompAuxNum", "CompAuxLib",
    "PieceRef", "PieceDate", "EcritureLib",
    "Debit", "Credit",
    "EcritureLet", "DateLet", "ValidDate",
    "Montantdevise", "Idevise"
]


# -----------------------------
# Utils
# -----------------------------
def normalize_space(s: str) -> str:
    return re.sub(r"\s+", " ", (s or "")).strip()


def safe_cell(x) -> str:
    if x is None:
        return ""
    if isinstance(x, float) and pd.isna(x):
        return ""
    s = str(x)
    if s.lower() == "nan":
        return ""
    return s


def to_decimal_fr(x) -> float:
    """Convert '1 002,00' / '556,80' / '0' / NaN -> float."""
    if x is None:
        return 0.0
    if isinstance(x, float) and pd.isna(x):
        return 0.0
    if isinstance(x, (int, float)):
        return float(x)

    s = safe_cell(x).strip()
    if s == "":
        return 0.0

    s = s.replace("\u00A0", " ").replace(" ", "")
    s = s.replace(",", ".")
    s = re.sub(r"[^0-9\.\-]", "", s)
    if s in ("", "-", ".", "-."):
        return 0.0
    try:
        return float(s)
    except ValueError:
        return 0.0


def excel_engine(filename: str) -> str:
    name = (filename or "").lower()
    return "xlrd" if name.endswith(".xls") else "openpyxl"


def extract_compte_num(v: str) -> str:
    s = normalize_space(safe_cell(v))
    s = re.sub(r"^\s*C\s+", "", s, flags=re.IGNORECASE)  # enlève éventuel indicateur "C"
    m = re.search(r"([0-9]{3,})", s)
    return m.group(1) if m else s


def parse_period(period: str) -> Tuple[int, int]:
    m = re.match(r"^\s*([0-1]?\d)\s*/\s*(20\d{2})\s*$", period or "")
    if not m:
        raise ValueError(f"Période introuvable ou format inattendu (attendu MM/YYYY). Trouvé: '{period}'")
    mois = int(m.group(1))
    annee = int(m.group(2))
    if not (1 <= mois <= 12):
        raise ValueError(f"Mois invalide dans la période: {period}")
    return annee, mois


def coerce_day(jour_cell) -> int:
    s = safe_cell(jour_cell)
    s = re.sub(r"\D", "", s)
    if s == "":
        return 1
    try:
        d = int(s)
        return d if 1 <= d <= 31 else 1
    except Exception:
        return 1


# -----------------------------
# Detect journal / period
# -----------------------------
def detect_journal_and_period(df_raw: pd.DataFrame) -> Tuple[str, str, str]:
    """
    Scanne le haut de page pour trouver:
      - JournalCode: 001
      - JournalLib : Ventes et prestations
      - Period     : 12/2025
    """
    journal_code, journal_lib, period = "", "", ""
    max_scan = min(len(df_raw), 140)

    for r in range(max_scan):
        row = [safe_cell(c) for c in df_raw.iloc[r].tolist()]
        line = normalize_space(" ".join([c for c in row if c]))

        if not period:
            m = re.search(r"\bP[ée]riode\b\s*([0-1]?\d\/20\d{2})", line, flags=re.IGNORECASE)
            if m:
                period = m.group(1)

        if not journal_code and ("Journal" in line or "JOURNAL" in line):
            m = re.search(r"\bJournal\b\s*([0-9]{1,3})\s+(.+)$", line, flags=re.IGNORECASE)
            if m:
                journal_code = m.group(1).zfill(3)
                journal_lib = normalize_space(m.group(2))

        # fallback: "001 Ventes et prestations"
        if not journal_code:
            m = re.search(r"\b([0-9]{3})\b\s+([A-Za-zÀ-ÿ].+)", line)
            if m and "Folio" not in line and "Période" not in line:
                if "/" not in m.group(2):
                    journal_code = m.group(1)
                    journal_lib = normalize_space(m.group(2))

    return journal_code, journal_lib, period


# -----------------------------
# Header detection by index
# -----------------------------
def find_header_row_and_indices(df_raw: pd.DataFrame) -> Tuple[int, Dict[str, int]]:
    """
    Trouve la ligne d'entête et renvoie l'index de ligne + mapping des colonnes par position.
    On cherche les cellules qui contiennent : Ecr, Jour, Pièce, Compte, Débit, Crédit
    """
    max_scan = min(len(df_raw), 260)

    def cellnorm(x: str) -> str:
        return normalize_space(safe_cell(x)).lower()

    for r in range(max_scan):
        row = [cellnorm(c) for c in df_raw.iloc[r].tolist()]

        # trouve indices probables
        idx = {}
        for j, v in enumerate(row):
            if "ecr" == v or v.startswith("ecr"):
                idx.setdefault("Ecr", j)
            if v == "jour":
                idx.setdefault("Jour", j)
            if "pièce" in v or "piece" in v or v.startswith("pi"):
                idx.setdefault("Piece", j)
            if "compte" in v:
                idx.setdefault("Compte", j)
            if "débit" in v or "debit" in v:
                idx.setdefault("Debit", j)
            if "crédit" in v or "credit" in v:
                idx.setdefault("Credit", j)
            if "libellé écriture" in v or "libelle ecriture" in v:
                idx.setdefault("LibEcriture", j)

        # règle: si on a les 6 clés principales, c'est bon
        if all(k in idx for k in ["Ecr", "Jour", "Piece", "Compte", "Debit", "Credit"]):
            return r, idx

    raise ValueError("Impossible de trouver la ligne d’entête (Ecr/Jour/Pièce/Compte/Débit/Crédit).")


def guess_compte_lib_col(row_headers: List[str], idx_compte: int, idx_debit: int) -> Optional[int]:
    """
    Dans beaucoup d’éditions, la colonne juste après 'Compte' est le libellé du compte (MO, TVA, etc.)
    Sauf si Débit arrive directement.
    """
    if idx_compte + 1 >= len(row_headers):
        return None
    if idx_compte + 1 == idx_debit:
        return None
    return idx_compte + 1


# -----------------------------
# Parse one sheet
# -----------------------------
def parse_sheet_to_fec(df_raw: pd.DataFrame, file_name: str, sheet_name: str) -> Tuple[pd.DataFrame, Dict[str, str]]:
    journal_code, journal_lib, period = detect_journal_and_period(df_raw)
    journal_code = journal_code or "000"
    journal_lib = journal_lib or "Journal"
    if not period:
        raise ValueError(f"[{file_name} / {sheet_name}] Période introuvable (ex: 12/2025).")

    annee, mois = parse_period(period)

    header_row, idx = find_header_row_and_indices(df_raw)

    # headers row (for guessing CompteLib)
    header_cells = [normalize_space(safe_cell(c)) for c in df_raw.iloc[header_row].tolist()]
    idx_compte_lib = guess_compte_lib_col(header_cells, idx["Compte"], idx["Debit"])

    # data below header
    data = df_raw.iloc[header_row + 1:].copy()
    data = data.dropna(how="all")

    fec_rows = []

    for _, row in data.iterrows():
        # fetch by index safely
        def get_i(i: int):
            if i is None:
                return ""
            try:
                return row.iloc[i]
            except Exception:
                return ""

        ecr = normalize_space(safe_cell(get_i(idx["Ecr"])))
        jour = get_i(idx["Jour"])
        piece = normalize_space(safe_cell(get_i(idx["Piece"])))
        compte_raw = get_i(idx["Compte"])
        compte = extract_compte_num(compte_raw)

        debit = to_decimal_fr(get_i(idx["Debit"]))
        credit = to_decimal_fr(get_i(idx["Credit"]))

        # stop/skip empty movement lines
        if piece == "" or compte == "":
            continue
        if abs(debit) < 0.0001 and abs(credit) < 0.0001:
            # certaines éditions ont des lignes de titre (rare) -> on ignore
            continue

        compte_lib = ""
        if idx_compte_lib is not None:
            compte_lib = normalize_space(safe_cell(get_i(idx_compte_lib)))

        lib_ecr = ""
        if "LibEcriture" in idx:
            lib_ecr = normalize_space(safe_cell(get_i(idx["LibEcriture"])))

        d = coerce_day(jour)
        ecr_date = date(annee, mois, d).strftime("%Y%m%d")

        fec_rows.append({
            "JournalCode": journal_code,
            "JournalLib": journal_lib,
            "EcritureNum": piece if piece else ecr,
            "EcritureDate": ecr_date,
            "CompteNum": compte,
            "CompteLib": compte_lib,
            "CompAuxNum": "",
            "CompAuxLib": "",
            "PieceRef": piece,
            "PieceDate": ecr_date,
            "EcritureLib": lib_ecr,
            "Debit": f"{round(debit, 2):.2f}",
            "Credit": f"{round(credit, 2):.2f}",
            "EcritureLet": "",
            "DateLet": "",
            "ValidDate": "",
            "Montantdevise": "",
            "Idevise": "",
            "_src_file": file_name,
            "_src_sheet": sheet_name,
            "_period": period,
        })

    if not fec_rows:
        raise ValueError(f"[{file_name} / {sheet_name}] Aucune ligne de mouvement détectée sous l’entête.")

    fec = pd.DataFrame(fec_rows)
    return fec, {
        "File": file_name,
        "Sheet": sheet_name,
        "JournalCode": journal_code,
        "JournalLib": journal_lib,
        "Period": period,
        "Rows": str(len(fec)),
    }


# -----------------------------
# FEC output + controls
# -----------------------------
def fec_to_text(df_fec: pd.DataFrame) -> str:
    out = io.StringIO()
    out.write("|".join(FEC_COLUMNS) + "\n")
    for _, r in df_fec.iterrows():
        out.write("|".join([str(r.get(c, "") if not pd.isna(r.get(c, "")) else "") for c in FEC_COLUMNS]) + "\n")
    return out.getvalue()


def balance_controls(df_fec: pd.DataFrame) -> Tuple[pd.DataFrame, Dict[str, float]]:
    df = df_fec.copy()
    df["_debit"] = df["Debit"].astype(str).replace("", "0").astype(float)
    df["_credit"] = df["Credit"].astype(str).replace("", "0").astype(float)

    by_piece = df.groupby(
        ["JournalCode", "JournalLib", "Period", "PieceRef", "EcritureDate"],
        dropna=False
    ).agg(
        Debit=("_debit", "sum"),
        Credit=("_credit", "sum"),
        Lignes=("PieceRef", "size"),
        Fichier=("_src_file", "first"),
        Feuille=("_src_sheet", "first"),
    ).reset_index()

    by_piece["Ecart"] = (by_piece["Debit"] - by_piece["Credit"]).round(2)
    by_piece = by_piece.sort_values("Ecart", key=lambda s: s.abs(), ascending=False)

    tot_deb = float(df["_debit"].sum())
    tot_cred = float(df["_credit"].sum())
    return by_piece, {
        "TotalDebit": tot_deb,
        "TotalCredit": tot_cred,
        "TotalEcart": round(tot_deb - tot_cred, 2),
    }


# -----------------------------
# Streamlit UI
# -----------------------------
st.set_page_config(page_title="Éditions -> FEC unique", layout="wide")
st.title("Éditions de journaux (XLS/XLSX) ➜ FEC unique (multi fichiers)")

st.write(
    "Charge plusieurs fichiers d’éditions (journaux et périodes différents). "
    "L’app convertit tout en un **seul FEC** et réalise un **contrôle Débit/Crédit**."
)

uploaded_files = st.file_uploader(
    "Fichiers XLS/XLSX (plusieurs)",
    type=["xls", "xlsx"],
    accept_multiple_files=True
)

c1, c2 = st.columns([1, 2])
with c1:
    parse_all_sheets = st.checkbox("Parser toutes les feuilles", value=True)
with c2:
    strict_block = st.checkbox("Bloquer le téléchargement si déséquilibre global", value=False)

if uploaded_files:
    fec_parts: List[pd.DataFrame] = []
    metas: List[Dict[str, str]] = []
    errors: List[str] = []

    with st.spinner("Conversion en cours…"):
        for uf in uploaded_files:
            file_name = uf.name
            eng = excel_engine(file_name)

            try:
                xls = pd.ExcelFile(uf, engine=eng)
            except Exception as e:
                errors.append(f"[{file_name}] Impossible d’ouvrir le fichier (engine={eng}) : {e}")
                continue

            sheets = list(xls.sheet_names)
            sheets_to_parse = sheets if parse_all_sheets else [sheets[0]]

            for sh in sheets_to_parse:
                try:
                    df_raw = pd.read_excel(xls, sheet_name=sh, header=None)  # uses xls engine
                    fec_df, meta = parse_sheet_to_fec(df_raw, file_name, sh)
                    fec_parts.append(fec_df)
                    metas.append(meta)
                except Exception as e:
                    errors.append(str(e))

    st.subheader("Fichiers / feuilles convertis")
    if metas:
        st.dataframe(pd.DataFrame(metas), use_container_width=True)
    else:
        st.warning("Aucune donnée convertie. Regarde la section erreurs.")

    if errors:
        with st.expander(f"Erreurs / feuilles ignorées ({len(errors)})", expanded=not metas):
            for err in errors:
                st.error(err)

    if fec_parts:
        fec_all = pd.concat(fec_parts, ignore_index=True)

        # Ajoute Period dans le DF pour les contrôles
        if "Period" not in fec_all.columns:
            fec_all["Period"] = fec_all["_period"]
        else:
            fec_all["Period"] = fec_all["Period"].fillna(fec_all["_period"])

        # Compose FEC standard (colonnes strictes)
        # (on garde Period en interne pour les contrôles, mais pas dans le fichier export)
        for c in FEC_COLUMNS:
            if c not in fec_all.columns:
                fec_all[c] = ""

        st.subheader("Contrôle Débit / Crédit")
        by_piece, totals = balance_controls(fec_all)

        st.write(
            f"Total Débit = **{totals['TotalDebit']:.2f}** ; "
            f"Total Crédit = **{totals['TotalCredit']:.2f}** ; "
            f"Écart = **{totals['TotalEcart']:.2f}**"
        )

        st.write("Top écarts par pièce (les plus gros déséquilibres en premier) :")
        st.dataframe(by_piece.head(60), use_container_width=True)

        st.subheader("Aperçu FEC (50 premières lignes)")
        st.dataframe(fec_all[FEC_COLUMNS].head(50), use_container_width=True)

        fec_text = fec_to_text(fec_all[FEC_COLUMNS])

        can_download = True
        if strict_block and abs(totals["TotalEcart"]) > 0.0001:
            can_download = False
            st.error("Téléchargement bloqué : le FEC global n’est pas équilibré (Débit ≠ Crédit).")

        if can_download:
            st.download_button(
                "Télécharger le FEC unique",
                data=fec_text.encode("utf-8"),
                file_name="FEC_multi_journaux.txt",
                mime="text/plain"
            )

        with st.expander("Voir le début du FEC (texte)"):
            st.code("\n".join(fec_text.splitlines()[:50]), language="text")

else:
    st.info("Charge un ou plusieurs fichiers Excel pour générer un FEC unique.")
