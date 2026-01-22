# app.py
# Streamlit app: Convert multiple "éditions de journaux" Excel (xls/xlsx) into a single FEC text file
# - Supports multiple uploaded files, different journals and periods
# - Detects header area automatically
# - Builds FEC lines and concatenates into one output
# - Adds per-piece and global debit/credit controls
#
# Requirements:
#   streamlit, pandas, openpyxl, xlrd
# Recommended (Streamlit Cloud): runtime.txt -> python-3.11

import io
import re
from datetime import date
from typing import Dict, List, Optional, Tuple

import pandas as pd
import streamlit as st


# -----------------------------
# FEC columns (standard order)
# -----------------------------
FEC_COLUMNS = [
    "JournalCode", "JournalLib", "EcritureNum", "EcritureDate",
    "CompteNum", "CompteLib", "CompAuxNum", "CompAuxLib",
    "PieceRef", "PieceDate", "EcritureLib",
    "Debit", "Credit",
    "EcritureLet", "DateLet", "ValidDate",
    "Montantdevise", "Idevise"
]


# -----------------------------
# Helpers
# -----------------------------
def normalize_space(s: str) -> str:
    return re.sub(r"\s+", " ", (s or "")).strip()


def to_decimal_fr(x) -> float:
    """Convert '1 002,00' / '556,80' / 556.80 / NaN -> float (or 0)."""
    if x is None:
        return 0.0
    if isinstance(x, float) and pd.isna(x):
        return 0.0
    if isinstance(x, (int, float)):
        return float(x)

    s = str(x).strip()
    if s == "" or s.lower() == "nan":
        return 0.0

    # remove NBSP and spaces (thousands)
    s = s.replace("\u00A0", " ").replace(" ", "")
    # comma decimal -> dot
    s = s.replace(",", ".")
    # keep digits / dot / minus
    s = re.sub(r"[^0-9\.\-]", "", s)

    if s in ("", "-", ".", "-."):
        return 0.0
    try:
        return float(s)
    except ValueError:
        return 0.0


def excel_engine_for_filename(name: str) -> str:
    name = (name or "").lower()
    if name.endswith(".xls"):
        return "xlrd"
    return "openpyxl"


def safe_str_cell(x) -> str:
    if x is None:
        return ""
    if isinstance(x, float) and pd.isna(x):
        return ""
    s = str(x)
    if s.lower() == "nan":
        return ""
    return s


def extract_compte_num(v: str) -> str:
    """Handle 'C 30020' or '30020' or ' 30020 ' -> '30020'."""
    s = normalize_space(safe_str_cell(v))
    s = re.sub(r"^\s*C\s+", "", s, flags=re.IGNORECASE)
    m = re.search(r"([0-9]{3,})", s)
    return m.group(1) if m else s


def detect_period_journal(df_raw: pd.DataFrame) -> Tuple[str, str, str]:
    """
    Try to detect:
      - JournalCode (e.g., 001)
      - JournalLib (e.g., Ventes et prestations)
      - Period (e.g., 12/2025)
    by scanning top rows.
    """
    journal_code = ""
    journal_lib = ""
    period = ""

    max_scan = min(len(df_raw), 120)
    for r in range(max_scan):
        row = [safe_str_cell(c) for c in df_raw.iloc[r].tolist()]
        line = normalize_space(" ".join([c for c in row if c]))

        if not period:
            m = re.search(r"\bP[ée]riode\b\s*([0-1]?\d\/20\d{2})", line, flags=re.IGNORECASE)
            if m:
                period = m.group(1)

        if ("Journal" in line or "JOURNAL" in line) and not journal_code:
            # ex: "Journal 001 Ventes et prestations"
            m = re.search(r"\bJournal\b\s*([0-9]{1,3})\s+(.+)$", line, flags=re.IGNORECASE)
            if m:
                journal_code = m.group(1).zfill(3)
                journal_lib = normalize_space(m.group(2))

        # fallback: "001 Ventes et prestations"
        if not journal_code:
            m = re.search(r"\b([0-9]{3})\b\s+([A-Za-zÀ-ÿ].+)", line)
            if m and "Folio" not in line and "Période" not in line and "/" not in m.group(2):
                journal_code = m.group(1)
                journal_lib = normalize_space(m.group(2))

    return journal_code, journal_lib, period


def find_header_row(df_raw: pd.DataFrame) -> Optional[int]:
    """
    Find header row containing most of:
      Ecr / Jour / Pièce / Compte / Débit / Crédit
    """
    required = ["ecr", "jour", "pi", "comp", "dé", "cr"]
    max_scan = min(len(df_raw), 250)
    for r in range(max_scan):
        row = [safe_str_cell(c).lower() for c in df_raw.iloc[r].tolist()]
        joined = " ".join(row)
        score = sum(1 for k in required if k in joined)
        if score >= 5:
            return r
    return None


def pick_col(columns: List[str], patterns: List[str], fallback_contains: Optional[str] = None) -> Optional[str]:
    for p in patterns:
        for c in columns:
            if re.search(p, c, flags=re.IGNORECASE):
                return c
    if fallback_contains:
        for c in columns:
            if fallback_contains.lower() in c.lower():
                return c
    return None


def parse_period_to_year_month(period: str) -> Tuple[int, int]:
    """
    period format: '12/2025'
    """
    m = re.match(r"^\s*([0-1]?\d)\s*/\s*(20\d{2})\s*$", period or "")
    if not m:
        raise ValueError(f"Période introuvable ou au mauvais format (attendu 'MM/YYYY') : '{period}'")
    mois = int(m.group(1))
    annee = int(m.group(2))
    if not (1 <= mois <= 12):
        raise ValueError(f"Mois invalide dans la période : {period}")
    return annee, mois


def coerce_day(jour_cell: str) -> int:
    s = safe_str_cell(jour_cell).strip()
    s = re.sub(r"\D", "", s)
    if s == "":
        return 1
    try:
        d = int(s)
        if 1 <= d <= 31:
            return d
    except Exception:
        pass
    return 1


def detect_sheets(xls: pd.ExcelFile) -> List[str]:
    # Keep all sheets, user can choose to parse all or select some
    return list(xls.sheet_names)


def parse_one_sheet_to_fec(
    df_raw: pd.DataFrame,
    file_name: str,
    sheet_name: str,
    default_journal_code: str = "001",
    default_journal_lib: str = "Journal",
) -> Tuple[pd.DataFrame, Dict[str, str]]:
    """
    Parse one raw sheet (header=None) into a FEC dataframe.
    """
    journal_code, journal_lib, period = detect_period_journal(df_raw)
    journal_code = journal_code or default_journal_code
    journal_lib = journal_lib or default_journal_lib

    header_idx = find_header_row(df_raw)
    if header_idx is None:
        raise ValueError(f"[{file_name} / {sheet_name}] Impossible de trouver l'entête du tableau (Ecr/Jour/Pièce/Compte/Débit/Crédit).")

    headers = [normalize_space(safe_str_cell(c)) for c in df_raw.iloc[header_idx].tolist()]
    data = df_raw.iloc[header_idx + 1:].copy()
    data.columns = headers
    data = data.dropna(how="all")

    cols = list(data.columns)

    # core columns
    col_ecr = pick_col(cols, [r"^Ecr\.?$", r"\bEcr\b"], fallback_contains="Ecr")
    col_jour = pick_col(cols, [r"\bJour\b"])
    col_piece = pick_col(cols, [r"Pi[eè]ce"])
    col_compte = pick_col(cols, [r"\bCompte\b"])
    col_debit = pick_col(cols, [r"D[ée]bit"])
    col_credit = pick_col(cols, [r"Cr[ée]dit"])

    if not all([col_ecr, col_jour, col_piece, col_compte, col_debit, col_credit]):
        missing = [n for n, v in [
            ("Ecr", col_ecr), ("Jour", col_jour), ("Pièce", col_piece),
            ("Compte", col_compte), ("Débit", col_debit), ("Crédit", col_credit)
        ] if not v]
        raise ValueError(f"[{file_name} / {sheet_name}] Colonnes manquantes: {', '.join(missing)}")

    # Libellé écriture
    col_lib_ecr = pick_col(cols, [r"Libell[ée]\s*[ée]criture", r"\bLibell[ée]\b"], fallback_contains="Libell")

    # CompteLib: either explicit column, or guess next to Compte if it looks textual
    col_compte_lib = None
    compte_idx = cols.index(col_compte)
    if compte_idx + 1 < len(cols):
        candidate = cols[compte_idx + 1]
        if candidate not in (col_debit, col_credit) and (
            re.search(r"libell", candidate, flags=re.IGNORECASE)
            or data[candidate].head(30).astype(str).str.contains(r"[A-Za-zÀ-ÿ]", regex=True).mean() > 0.6
        ):
            col_compte_lib = candidate

    # Period -> year/month
    annee, mois = parse_period_to_year_month(period)

    # Filter rows that look like movements
    data[col_piece] = data[col_piece].astype(str).map(lambda x: normalize_space(safe_str_cell(x)))
    data[col_compte] = data[col_compte].astype(str).map(lambda x: normalize_space(safe_str_cell(x)))

    data = data[(data[col_piece] != "") & (data[col_compte] != "")]

    fec_rows = []
    for _, r in data.iterrows():
        piece = normalize_space(safe_str_cell(r.get(col_piece)))
        ecr_num = piece or normalize_space(safe_str_cell(r.get(col_ecr)))
        jour = r.get(col_jour)
        d = coerce_day(jour)
        ecr_date = date(annee, mois, d).strftime("%Y%m%d")

        compte_num = extract_compte_num(r.get(col_compte))
        compte_lib = normalize_space(safe_str_cell(r.get(col_compte_lib))) if col_compte_lib else ""

        lib_ecr = normalize_space(safe_str_cell(r.get(col_lib_ecr))) if col_lib_ecr else ""

        debit = to_decimal_fr(r.get(col_debit))
        credit = to_decimal_fr(r.get(col_credit))

        fec_rows.append({
            "JournalCode": journal_code,
            "JournalLib": journal_lib,
            "EcritureNum": ecr_num,
            "EcritureDate": ecr_date,
            "CompteNum": compte_num,
            "CompteLib": compte_lib,
            "CompAuxNum": "",
            "CompAuxLib": "",
            "PieceRef": piece,
            "PieceDate": ecr_date,
            "EcritureLib": lib_ecr,
            "Debit": round(debit, 2),
            "Credit": round(credit, 2),
            "EcritureLet": "",
            "DateLet": "",
            "ValidDate": "",
            "Montantdevise": "",
            "Idevise": "",
            # internal trace (not exported)
            "_src_file": file_name,
            "_src_sheet": sheet_name,
        })

    fec = pd.DataFrame(fec_rows)

    # Ensure standard columns exist
    for c in FEC_COLUMNS:
        if c not in fec.columns:
            fec[c] = ""

    # Format debit/credit as 2 decimals (FEC expects numeric; we keep as string with dot)
    fec["Debit"] = fec["Debit"].map(lambda x: f"{float(x):.2f}" if str(x) != "" else "")
    fec["Credit"] = fec["Credit"].map(lambda x: f"{float(x):.2f}" if str(x) != "" else "")

    meta = {
        "JournalCode": journal_code,
        "JournalLib": journal_lib,
        "Period": period,
        "File": file_name,
        "Sheet": sheet_name,
        "Rows": str(len(fec)),
    }
    return fec[FEC_COLUMNS + ["_src_file", "_src_sheet"]], meta


def fec_to_text(df_fec: pd.DataFrame) -> str:
    out = io.StringIO()
    out.write("|".join(FEC_COLUMNS) + "\n")
    for _, r in df_fec.iterrows():
        vals = []
        for c in FEC_COLUMNS:
            v = r.get(c, "")
            if pd.isna(v):
                v = ""
            vals.append(str(v))
        out.write("|".join(vals) + "\n")
    return out.getvalue()


def add_balance_controls(df_fec: pd.DataFrame) -> Tuple[pd.DataFrame, Dict[str, float]]:
    """
    Returns per-piece balance and global totals.
    """
    df = df_fec.copy()

    # numeric columns
    df["_debit_num"] = df["Debit"].astype(str).str.replace(",", ".", regex=False).replace("", "0").astype(float)
    df["_credit_num"] = df["Credit"].astype(str).str.replace(",", ".", regex=False).replace("", "0").astype(float)

    by_piece = df.groupby(
        ["JournalCode", "PieceRef", "EcritureDate"],
        dropna=False
    ).agg(
        Debit=("_debit_num", "sum"),
        Credit=("_credit_num", "sum"),
        Lignes=("PieceRef", "size"),
    ).reset_index()

    by_piece["Ecart"] = (by_piece["Debit"] - by_piece["Credit"]).round(2)

    totals = {
        "TotalDebit": float(df["_debit_num"].sum()),
        "TotalCredit": float(df["_credit_num"].sum()),
    }
    totals["TotalEcart"] = round(totals["TotalDebit"] - totals["TotalCredit"], 2)

    # show biggest imbalances first
    by_piece = by_piece.sort_values("Ecart", key=lambda s: s.abs(), ascending=False)

    return by_piece, totals


# -----------------------------
# Streamlit UI
# -----------------------------
st.set_page_config(page_title="Multi-fichiers -> FEC", layout="wide")
st.title("Convertisseur multi-fichiers d’éditions (XLS/XLSX) vers un FEC unique")

st.write(
    "Charge **plusieurs fichiers** d’éditions de journaux (périodes et journaux différents possibles). "
    "L’app produit **un seul fichier FEC** (texte séparé par `|`) et affiche un **contrôle Débit/Crédit**."
)

uploaded_files = st.file_uploader(
    "Fichiers XLS/XLSX (plusieurs)",
    type=["xls", "xlsx"],
    accept_multiple_files=True
)

colA, colB, colC = st.columns([1, 1, 2])
with colA:
    parse_all_sheets = st.checkbox("Parser toutes les feuilles", value=True)
with colB:
    strict_balance = st.checkbox("Bloquer téléchargement si déséquilibre global", value=False)
with colC:
    st.caption(
        "Si tu décoches 'Parser toutes les feuilles', l’app prendra uniquement la 1ère feuille de chaque fichier. "
        "Pour éviter les surprises, l’option 'bloquer' peut imposer Débit = Crédit global."
    )

if uploaded_files:
    fec_parts: List[pd.DataFrame] = []
    metas: List[Dict[str, str]] = []
    errors: List[str] = []

    with st.spinner("Lecture et conversion des fichiers…"):
        for uf in uploaded_files:
            file_name = uf.name
            engine = excel_engine_for_filename(file_name)

            try:
                xls = pd.ExcelFile(uf, engine=engine)
            except Exception as e:
                errors.append(f"[{file_name}] Impossible d’ouvrir le fichier Excel ({engine}) : {e}")
                continue

            sheets = detect_sheets(xls)
            if not sheets:
                errors.append(f"[{file_name}] Aucune feuille détectée.")
                continue

            sheets_to_parse = sheets if parse_all_sheets else [sheets[0]]

            for sh in sheets_to_parse:
                try:
                    df_raw = pd.read_excel(xls, sheet_name=sh, header=None)  # uses xls engine
                    fec_df, meta = parse_one_sheet_to_fec(
                        df_raw=df_raw,
                        file_name=file_name,
                        sheet_name=sh,
                        default_journal_code="001",
                        default_journal_lib="Journal",
                    )
                    if len(fec_df) > 0:
                        fec_parts.append(fec_df)
                        metas.append(meta)
                except Exception as e:
                    # Keep parsing other sheets/files
                    errors.append(str(e))

    st.subheader("Résultat de conversion")
    if metas:
        st.dataframe(pd.DataFrame(metas), use_container_width=True)
    else:
        st.warning("Aucune donnée convertie. Vérifie le format des fichiers (entête du tableau, période, etc.).")

    if errors:
        with st.expander(f"Erreurs / feuilles ignorées ({len(errors)})"):
            for err in errors:
                st.error(err)

    if fec_parts:
        fec_all = pd.concat(fec_parts, ignore_index=True)

        # Optional: remove empty lines
        fec_all = fec_all[(fec_all["PieceRef"].astype(str).str.strip() != "") & (fec_all["CompteNum"].astype(str).str.strip() != "")]

        # Controls
        by_piece, totals = add_balance_controls(fec_all)

        st.subheader("Contrôle Débit / Crédit (global)")
        st.write(
            f"Total Débit = **{totals['TotalDebit']:.2f}** ; "
            f"Total Crédit = **{totals['TotalCredit']:.2f}** ; "
            f"Écart = **{totals['TotalEcart']:.2f}**"
        )

        st.subheader("Contrôle Débit / Crédit par pièce (top écarts)")
        st.dataframe(by_piece.head(50), use_container_width=True)

        st.subheader("Aperçu du FEC (50 premières lignes)")
        st.dataframe(fec_all[FEC_COLUMNS].head(50), use_container_width=True)

        fec_text = fec_to_text(fec_all[FEC_COLUMNS])

        # filename: single output
        out_name = "FEC_multi_journaux.txt"

        can_download = True
        if strict_balance and abs(totals["TotalEcart"]) > 0.0001:
            can_download = False
            st.error("Téléchargement bloqué : le FEC global n’est pas équilibré (Débit ≠ Crédit).")

        if can_download:
            st.download_button(
                "Télécharger le FEC unique",
                data=fec_text.encode("utf-8"),
                file_name=out_name,
                mime="text/plain"
            )

        with st.expander("Voir le début du fichier FEC (texte)"):
            st.code("\n".join(fec_text.splitlines()[:40]), language="text")

else:
    st.info("Charge un ou plusieurs fichiers Excel pour générer un FEC unique.")
