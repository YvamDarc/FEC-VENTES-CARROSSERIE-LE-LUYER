# app.py
import re
import io
from datetime import date
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

def to_decimal_fr(x):
    """Convertit '1 002,00' / '556,80' / 556.80 / NaN -> float (ou 0)."""
    if x is None or (isinstance(x, float) and pd.isna(x)):
        return 0.0
    if isinstance(x, (int, float)):
        return float(x)
    s = str(x).strip()
    if s == "" or s.lower() == "nan":
        return 0.0
    # retire espaces milliers
    s = s.replace("\u00A0", " ").replace(" ", "")
    # virgule décimale -> point
    s = s.replace(",", ".")
    # garde uniquement chiffres/point/signe
    s = re.sub(r"[^0-9\.\-]", "", s)
    if s in ("", "-", ".", "-."):
        return 0.0
    try:
        return float(s)
    except ValueError:
        return 0.0

def find_period_and_journal(df_raw: pd.DataFrame):
    """
    Tente de détecter :
      - JournalCode (ex: 001)
      - JournalLib (ex: Ventes et prestations)
      - Période (mois/année) ex: 12/2025
    en scannant les cellules texte.
    """
    journal_code = ""
    journal_lib = ""
    period = ""

    # concat lignes en texte
    for r in range(min(len(df_raw), 80)):
        row = df_raw.iloc[r].astype(str).fillna("").tolist()
        line = " ".join([c for c in row if c and c.lower() != "nan"])
        line = re.sub(r"\s+", " ", line).strip()

        # Période
        if not period:
            m = re.search(r"\bP[ée]riode\b\s*([0-1]?\d\/20\d{2})", line, flags=re.IGNORECASE)
            if m:
                period = m.group(1)

        # Journal
        if ("Journal" in line or "JOURNAL" in line) and not journal_code:
            # Ex: "Journal 001 Ventes et prestations"
            m = re.search(r"\bJournal\b\s*([0-9]{1,3})\s+(.+)$", line, flags=re.IGNORECASE)
            if m:
                journal_code = m.group(1).zfill(3)
                journal_lib = m.group(2).strip()
        # Autre forme : "001 Ventes et prestations"
        if not journal_code:
            m = re.search(r"\b([0-9]{3})\b\s+([A-Za-zÀ-ÿ].+)", line)
            if m and "Folio" not in line:
                journal_code = m.group(1)
                # évite de prendre des choses type "12/2025"
                if "/" not in m.group(2):
                    journal_lib = m.group(2).strip()

    return journal_code, journal_lib, period

def find_header_row(df_raw: pd.DataFrame):
    """
    Cherche la ligne d'entête contenant au moins Ecr / Jour / Pièce / Compte / Débit / Crédit.
    Retourne l'index de ligne.
    """
    required = ["ecr", "jour", "pi", "comp", "dé", "cr"]
    for r in range(min(len(df_raw), 200)):
        row = df_raw.iloc[r].astype(str).fillna("").str.lower().tolist()
        joined = " ".join(row)
        score = sum(1 for k in required if k in joined)
        if score >= 5:
            return r
    return None

def normalize_headers(header_row):
    """
    Normalise les noms de colonnes.
    """
    headers = []
    for h in header_row:
        s = str(h).strip()
        s = re.sub(r"\s+", " ", s)
        headers.append(s)
    return headers

def parse_table(df_raw: pd.DataFrame):
    journal_code, journal_lib, period = find_period_and_journal(df_raw)
    header_idx = find_header_row(df_raw)
    if header_idx is None:
        raise ValueError("Impossible de trouver la ligne d'entête (Ecr/Jour/Pièce/Compte/Débit/Crédit).")

    headers = normalize_headers(df_raw.iloc[header_idx].tolist())
    data = df_raw.iloc[header_idx + 1:].copy()
    data.columns = headers

    # supprime lignes vides
    data = data.dropna(how="all")
    # enlève les lignes où aucune info "Pièce" ou "Compte" n'apparait
    # (la colonne peut s'appeler "Pièce" ou "Piece" suivant export)
    possible_piece_cols = [c for c in data.columns if re.search(r"pi[eè]ce", c, flags=re.IGNORECASE)]
    possible_compte_cols = [c for c in data.columns if re.search(r"compte", c, flags=re.IGNORECASE)]
    if not possible_piece_cols or not possible_compte_cols:
        raise ValueError("Colonnes 'Pièce' ou 'Compte' introuvables après lecture.")

    piece_col = possible_piece_cols[0]
    compte_col = possible_compte_cols[0]

    # détecte colonnes Débit / Crédit
    debit_cols = [c for c in data.columns if re.search(r"d[ée]bit", c, flags=re.IGNORECASE)]
    credit_cols = [c for c in data.columns if re.search(r"cr[ée]dit", c, flags=re.IGNORECASE)]
    if not debit_cols or not credit_cols:
        raise ValueError("Colonnes 'Débit' et/ou 'Crédit' introuvables.")

    debit_col = debit_cols[0]
    credit_col = credit_cols[0]

    # Colonnes Ecr / Jour
    ecr_cols = [c for c in data.columns if re.fullmatch(r"Ecr\.?|Ecr", c, flags=re.IGNORECASE)]
    if not ecr_cols:
        # fallback : contient "Ecr"
        ecr_cols = [c for c in data.columns if "ecr" in c.lower()]
    if not ecr_cols:
        raise ValueError("Colonne 'Ecr.' introuvable.")
    ecr_col = ecr_cols[0]

    jour_cols = [c for c in data.columns if re.search(r"\bjour\b", c, flags=re.IGNORECASE)]
    if not jour_cols:
        raise ValueError("Colonne 'Jour' introuvable.")
    jour_col = jour_cols[0]

    # Certaines éditions ont 2 colonnes de libellés :
    # - une pour le libellé du compte (VENTE PIECE, MO, TVA...)
    # - une pour le libellé d'écriture (Fact. xxxx - Client ...)
    # On cherche une colonne "Libellé écriture" / "Libellé"
    lib_ecr_cols = [c for c in data.columns if re.search(r"libell[ée]\s*[ée]criture", c, flags=re.IGNORECASE)]
    lib_cols = [c for c in data.columns if re.fullmatch(r"Libell[ée]\s*", c, flags=re.IGNORECASE) or "libell" in c.lower()]

    lib_ecr_col = lib_ecr_cols[0] if lib_ecr_cols else (lib_cols[-1] if lib_cols else None)

    # Le libellé du compte est parfois dans une colonne entre Compte et Libellé écriture
    # (comme ton exemple : "Compte" puis "Libellé écriture" mais en fait tu as les 2)
    # Ici on tente de prendre la colonne juste après "Compte" si elle ressemble à un libellé de compte.
    compte_idx = list(data.columns).index(compte_col)
    compte_lib_col = None
    if compte_idx + 1 < len(data.columns):
        candidate = data.columns[compte_idx + 1]
        # si la candidate n'est pas Débit/Crédit et contient souvent des mots (pas des montants)
        if candidate not in (debit_col, credit_col) and "libell" in candidate.lower():
            compte_lib_col = candidate
        else:
            # parfois la colonne n'est pas nommée "Libellé" mais contient VENTE PIECE/MO/TVA...
            # on teste sur quelques lignes
            sample = data[candidate].head(20).astype(str)
            if candidate not in (debit_col, credit_col) and sample.str.contains(r"[A-Za-zÀ-ÿ]", regex=True).mean() > 0.6:
                compte_lib_col = candidate

    # période -> (mois, année)
    if period:
        m = re.match(r"^\s*([0-1]?\d)\s*/\s*(20\d{2})\s*$", period)
        if not m:
            raise ValueError(f"Période détectée mais format inattendu: {period}")
        mois = int(m.group(1))
        annee = int(m.group(2))
    else:
        # fallback : on essaie de trouver un "12/2025" dans la page
        mois, annee = 1, 2000

    # Nettoyage et extraction des comptes : parfois "C 30020" ou colonne séparée
    def extract_compte_num(v):
        if v is None or (isinstance(v, float) and pd.isna(v)):
            return ""
        s = str(v).strip()
        s = re.sub(r"\s+", " ", s)
        # enlève un éventuel indicateur "C"
        s = re.sub(r"^\s*C\s+", "", s, flags=re.IGNORECASE)
        # garde les chiffres (compte)
        m = re.search(r"([0-9]{3,})", s)
        return m.group(1) if m else s

    # Stopper à la première grande zone vide de mouvements
    # (si l'édition contient des pages/folios)
    # On garde les lignes où piece et compte existent.
    data[piece_col] = data[piece_col].astype(str).str.strip()
    data[compte_col] = data[compte_col].astype(str).str.strip()
    data = data[(data[piece_col] != "") & (data[piece_col].str.lower() != "nan") & (data[compte_col] != "")]

    rows = []
    for _, r in data.iterrows():
        piece = str(r.get(piece_col, "")).strip()
        ecr = str(r.get(ecr_col, "")).strip()
        jour = str(r.get(jour_col, "")).strip()

        compte_num = extract_compte_num(r.get(compte_col, ""))
        compte_lib = str(r.get(compte_lib_col, "")).strip() if compte_lib_col else ""

        lib_ecr = str(r.get(lib_ecr_col, "")).strip() if lib_ecr_col else ""
        debit = to_decimal_fr(r.get(debit_col, 0))
        credit = to_decimal_fr(r.get(credit_col, 0))

        # Date d'écriture : année/mois de période + jour
        try:
            d = int(re.sub(r"\D", "", jour)) if jour else 1
            ecr_date = date(annee, mois, d).strftime("%Y%m%d")
        except Exception:
            ecr_date = date(annee, mois, 1).strftime("%Y%m%d")

        fec_row = {
            "JournalCode": journal_code or "001",
            "JournalLib": journal_lib or "Ventes",
            "EcritureNum": piece if piece and piece.lower() != "nan" else (ecr or ""),
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
            "Idevise": ""
        }
        rows.append(fec_row)

    fec = pd.DataFrame(rows, columns=FEC_COLUMNS)

    # sécurise formats
    fec["Debit"] = fec["Debit"].map(lambda x: f"{x:.2f}" if x != "" else "")
    fec["Credit"] = fec["Credit"].map(lambda x: f"{x:.2f}" if x != "" else "")

    return fec, {"journal_code": journal_code, "journal_lib": journal_lib, "period": period}

def fec_to_text(df_fec: pd.DataFrame) -> str:
    # FEC = séparateur |
    # Valeurs vides -> vide
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

def check_balance(df_fec: pd.DataFrame):
    # Totaux par pièce + global
    df = df_fec.copy()
    df["Debit_num"] = df["Debit"].astype(str).str.replace(",", ".", regex=False).astype(float)
    df["Credit_num"] = df["Credit"].astype(str).str.replace(",", ".", regex=False).astype(float)

    by_piece = df.groupby("PieceRef", dropna=False).agg(
        Debit=("Debit_num", "sum"),
        Credit=("Credit_num", "sum"),
        Lignes=("PieceRef", "size")
    ).reset_index()

    by_piece["Ecart"] = (by_piece["Debit"] - by_piece["Credit"]).round(2)

    total_debit = float(df["Debit_num"].sum())
    total_credit = float(df["Credit_num"].sum())
    total_ecart = round(total_debit - total_credit, 2)

    return by_piece, total_debit, total_credit, total_ecart


st.set_page_config(page_title="Convertisseur Éditions de ventes -> FEC", layout="wide")
st.title("Convertisseur d’éditions de ventes (XLS/XLSX) vers FEC")

st.write(
    "Charge une édition de ventes (type *Journaux comptables – Ventes et prestations*) "
    "et télécharge le fichier FEC (texte séparé par `|`)."
)

uploaded = st.file_uploader("Fichier XLS/XLSX", type=["xls", "xlsx"])

if uploaded:
    # Lecture brute (sans header), toutes cellules
    xls = pd.ExcelFile(uploaded)
    sheet = st.selectbox("Feuille à importer", xls.sheet_names, index=0)

    df_raw = pd.read_excel(xls, sheet_name=sheet, header=None, engine="openpyxl")

    try:
        fec_df, meta = parse_table(df_raw)

        st.subheader("Détection")
        c1, c2, c3 = st.columns(3)
        c1.metric("Journal", meta.get("journal_code") or "—")
        c2.metric("Libellé", meta.get("journal_lib") or "—")
        c3.metric("Période", meta.get("period") or "—")

        st.subheader("Aperçu des écritures FEC")
        st.dataframe(fec_df.head(50), use_container_width=True)

        st.subheader("Contrôle d’équilibre")
        by_piece, td, tc, te = check_balance(fec_df)
        st.write(f"Total Débit = **{td:.2f}** ; Total Crédit = **{tc:.2f}** ; Écart = **{te:.2f}**")
        st.dataframe(by_piece.sort_values("Ecart", key=lambda s: s.abs(), ascending=False).head(30),
                     use_container_width=True)

        fec_text = fec_to_text(fec_df)
        default_name = f"FEC_{meta.get('journal_code') or 'JRN'}_{(meta.get('period') or 'PERIODE').replace('/','')}.txt"

        st.download_button(
            "Télécharger le FEC",
            data=fec_text.encode("utf-8"),
            file_name=default_name,
            mime="text/plain"
        )

        with st.expander("Voir le contenu texte (début)"):
            st.code("\n".join(fec_text.splitlines()[:30]), language="text")

    except Exception as e:
        st.error(str(e))
        st.info(
            "Astuce : si ton XLS a une mise en page très 'imprimante' (cellules fusionnées, colonnes décalées), "
            "essaie d’exporter une version plus 'table' (ou de fournir une export balance/journaux)."
        )

else:
    st.caption("➡️ Charge un fichier pour démarrer.")
