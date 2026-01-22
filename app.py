from __future__ import annotations

import os
import re
from pathlib import Path
from typing import Optional, Tuple, List

import pandas as pd


HEADER_KEYWORDS = [
    "date", "journal", "pièce", "piece", "n°", "numero", "numéro",
    "compte", "libellé", "libelle", "débit", "debit", "crédit", "credit",
    "contrepartie", "tiers", "référence", "reference", "écriture", "ecriture"
]

# Certains exports ont des variantes
CANON_COLS = {
    "date": ["date", "date piece", "date pièce"],
    "journal": ["journal", "code journal"],
    "piece": ["pièce", "piece", "n° pièce", "no piece", "num piece", "numéro pièce", "numero piece"],
    "compte": ["compte", "n° compte", "numero compte", "no compte"],
    "tiers": ["tiers", "compte tiers", "auxiliaire", "client", "fournisseur"],
    "libelle": ["libellé", "libelle", "intitulé", "intitule", "description"],
    "debit": ["débit", "debit"],
    "credit": ["crédit", "credit"],
}


def _norm(s: str) -> str:
    s = str(s).strip().lower()
    # simplifie accents et ponctuation basique
    s = s.replace("é", "e").replace("è", "e").replace("ê", "e") \
         .replace("à", "a").replace("ù", "u").replace("ç", "c")
    s = re.sub(r"\s+", " ", s)
    return s


def _score_header_row(row: pd.Series) -> int:
    vals = [_norm(v) for v in row.tolist()]
    hits = 0
    for v in vals:
        if not v or v == "nan":
            continue
        if any(k in v for k in HEADER_KEYWORDS):
            hits += 1
    return hits


def _find_header_row(df_raw: pd.DataFrame, max_scan_rows: int = 80) -> Optional[int]:
    # On scanne les premières lignes pour trouver celle qui ressemble à un header
    scan = df_raw.head(max_scan_rows)
    best_i, best_score = None, 0
    for i in range(len(scan)):
        score = _score_header_row(scan.iloc[i])
        if score > best_score:
            best_score = score
            best_i = i
    # seuil : au moins 2 mots-clés trouvés
    if best_score >= 2:
        return best_i
    return None


def _cleanup_df(df: pd.DataFrame) -> pd.DataFrame:
    # drop colonnes totalement vides
    df = df.dropna(axis=1, how="all")
    # drop lignes totalement vides
    df = df.dropna(axis=0, how="all")

    # trim colonnes
    df.columns = [str(c).strip() for c in df.columns]
    return df


def _canonize_columns(df: pd.DataFrame) -> pd.DataFrame:
    # Normalisation des noms
    colmap = {}
    norm_cols = {_norm(c): c for c in df.columns}
    for canon, variants in CANON_COLS.items():
        for v in variants:
            nv = _norm(v)
            # match exact
            if nv in norm_cols:
                colmap[norm_cols[nv]] = canon
                break
            # match contains (au cas où: "Montant Débit", etc.)
            for nc, orig in norm_cols.items():
                if nv in nc:
                    colmap[orig] = canon
                    break
            if canon in colmap.values():
                break

    if colmap:
        df = df.rename(columns=colmap)

    return df


def _try_read_excel_sheets(path: str) -> List[Tuple[str, pd.DataFrame]]:
    # Lit toutes les feuilles sans supposer le header
    xls = pd.ExcelFile(path)
    out = []
    for sh in xls.sheet_names:
        df = pd.read_excel(path, sheet_name=sh, header=None, dtype=str)
        out.append((sh, df))
    return out


def _try_read_html_tables(path: str) -> List[pd.DataFrame]:
    # Si .xls déguisé en HTML, read_html marche souvent parfaitement
    tables = pd.read_html(path, header=None)
    # read_html renvoie des DataFrames
    return tables


def read_journal_file(path: str) -> pd.DataFrame:
    """
    Charge un fichier de journaux (.xls/.xlsx ou .xls HTML déguisé),
    détecte la feuille + la ligne d'entête et retourne un DF propre.
    """
    path = str(path)
    candidates: List[Tuple[str, pd.DataFrame]] = []

    # 1) Tentative Excel classique (toutes feuilles)
    try:
        sheets = _try_read_excel_sheets(path)
        candidates.extend([(f"sheet:{name}", df) for name, df in sheets])
    except Exception:
        # 2) Tentative HTML
        try:
            tables = _try_read_html_tables(path)
            candidates.extend([(f"html_table:{i}", df) for i, df in enumerate(tables)])
        except Exception as e:
            raise RuntimeError(f"Impossible de lire le fichier: {path}") from e

    # Pour chaque candidate, on cherche un header puis on reconstruit le DF
    best_df = None
    best_quality = -1

    for origin, df_raw in candidates:
        df_raw = df_raw.copy()
        df_raw = _cleanup_df(df_raw)

        if df_raw.empty or df_raw.shape[1] < 2:
            continue

        header_i = _find_header_row(df_raw)
        if header_i is None:
            continue

        header = df_raw.iloc[header_i].tolist()
        df = df_raw.iloc[header_i + 1 :].copy()
        df.columns = [str(h).strip() for h in header]
        df = _cleanup_df(df)

        # évalue la “qualité” : nb de colonnes utiles détectées
        df2 = _canonize_columns(df)
        quality = sum(1 for c in ["date", "compte", "libelle", "debit", "credit"] if c in df2.columns)

        # bonus si on a un volume de lignes correct
        if len(df2) > 10:
            quality += 1

        if quality > best_quality:
            best_quality = quality
            best_df = df2

    if best_df is None:
        raise ValueError(
            f"Aucune table exploitable détectée dans {path}. "
            "Le fichier est probablement une édition très mise en page (impression) "
            "ou sans entête identifiable."
        )

    return best_df


def import_folder(folder: str, pattern: Tuple[str, ...] = (".xls", ".xlsx")) -> pd.DataFrame:
    folder = str(folder)
    files = sorted(
        str(p) for p in Path(folder).rglob("*")
        if p.is_file() and p.suffix.lower() in pattern
    )
    dfs = []
    errors = []

    for f in files:
        try:
            df = read_journal_file(f)
            df["__source_file"] = os.path.basename(f)
            dfs.append(df)
        except Exception as e:
            errors.append((f, str(e)))

    out = pd.concat(dfs, ignore_index=True) if dfs else pd.DataFrame()

    if errors:
        # tu peux logger/print si tu veux
        err_df = pd.DataFrame(errors, columns=["file", "error"])
        print("Fichiers en erreur (à vérifier) :")
        print(err_df.to_string(index=False))

    return out
