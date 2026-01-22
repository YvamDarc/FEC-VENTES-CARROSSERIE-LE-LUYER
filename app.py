# app.py
from __future__ import annotations

import os
import re
import io
import zipfile
import tempfile
from pathlib import Path
from typing import Optional, Tuple, List

import streamlit as st
import pandas as pd


# -----------------------------
# UI / sécurité anti-page-blanche
# -----------------------------
st.set_page_config(page_title="Import journaux", layout="wide")
st.title("Import journaux comptables — multi-fichiers")

st.success("✅ L’app démarre et Streamlit rend la page.")

# Astuce: tout le reste est encapsulé pour afficher les erreurs dans l'UI
try:
    # -----------------------------
    # Helpers
    # -----------------------------
    def _clean_colname(x: str) -> str:
        x = "" if x is None else str(x)
        x = x.strip()
        x = re.sub(r"\s+", " ", x)
        return x

    def _safe_read_csv(path: str | Path) -> pd.DataFrame:
        # Essaye quelques encodages usuels FR
        for enc in ("utf-8", "utf-8-sig", "cp1252", "latin1"):
            try:
                return pd.read_csv(path, dtype=str, encoding=enc, sep=None, engine="python")
            except Exception:
                continue
        # dernier recours
        return pd.read_csv(path, dtype=str, encoding_errors="replace", sep=None, engine="python")

    def _safe_read_excel(path: str | Path) -> pd.DataFrame:
        """
        Stratégie robuste :
        - read_excel classique
        - si échec sur .xls : tenter read_html (souvent un export HTML renommé .xls)
        """
        p = str(path).lower()
        # 1) Excel natif
        try:
            df = pd.read_excel(path, dtype=str)
            return df
        except Exception as e_excel:
            # 2) Fallback HTML pour certains .xls
            if p.endswith(".xls"):
                try:
                    tables = pd.read_html(path)  # type: ignore
                    if tables and len(tables) > 0:
                        return tables[0].astype(str)
                except Exception:
                    pass
            raise e_excel

    def read_any(path: str | Path) -> Tuple[pd.DataFrame, str]:
        """
        Retourne (df, mode_lecture).
        """
        p = str(path).lower()
        if p.endswith(".csv") or p.endswith(".txt"):
            df = _safe_read_csv(path)
            return df, "csv"
        if p.endswith(".xlsx") or p.endswith(".xlsm") or p.endswith(".xls"):
            df = _safe_read_excel(path)
            return df, "excel/html"
        raise ValueError(f"Format non supporté: {path}")

    def list_files_from_uploaded_zip(zip_bytes: bytes) -> List[Tuple[str, bytes]]:
        """
        Retourne une liste de (filename, content_bytes) depuis un zip uploadé.
        """
        out = []
        with zipfile.ZipFile(io.BytesIO(zip_bytes)) as z:
            for name in z.namelist():
                if name.endswith("/") or name.startswith("__MACOSX/"):
                    continue
                out.append((name, z.read(name)))
        return out

    def save_bytes_to_tempfile(filename: str, content: bytes) -> str:
        suffix = "." + filename.split(".")[-1].lower()
        with tempfile.NamedTemporaryFile(delete=False, suffix=suffix) as tmp:
            tmp.write(content)
            return tmp.name

    def normalize_df(df: pd.DataFrame) -> pd.DataFrame:
        """
        Normalisation légère (à adapter à tes journaux) :
        - forcer string
        - nettoyer noms de colonnes
        - retirer colonnes vides
        """
        # colonnes
        df = df.copy()
        df.columns = [_clean_colname(c) for c in df.columns]

        # valeurs en str
        for c in df.columns:
            df[c] = df[c].astype(str)

        # retire colonnes entièrement vides (ou "nan")
        def is_empty_series(s: pd.Series) -> bool:
            x = s.fillna("").astype(str).str.strip()
            x = x.replace({"nan": "", "None": ""})
            return (x == "").all()

        df = df.loc[:, [c for c in df.columns if not is_empty_series(df[c])]]
        return df

    def concat_with_origin(dfs: List[Tuple[pd.DataFrame, str, str]]) -> pd.DataFrame:
        """
        dfs: [(df, source_name, read_mode), ...]
        """
        out = []
        for df, src, mode in dfs:
            d = df.copy()
            d.insert(0, "_source_file", src)
            d.insert(1, "_read_mode", mode)
            out.append(d)
        if not out:
            return pd.DataFrame()
        return pd.concat(out, ignore_index=True)

    # -----------------------------
    # Sidebar : modes d'import
    # -----------------------------
    st.sidebar.header("Import")

    import_mode = st.sidebar.radio(
        "Tu importes comment ?",
        ["Uploader des fichiers", "Uploader un ZIP", "Scanner un dossier (local)"],
        index=0,
    )

    st.sidebar.caption("Formats acceptés : .xls, .xlsx, .csv, .txt")

    # Option : lecture rapide vs complète (si fichiers lourds)
    nrows_opt = st.sidebar.number_input(
        "Limiter le nombre de lignes lues (0 = tout)",
        min_value=0, max_value=2_000_000, value=0, step=1000
    )

    # -----------------------------
    # Collecte des fichiers
    # -----------------------------
    selected_paths: List[Tuple[str, str]] = []  # (display_name, filepath)

    if import_mode == "Uploader des fichiers":
        uploaded_files = st.file_uploader(
            "Dépose tes fichiers (multi) :",
            type=["xls", "xlsx", "csv", "txt"],
            accept_multiple_files=True
        )

        if uploaded_files:
            for uf in uploaded_files:
                tmp_path = save_bytes_to_tempfile(uf.name, uf.getbuffer().tobytes())
                selected_paths.append((uf.name, tmp_path))

    elif import_mode == "Uploader un ZIP":
        uploaded_zip = st.file_uploader("Dépose un ZIP contenant tes fichiers :", type=["zip"])
        if uploaded_zip is not None:
            files_in_zip = list_files_from_uploaded_zip(uploaded_zip.getbuffer().tobytes())
            keep = []
            for name, content in files_in_zip:
                low = name.lower()
                if low.endswith((".xls", ".xlsx", ".csv", ".txt")):
                    keep.append((name, content))

            st.info(f"{len(keep)} fichier(s) détecté(s) dans le ZIP (formats supportés).")
            for name, content in keep:
                tmp_path = save_bytes_to_tempfile(name, content)
                selected_paths.append((name, tmp_path))

    else:  # Scanner un dossier (local)
        folder = st.text_input("Chemin dossier (ex: /mnt/data ou C:\\\\...)")
        recursive = st.checkbox("Inclure sous-dossiers", value=True)
        if folder:
            p = Path(folder)
            if not p.exists():
                st.error("Dossier introuvable.")
            else:
                pattern = "**/*" if recursive else "*"
                files = [f for f in p.glob(pattern) if f.is_file()]
                files = [f for f in files if f.suffix.lower() in (".xls", ".xlsx", ".csv", ".txt")]
                st.info(f"{len(files)} fichier(s) trouvé(s).")
                # on ne copie pas en temp : on lit direct
                for f in files:
                    selected_paths.append((f.name, str(f)))

    # -----------------------------
    # Lecture & concat
    # -----------------------------
    if not selected_paths:
        st.warning("Aucun fichier importé pour le moment.")
        st.stop()

    st.subheader("Fichiers sélectionnés")
    st.write(pd.DataFrame(selected_paths, columns=["Nom", "Chemin temporaire / local"]))

    if st.button("🚀 Lancer l'import", type="primary"):
        dfs: List[Tuple[pd.DataFrame, str, str]] = []
        progress = st.progress(0)
        status = st.empty()

        total = len(selected_paths)

        for i, (name, path) in enumerate(selected_paths, start=1):
            status.write(f"Lecture ({i}/{total}) : **{name}**")
            try:
                df, mode = read_any(path)

                # option limite lignes
                if nrows_opt and nrows_opt > 0:
                    df = df.head(int(nrows_opt))

                df = normalize_df(df)
                dfs.append((df, name, mode))

            except Exception as e:
                st.error(f"❌ Erreur de lecture sur {name}")
                st.exception(e)

            progress.progress(int(i / total * 100))

        status.write("Terminé.")

        if not dfs:
            st.error("Aucun fichier n’a pu être lu.")
            st.stop()

        result = concat_with_origin(dfs)

        st.success(f"✅ Import terminé : {len(dfs)} fichier(s) lus, {result.shape[0]} lignes, {result.shape[1]} colonnes.")

        st.subheader("Aperçu (top 200 lignes)")
        st.dataframe(result.head(200), use_container_width=True)

        st.subheader("Colonnes détectées")
        st.write(list(result.columns))

        # Stats simples
        st.subheader("Contrôles rapides")
        st.write("Nb de lignes par fichier :")
        st.dataframe(
            result["_source_file"].value_counts().rename_axis("fichier").reset_index(name="lignes"),
            use_container_width=True
        )

        # Export CSV
        st.subheader("Export")
        csv_bytes = result.to_csv(index=False).encode("utf-8-sig")
        st.download_button(
            "⬇️ Télécharger le CSV consolidé",
            data=csv_bytes,
            file_name="journaux_consolides.csv",
            mime="text/csv"
        )

        # Nettoyage fichiers temporaires (upload seulement)
        # (on supprime seulement ceux qui sont dans /tmp)
        for _, path in selected_paths:
            try:
                if isinstance(path, str) and os.path.exists(path) and ("tmp" in path or "temp" in path):
                    os.unlink(path)
            except Exception:
                pass

except Exception as e:
    st.error("💥 Erreur au chargement de l'app (anti page blanche)")
    st.exception(e)
    st.stop()
