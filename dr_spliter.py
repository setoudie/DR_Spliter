import streamlit as st
import pandas as pd
from io import BytesIO
import zipfile
import re

# ----------------------------------------------------------
# 🖥️ CONFIGURATION DE LA PAGE
# ----------------------------------------------------------
st.set_page_config(
    page_title="DRV Splitter",
    page_icon="📊",
    layout="centered"
)

# ----------------------------------------------------------
# 🎨 CSS DESIGN (version améliorée)
# ----------------------------------------------------------
st.markdown("""
<style>
    .header-title {
        color: #1e3a8a;
        font-size: 3rem !important;
        text-align: center;
        margin-bottom: 20px;
        font-weight: 700;
    }
    .step-box {
        background: #f3f4f6;
        padding: 18px;
        border-radius: 12px;
        border-left: 6px solid #4f46e5;
        margin-bottom: 20px;
    }
    .success-box {
        background-color: #d1fae5;
        border-radius: 10px;
        padding: 15px;
        margin: 15px 0;
        border-left: 5px solid #059669;
    }
    .error-box {
        background-color: #fee2e2;
        border-radius: 10px;
        padding: 15px;
        margin: 15px 0;
        border-left: 5px solid #dc2626;
    }
    .info-box {
        background-color: #e0f2fe;
        border-radius: 10px;
        padding: 15px;
        margin: 15px 0;
        border-left: 5px solid #0284c7;
    }
    .stDownloadButton>button {
        background-color: #4f46e5 !important;
        color: white !important;
        font-weight: bold;
        border-radius: 10px;
        padding: 12px 28px;
        transition: 0.25s;
    }
    .stDownloadButton>button:hover {
        background-color: #3730a3 !important;
        transform: scale(1.05);
    }
</style>
""", unsafe_allow_html=True)

# ----------------------------------------------------------
# 🏷️ TITRE
# ----------------------------------------------------------
st.markdown('<h1 class="header-title">✨ DRV SPLITTER – Version Avancée</h1>', unsafe_allow_html=True)

# ----------------------------------------------------------
# 📤 UPLOAD FICHIER
# ----------------------------------------------------------
uploaded_file = st.file_uploader(
    "Téléverse ton fichier Excel (.xlsx)",
    type=["xlsx"]
)

if not uploaded_file:
    st.markdown('<div class="info-box">📌 En attente d’un fichier Excel…</div>', unsafe_allow_html=True)
    st.stop()

# ----------------------------------------------------------
# 📄 LECTURE DES FEUILLES DU FICHIER
# ----------------------------------------------------------
try:
    excel_file = pd.ExcelFile(uploaded_file)
    sheet = st.selectbox("📄 Choisis la feuille sur laquelle travailler :", excel_file.sheet_names)

    df = pd.read_excel(excel_file, sheet_name=sheet)

    st.success(f"Feuille chargée : **{sheet}** ({len(df)} lignes)")

except Exception as e:
    st.markdown('<div class="error-box">❌ Impossible de lire le fichier.</div>', unsafe_allow_html=True)
    st.exception(e)
    st.stop()

# ----------------------------------------------------------
# 🔍 CHOIX DE LA COLONNE À SPLITTER
# ----------------------------------------------------------
column = st.selectbox(
    "🔎 Sélectionne la colonne sur laquelle découper :",
    df.columns
)

# ----------------------------------------------------------
# 📦 CHOIX DU MODE DE SORTIE
# ----------------------------------------------------------
output_mode = st.radio(
    "🗂️ Comment veux-tu recevoir le résultat ?",
    [
        "Un seul fichier Excel (plusieurs onglets)",
        "Plusieurs fichiers Excel séparés (ZIP)"
    ]
)

# ----------------------------------------------------------
# 🚀 BOUTON DE TRAITEMENT
# ----------------------------------------------------------
if st.button("🚀 Lancer le Split"):
    try:
        grouped = df.groupby(column)

        st.info(f"Découpage en **{len(grouped)} groupes**…")

        # ------------------------------------------------------
        # MODE 1 – UN SEUL FICHIER EXCEL AVEC FEUILLES
        # ------------------------------------------------------
        if output_mode == "Un seul fichier Excel (plusieurs onglets)":
            output = BytesIO()
            with pd.ExcelWriter(output, engine="xlsxwriter") as writer:

                for name, group in grouped:
                    sheet_name = re.sub(r'[\\/*?:\[\]]', '', str(name))[:31] or "inconnu"
                    group.to_excel(writer, sheet_name=sheet_name, index=False)

            output.seek(0)
            st.balloons()
            st.download_button(
                label="📥 Télécharger le fichier Excel",
                data=output,
                file_name=f"SPLIT_{column}.xlsx",
                mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
            )

        # ------------------------------------------------------
        # MODE 2 – PLUSIEURS EXCEL DANS UN ZIP
        # ------------------------------------------------------
        else:
            zip_buffer = BytesIO()
            with zipfile.ZipFile(zip_buffer, "w") as zipf:
                for name, group in grouped:
                    clean_name = re.sub(r'[\\/*?:\[\]]', '', str(name)) or "inconnu"
                    file_bytes = BytesIO()
                    group.to_excel(file_bytes, index=False)
                    file_bytes.seek(0)
                    zipf.writestr(f"{clean_name}.xlsx", file_bytes.read())

            zip_buffer.seek(0)
            st.balloons()
            st.download_button(
                label="📥 Télécharger le ZIP",
                data=zip_buffer,
                file_name=f"SPLIT_{column}.zip",
                mime="application/zip"
            )

    except Exception as e:
        st.markdown('<div class="error-box">❌ Une erreur est survenue.</div>', unsafe_allow_html=True)
        st.exception(e)

# ----------------------------------------------------------
# 👀 APERÇU DES DONNÉES
# ----------------------------------------------------------
st.markdown("### 👀 Aperçu des 5 premières lignes")
st.dataframe(df.head())

# ----------------------------------------------------------
# FOOTER
# ----------------------------------------------------------
st.markdown("---")
st.caption("Made with ❤️ by Seny for DIANKHA | v2.0 (Optimisé & Stylé)")
