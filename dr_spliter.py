import streamlit as st
import pandas as pd
from io import BytesIO
import zipfile
import re

# ----------------------------------------------------------
# 🖥️ CONFIGURATION
# ----------------------------------------------------------
st.set_page_config(
    page_title="DRV Splitter",
    page_icon="📊",
    layout="centered"
)

# ----------------------------------------------------------
# 🔧 FONCTIONS
# ----------------------------------------------------------
def normalize_value(x):
    """Nettoyage pour un regroupement cohérent."""
    if pd.isna(x):
        return "INCONNU"
    x = str(x).upper()
    x = re.sub(r'\s+', '', x)        # remove spaces
    x = re.sub(r'[-_]', '', x)       # remove dashes/underscores
    x = re.sub(r'[^A-Z0-9]', '', x)  # keep only letters/numbers
    return x if x else "INCONNU"

def restore_human_format(value):
    """Transforme DAKAR1 → DAKAR 1 ; DAKAR2 → DAKAR 2."""
    match = re.match(r"^([A-Z]+)([0-9]+)$", value)
    if match:
        return f"{match.group(1)} {match.group(2)}"
    return value


# ----------------------------------------------------------
# 🎨 CSS
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
# 📤 UPLOAD
# ----------------------------------------------------------
uploaded_file = st.file_uploader("Téléverse ton fichier Excel (.xlsx)", type=["xlsx"])

if not uploaded_file:
    st.info("📌 En attente d’un fichier Excel…")
    st.stop()


# ----------------------------------------------------------
# 📄 LECTURE DES FEUILLES
# ----------------------------------------------------------
try:
    excel_file = pd.ExcelFile(uploaded_file)
    sheet = st.selectbox("📄 Choisis la feuille :", excel_file.sheet_names)

    df = pd.read_excel(excel_file, sheet_name=sheet)
    st.success(f"Feuille **{sheet}** chargée ({len(df)} lignes).")

except Exception as e:
    st.error("❌ Impossible de lire le fichier.")
    st.exception(e)
    st.stop()


# ----------------------------------------------------------
# 🧩 CHOIX COLONNE
# ----------------------------------------------------------
column = st.selectbox("🔎 Colonne à utiliser pour découper :", df.columns)


# ----------------------------------------------------------
# 🔍 NORMALISATION + VUE DES VALEURS
# ----------------------------------------------------------
df["__normalized__"] = df[column].apply(normalize_value)

st.markdown("### 🔍 Valeurs détectées (après normalisation)")

zone_counts = df["__normalized__"].value_counts().rename_axis("Valeur").reset_index(name="Nombre")
zone_counts["Valeur"] = zone_counts["Valeur"].apply(restore_human_format)

st.dataframe(zone_counts)


# ----------------------------------------------------------
# 📦 CHOIX MODE DE SORTIE
# ----------------------------------------------------------
output_mode = st.radio(
    "🗂️ Format du résultat final :",
    [
        "Un seul fichier Excel (plusieurs onglets)",
        "Plusieurs fichiers Excel séparés (ZIP)"
    ]
)


# ----------------------------------------------------------
# 🚀 SPLIT
# ----------------------------------------------------------
if st.button("🚀 Lancer le Split"):

    try:
        grouped = df.groupby("__normalized__")

        st.info(f"Découpage en **{len(grouped)} groupes**…")

        # ------------------------------------------------------
        # MODE 1 – UN SEUL FICHIER EXCEL
        # ------------------------------------------------------
        if output_mode == "Un seul fichier Excel (plusieurs onglets)":

            output = BytesIO()

            with pd.ExcelWriter(output, engine="xlsxwriter") as writer:
                for name, group in grouped:
                    sheet_name = restore_human_format(str(name))[:31]
                    group.to_excel(writer, sheet_name=sheet_name, index=False)

            output.seek(0)
            st.balloons()
            st.download_button(
                "📥 Télécharger le fichier Excel",
                data=output,
                file_name=f"SPLIT_{column}.xlsx",
                mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
            )


        # ------------------------------------------------------
        # MODE 2 – PLUSIEURS FICHIERS ZIP
        # ------------------------------------------------------
        else:

            zip_buffer = BytesIO()

            with zipfile.ZipFile(zip_buffer, "w") as zipf:
                for name, group in grouped:
                    final_name = restore_human_format(str(name))
                    file_bytes = BytesIO()
                    group.to_excel(file_bytes, index=False)
                    file_bytes.seek(0)
                    zipf.writestr(f"{final_name}.xlsx", file_bytes.read())

            zip_buffer.seek(0)
            st.balloons()
            st.download_button(
                "📥 Télécharger le ZIP",
                data=zip_buffer,
                file_name=f"SPLIT_{column}.zip",
                mime="application/zip"
            )

    except Exception as e:
        st.error("❌ Une erreur est survenue.")
        st.exception(e)


# ----------------------------------------------------------
# 👀 APERÇU
# ----------------------------------------------------------
st.markdown("### 👀 Aperçu (5 premières lignes)")
st.dataframe(df.head())


# ----------------------------------------------------------
# FOOTER
# ----------------------------------------------------------
st.markdown("---")
st.caption("Made with ❤️ by Seny for DIANKHA | v2.2 (Normalisation + Restore Format + Preview)")
