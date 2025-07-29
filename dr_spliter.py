import streamlit as st
import pandas as pd
from io import BytesIO
import re

# Configuration de la page
st.set_page_config(
    page_title="DR Spliter",
    page_icon="📊",
    layout="centered",
    initial_sidebar_state="expanded"
)

excel_img_link = "https://encrypted-tbn0.gstatic.com/images?q=tbn:ANd9GcSjm9RgAqdYle_Plh0SHAKY6OA3LOhqxwguYQ&s"
# Style CSS personnalisé
st.markdown("""
<style>
    .header-title {
        color: #1e3a8a;
        font-size: 2.5rem !important;
        text-align: center;
        padding: 10px;
        margin-bottom: 30px;
    }
    .success-box {
        background-color: #d1fae5;
        border-radius: 10px;
        padding: 15px;
        margin: 15px 0;
    }
    .error-box {
        background-color: #fee2e2;
        border-radius: 10px;
        padding: 15px;
        margin: 15px 0;
    }
    .info-box {
        background-color: #dbeafe;
        border-radius: 10px;
        padding: 15px;
        margin: 15px 0;
    }
    .stDownloadButton>button {
        background-color: #4f46e5 !important;
        color: white !important;
        font-weight: bold;
        border-radius: 8px;
        padding: 10px 24px;
        transition: all 0.3s;
    }
    .stDownloadButton>button:hover {
        background-color: #3730a3 !important;
        transform: scale(1.05);
    }
    .file-name {
        font-style: italic;
        word-break: break-all;
    }
</style>
""", unsafe_allow_html=True)

# Titre avec emojis et style
st.markdown('<h1 class="header-title">✨ Séparation Excel par Zone DRV</h1>', unsafe_allow_html=True)

# Zone d'upload
with st.container():
    st.subheader("📤 Téléversement du Fichier")
    uploaded_file = st.file_uploader(
        "Glissez-déposez votre fichier Excel ici",
        type=["xlsx"],
        help="Format supporté: .xlsx (Excel)",
        label_visibility="collapsed"
    )

if uploaded_file:
    try:
        # Afficher les informations du fichier
        file_details = st.expander("📝 Détails du fichier", expanded=True)
        with file_details:
            st.caption(f"**Nom du fichier:** <span class='file-name'>{uploaded_file.name}</span>",
                       unsafe_allow_html=True)
            st.caption(f"**Taille:** {(uploaded_file.size / 1024):.2f} KB")

        # Lecture du fichier
        with st.spinner("🔍 Analyse du fichier en cours..."):
            df = pd.read_excel(uploaded_file)

            if "zone_drvnew" not in df.columns:
                st.markdown('<div class="error-box">❌ Colonne "zone_drvnew" introuvable dans le fichier</div>',
                            unsafe_allow_html=True)
                st.error("Vérifiez que votre fichier contient bien cette colonne")
            else:
                # Statistiques
                zone_counts = df["zone_drvnew"].value_counts()
                unique_zones = len(zone_counts)

                st.markdown(f'<div class="success-box">✅ Fichier chargé avec succès!<br>'
                            f'• Zones détectées: {unique_zones}<br>'
                            f'• Lignes totales: {len(df)}</div>',
                            unsafe_allow_html=True)

                # Traitement
                with st.spinner("⚙️ Découpage des données par zone..."):
                    grouped = df.groupby("zone_drvnew")

                    output = BytesIO()
                    with pd.ExcelWriter(output, engine="xlsxwriter") as writer:
                        for name, group in grouped:
                            # Nettoyage du nom de feuille
                            sheet_name = re.sub(r'[\\/*?:\[\]]', '', str(name))
                            sheet_name = sheet_name[:31] if name else "inconnu"

                            if sheet_name == "":
                                sheet_name = "zone_vide"

                            group.to_excel(writer, sheet_name=sheet_name, index=False)

                    output.seek(0)

                # Résultat
                st.balloons()
                st.markdown(f'<div class="success-box">✨ Traitement terminé!<br>'
                            f'• Fichier découpé en {unique_zones} feuilles</div>',
                            unsafe_allow_html=True)

                # Bouton de téléchargement
                st.download_button(
                    label="📥 Télécharger le Fichier Séparé",
                    data=output.getvalue(),
                    file_name=f"ZONES_{uploaded_file.name}",
                    mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                    help="Cliquez pour télécharger le fichier séparé par zones"
                )

                # Aperçu des données
                st.subheader("👀 Aperçu des Données")
                st.dataframe(df.head(5))

    except Exception as e:
        st.markdown(f'<div class="error-box">❌ Erreur de traitement</div>', unsafe_allow_html=True)
        st.exception(e)
else:
    st.markdown('<div class="info-box">📌 Veuillez téléverser un fichier Excel pour commencer</div>',
                unsafe_allow_html=True)
    # st.image(excel_img_link, width=300, caption="Séparateur de fichiers Excel par zones")

# Pied de page
st.markdown("---")
st.caption("Made with ❤️ by Seny for DIANKHA | v1.2")