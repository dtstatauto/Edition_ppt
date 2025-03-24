import streamlit as st
import pandas as pd
from pptx import Presentation
from docx import Document
from io import BytesIO
import utils 

# Affichage du logo
st.image("templates/logo-aon.jpg", width=150)

# Interface utilisateur Streamlit
st.title("Générateur automatique de fichiers")

# Chemins vers les modèles PowerPoint et Word
chemin_template_flottes = "templates/ppt_flottes.pptx"
chemin_template_mission = "templates/ppt_missions.pptx"
chemin_template_word = "templates/word.docx"

# Dictionnaires des espaces réservés pour les contrats "flottes" et "missions"
placeholders_flottes = {
    "client": "client", "date": "date", "nom": "nom", "adresse": "adresse",
    "cp": "cp", "ville": "ville", "leger": "leger", "lourd": "lourd",
    "semirem": "semirem", "reminf": "reminf", "deuxroues": "droue",
    "engins": "engins", "assureur": "ass", "echeann": "echeann",
    "frac": "frac", "reg": "reg", "fin": "fin"
}

placeholders_missions = {
    "client": "client", "date": "date", "nom": "nom", "adresse": "adresse",
    "cp": "cp", "ville": "ville", "assureur": "ass", "echeann": "echeann",
    "frac": "frac", "fin": "fin"
}

placeholders_word = {
    "nom": "nom", "date": "date", "adresse": "adresse",
    "cp": "cp", "ville": "ville", "assureur": "assureur", "camionette": "camionette",
    "camion": "camion", "deuxroues": "deuxroues", "engins": "engins", "autre": "autre",
    "effet": "effet", "siret": "siret", "activite": "activite", "risque": "risque"
}

# Téléchargement du fichier Excel
excel_file = st.file_uploader("Choisissez le fichier Excel", type=["xlsx", "xls"])

if excel_file is not None:
    # Lire les noms des feuilles du fichier Excel
    xls = pd.ExcelFile(excel_file)
    sheet_names = xls.sheet_names

    # Sélection de la feuille par l'utilisateur
    selected_sheet = st.selectbox("Choisissez l'onglet à traiter", sheet_names)

    # Lire les données de la feuille sélectionnée
    df = pd.read_excel(excel_file, sheet_name=selected_sheet)

    # Afficher les données de la feuille sélectionnée
    st.write(f"Données de l'onglet '{selected_sheet}':")
    st.dataframe(df)

    # Sélection du client par l'utilisateur
    clients = df['client'].unique()
    client_selection = st.selectbox("Choisissez un client", clients)

    # Cases à cocher pour sélectionner les modes
    mode_flotte = st.checkbox("Mode Flotte")
    mode_mission = st.checkbox("Mode Mission")
    mode_word = st.checkbox("Mode Word")

    # Bouton pour générer les présentations "Flottes"
    if mode_flotte and st.button("Générer PowerPoint Flottes"):
        ppt_files = utils.generate_ppt(chemin_template_flottes, excel_file, selected_sheet, client_selection, placeholders_flottes)
        for filename, ppt_io in ppt_files:
            st.download_button(
                label=f"📥 Télécharger {filename}",
                data=ppt_io,
                file_name=filename,
                mime="application/vnd.openxmlformats-officedocument.presentationml.presentation"
            )
        st.balloons()  # Affiche une animation de ballons après la génération des présentations

    # Bouton pour générer les présentations "Missions"
    if mode_mission and st.button("Générer PowerPoint Mission"):
        ppt_files = utils.generate_ppt(chemin_template_mission, excel_file, selected_sheet, client_selection, placeholders_missions)
        for filename, ppt_io in ppt_files:
            st.download_button(
                label=f"📥 Télécharger {filename}",
                data=ppt_io,
                file_name=filename,
                mime="application/vnd.openxmlformats-officedocument.presentationml.presentation"
            )
        st.balloons()  # Affiche une animation de ballons après la génération des présentations

    # Bouton pour générer le document Word
    if mode_word and st.button("Générer Word"):
        word_file = utils.remplir_document_word(chemin_template_word, excel_file, selected_sheet, client_selection, placeholders_word)
        st.download_button(
            label="📥 Télécharger le document Word",
            data=word_file,
            file_name="document_rempli.docx",
            mime="application/vnd.openxmlformats-officedocument.wordprocessingml.document"
        )
        st.balloons()  # Affiche une animation de ballons après la génération du document Word
else:
    st.warning("Veuillez télécharger un fichier Excel.")
