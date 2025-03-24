import streamlit as st
import pandas as pd
from pptx import Presentation
from io import BytesIO

def edition_ppt(template_path, excel_file, client_selection):
    """ Remplit le PowerPoint modèle avec les données de chaque ligne du fichier Excel. """

    data = pd.read_excel(excel_file)
    filtered_data = data[(data['client'] == client_selection)]

    ppt_files = []

    for index, row in filtered_data.iterrows():
        prs = Presentation(template_path)

        for slide in prs.slides:
            for shape in slide.shapes:
                if shape.has_text_frame:
                    for paragraph in shape.text_frame.paragraphs:
                        for run in paragraph.runs:
                            run.text = run.text.replace("client", str(row["client"]))
                            run.text = run.text.replace("date", str(row["date"]))
                            run.text = run.text.replace("nom", str(row["nom"]))
                            run.text = run.text.replace("adresse", str(row["adresse"]))
                            run.text = run.text.replace("cp", str(row["cp"]))
                            run.text = run.text.replace("ville", str(row["ville"]))
                            run.text = run.text.replace("leger", str(row["leger"]))
                            run.text = run.text.replace("lourd", str(row["lourd"]))
                            run.text = run.text.replace("semirem", str(row["semirem"]))
                            run.text = run.text.replace("reminf", str(row["reminf"]))
                            run.text = run.text.replace("deuxroues", str(row["droue"]))
                            run.text = run.text.replace("engins", str(row["engins"]))
                            run.text = run.text.replace("assureur", str(row["ass"]))
                            run.text = run.text.replace("echeann", str(row["echeann"]))
                            run.text = run.text.replace("frac", str(row["frac"]))
                            run.text = run.text.replace("reg", str(row["reg"]))
                            run.text = run.text.replace("fin", str(row["fin"]))
                            
                            
                            
                            
        ppt_io = BytesIO()
        prs.save(ppt_io)
        ppt_io.seek(0)
        ppt_files.append((f"PROJET_REMISE_OFFRES_CONVENTIONS_{index + 1}.pptx", ppt_io))

    return ppt_files

# Interface utilisateur Streamlit
st.title("Générateur de PowerPoint")

# Chemin vers le modèle PowerPoint
chemin_template = "templates/ppt_flottes.pptx"


# Téléchargement du fichier Excel
excel_file = st.file_uploader("Choisissez un fichier Excel", type=["xlsx"])

if excel_file is not None:
    data = pd.read_excel(excel_file)
    clients = data['client'].unique()

    # Sélection du client et du type de risque
    client_selection = st.selectbox("Choisissez un client", clients)

    if st.button("Générer PowerPoint"):
        ppt_files = edition_ppt(chemin_template, excel_file, client_selection)
        for filename, ppt_io in ppt_files:
            st.download_button(
                label=f"Télécharger {filename}",
                data=ppt_io,
                file_name=filename,
                mime="application/vnd.openxmlformats-officedocument.presentationml.presentation"
            )
else:
    st.warning("Veuillez télécharger un fichier Excel.")
