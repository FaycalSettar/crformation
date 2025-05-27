import streamlit as st
import pandas as pd
from docx import Document
import os
import tempfile
from zipfile import ZipFile
import re

st.set_page_config(page_title="Compte Rendu Formation", layout="centered")
st.title("Générateur de Comptes Rendus de Formation")

def remplacer_placeholders(paragraph, replacements):
    for key, value in replacements.items():
        if key in paragraph.text:
            inline = paragraph.runs
            for i in range(len(inline)):
                if key in inline[i].text:
                    inline[i].text = inline[i].text.replace(key, value)

def cocher_checkbox(doc, choix):
    for para in doc.paragraphs:
        text = para.text.strip()
        for item, valeur in choix.items():
            if item in text:
                if valeur in text:
                    para.text = text.replace("{{checkbox}}", "☑")
                else:
                    para.text = text.replace("{{checkbox}}", "☐")

with st.expander("Étape 1 : Importer les fichiers", expanded=True):
    excel_file = st.file_uploader("Fichier Excel des participants", type="xlsx")
    word_file = st.file_uploader("Modèle Word de compte rendu", type="docx")

if excel_file and word_file:
    df = pd.read_excel(excel_file)
    if "session" not in df.columns or "Nom" not in df.columns or "Prénom" not in df.columns:
        st.error("Le fichier Excel doit contenir les colonnes : session, Prénom, Nom")
    else:
        sessions = df.groupby("session")

        st.markdown("### Étape 2 : Configurer les réponses manuelles")
        session_configs = {}

        for session_id, participants in sessions:
            st.subheader(f"Session : {session_id}")
            nom_formateur = st.text_input(f"Nom et prénom du formateur – Session {session_id}", key=f"formateur_{session_id}")
            formation_dispensee = st.text_input(f"Formation dispensée – Session {session_id}", key=f"formation_{session_id}")
            duree = st.text_input(f"Durée (en heures) – Session {session_id}", key=f"duree_{session_id}")

            satisfaction = st.radio(f"Niveau global de satisfaction – Session {session_id}",
                                    ["Très satisfait", "Satisfait", "Moyennement satisfait", "Insatisfait", "Non satisfait"],
                                    key=f"satisfaction_{session_id}")

            motivation = st.radio(f"Motivation des participants – Session {session_id}",
                                  ["Très motivés", "Motivés", "Pas motivés"],
                                  key=f"motivation_{session_id}")

            assiduite = st.radio(f"Assiduité des participants – Session {session_id}",
                                 ["Très motivés", "Motivés", "Pas motivés"],
                                 key=f"assiduite_{session_id}")

            reponses = st.radio(f"Réponses aux questions – Session {session_id}",
                                ["Toutes les questions", "A peu près toutes", "Il y a quelques sujets sur lesquels je n'avais pas les réponses", "Je n'ai pas pu répondre à la majorité des questions"],
                                key=f"reponses_{session_id}")

            adaptation = st.radio(f"Adaptation du déroulé – Session {session_id}",
                                  ["Oui", "Non"],
                                  key=f"adaptation_{session_id}")

            suivi = st.radio(f"Fichier de suivi mis à jour – Session {session_id}",
                             ["Oui", "Non", "Non concerné"],
                             key=f"suivi_{session_id}")

            pistes = st.text_area(f"Pistes d'amélioration – Session {session_id}", key=f"pistes_{session_id}")
            observations = st.text_area(f"Autres observations – Session {session_id}", key=f"observations_{session_id}")

            session_configs[session_id] = {
                "nom_formateur": nom_formateur,
                "formation_dispensee": formation_dispensee,
                "duree": duree,
                "nb_participants": str(len(participants)),
                "satisfaction": satisfaction,
                "motivation": motivation,
                "assiduite": assiduite,
                "reponses": reponses,
                "adaptation": adaptation,
                "suivi": suivi,
                "pistes": pistes,
                "observations": observations
            }

        if st.button("Générer les comptes rendus", type="primary"):
            with tempfile.TemporaryDirectory() as tmpdir:
                zip_path = os.path.join(tmpdir, "Comptes_Rendus.zip")
                recap_data = []

                with ZipFile(zip_path, 'w') as zipf:
                    for session_id, config in session_configs.items():
                        doc = Document(word_file)
                        replacements = {
                            "{{nom}}": config["nom_formateur"],
                            "{{ref_session}}": session_id,
                            "{{formation_dispensee}}": config["formation_dispensee"],
                            "{{duree_formation}}": config["duree"],
                            "{{nb_participants}}": config["nb_participants"]
                        }
                        for para in doc.paragraphs:
                            remplacer_placeholders(para, replacements)
                        for table in doc.tables:
                            for row in table.rows:
                                for cell in row.cells:
                                    for para in cell.paragraphs:
                                        remplacer_placeholders(para, replacements)

                        cocher_checkbox(doc, config)

                        doc.add_paragraph("Avis & piste d'amélioration de la formation :\n" + config["pistes"])
                        doc.add_paragraph("Autres observations :\n" + config["observations"])

                        filename = f"Compte_Rendu_{session_id.replace(' ', '_')}.docx"
                        filepath = os.path.join(tmpdir, filename)
                        doc.save(filepath)
                        zipf.write(filepath, filename)

                        recap_data.append({
                            "Session": session_id,
                            "Formateur": config["nom_formateur"],
                            "Participants": config["nb_participants"],
                            "Satisfaction": config["satisfaction"]
                        })

                    df_recap = pd.DataFrame(recap_data)
                    recap_path = os.path.join(tmpdir, "Recapitulatif_Compte_Rendus.xlsx")
                    df_recap.to_excel(recap_path, index=False)
                    zipf.write(recap_path, "Recapitulatif_Compte_Rendus.xlsx")

                with open(zip_path, "rb") as f:
                    st.success("✅ Comptes rendus générés avec succès !")
                    st.download_button("📥 Télécharger l'archive ZIP", data=f, file_name="Comptes_Rendus.zip", mime="application/zip")
