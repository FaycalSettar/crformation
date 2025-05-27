import streamlit as st
import pandas as pd
from docx import Document
import os
import tempfile
from zipfile import ZipFile

st.set_page_config(page_title="Compte Rendu Formation", layout="centered")
st.title("Générateur de Comptes Rendus de Formation")

def remplacer_placeholders(paragraph, replacements):
    for key, value in replacements.items():
        if key in paragraph.text:
            for run in paragraph.runs:
                if key in run.text:
                    run.text = run.text.replace(key, value)

def cocher_checkbox(doc, choix):
    for para in doc.paragraphs:
        text = para.text.strip()
        for item, valeur in choix.items():
            if item in text:
                if valeur in text:
                    para.text = text.replace("{{checkbox}}", "☑")
                else:
                    para.text = text.replace("{{checkbox}}", "☐")

with st.expander("Étape 1 : Import des fichiers", expanded=True):
    excel_file = st.file_uploader("Fichier Excel des participants", type="xlsx")
    word_file = st.file_uploader("Modèle Word du compte rendu", type="docx")

if excel_file and word_file:
    df = pd.read_excel(excel_file)
    df.columns = df.columns.str.strip()  # Nettoyage des noms de colonnes

    required_cols = {"session", "Prénom", "Nom", "formation", "nb d'heure", "formateur"}
    if not required_cols.issubset(df.columns):
        st.error(f"❌ Colonnes manquantes : {', '.join(required_cols - set(df.columns))}")
        st.stop()

    sessions = df.groupby("session")

    st.markdown("### Étape 2 : Complétez les informations manuellement")
    session_configs = {}

    for session_id, participants in sessions:
        st.subheader(f"Session : {session_id}")

        satisfaction = st.radio(f"Satisfaction globale – Session {session_id}",
                                ["Très satisfait", "Satisfait", "Moyennement satisfait", "Insatisfait", "Non satisfait"],
                                key=f"satisfaction_{session_id}")

        motivation = st.radio(f"Motivation des participants – Session {session_id}",
                              ["Très motivés", "Motivés", "Pas motivés"],
                              key=f"motivation_{session_id}")

        assiduite = st.radio(f"Assiduité des participants – Session {session_id}",
                             ["Très motivés", "Motivés", "Pas motivés"],
                             key=f"assiduite_{session_id}")

        reponses = st.radio(f"Réponses données – Session {session_id}",
                            ["Toutes les questions", "A peu près toutes",
                             "Il y a quelques sujets sur lesquels je n'avais pas les réponses",
                             "Je n'ai pas pu répondre à la majorité des questions"],
                            key=f"reponses_{session_id}")

        adaptation = st.radio(f"Adaptation du déroulé – Session {session_id}", ["Oui", "Non"],
                              key=f"adaptation_{session_id}")

        suivi = st.radio(f"Fichier de suivi mis à jour – Session {session_id}",
                         ["Oui", "Non", "Non concerné"], key=f"suivi_{session_id}")

        pistes = st.text_area(f"Pistes d'amélioration – Session {session_id}", key=f"pistes_{session_id}")
        observations = st.text_area(f"Observations libres – Session {session_id}", key=f"obs_{session_id}")

        session_configs[session_id] = {
            "satisfaction": satisfaction,
            "motivation": motivation,
            "assiduite": assiduite,
            "reponses": reponses,
            "adaptation": adaptation,
            "suivi": suivi,
            "pistes": pistes,
            "observations": observations
        }

    if st.button("📄 Générer les comptes rendus", type="primary"):
        with tempfile.TemporaryDirectory() as tmpdir:
            zip_path = os.path.join(tmpdir, "Comptes_Rendus.zip")
            recap = []

            with ZipFile(zip_path, 'w') as zipf:
                for session_id, participants in sessions:
                    doc = Document(word_file)

                    # Infos de session
                    first_row = participants.iloc[0]
                    nb_participants = len(participants)

                    replacements = {
                        "{{formateur}}": str(first_row["formateur"]),
                        "{{ref_session}}": str(session_id),
                        "{{formation_dispensee}}": str(first_row["formation"]),
                        "{{duree_formation}}": str(first_row["nb d'heure"]),
                        "{{nb_participants}}": str(nb_participants)
                    }

                    for para in doc.paragraphs:
                        remplacer_placeholders(para, replacements)
                    for table in doc.tables:
                        for row in table.rows:
                            for cell in row.cells:
                                for para in cell.paragraphs:
                                    remplacer_placeholders(para, replacements)

                    # Cocher les cases
                    cocher_checkbox(doc, session_configs[session_id])

                    # Ajouter observations
                    doc.add_paragraph("Avis & piste d'amélioration :\n" + session_configs[session_id]["pistes"])
                    doc.add_paragraph("Autres observations :\n" + session_configs[session_id]["observations"])

                    filename = f"Compte_Rendu_{session_id}.docx"
                    path = os.path.join(tmpdir, filename)
                    doc.save(path)
                    zipf.write(path, arcname=filename)

                    recap.append({
                        "Session": session_id,
                        "Formateur": first_row["formateur"],
                        "Participants": nb_participants,
                        "Satisfaction": session_configs[session_id]["satisfaction"]
                    })

                recap_path = os.path.join(tmpdir, "Recapitulatif.xlsx")
                pd.DataFrame(recap).to_excel(recap_path, index=False)
                zipf.write(recap_path, arcname="Recapitulatif.xlsx")

            with open(zip_path, "rb") as f:
                st.success("✅ Comptes rendus générés avec succès !")
                st.download_button("📥 Télécharger l'archive ZIP", data=f, file_name="Comptes_Rendus.zip", mime="application/zip")
