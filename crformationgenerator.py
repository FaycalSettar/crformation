import streamlit as st
import pandas as pd
from docx import Document
import os
import tempfile
from zipfile import ZipFile
import random
import re

st.set_page_config(page_title="Générateur de QCM par session", layout="centered")
st.title("📄 Générateur de QCM par session (figé ou aléatoire)")

# Fonction de remplacement des balises
def remplacer_placeholders(paragraph, replacements):
    for key, val in replacements.items():
        if key in paragraph.text:
            for run in paragraph.runs:
                if key in run.text:
                    run.text = run.text.replace(key, val)

# Détection des blocs de checkbox
CHECKBOX_GROUPS = {
    "satisfaction": ["Très satisfait", "Satisfait", "Moyennement satisfait", "Insatisfait", "Non satisfait"],
    "motivation": ["Très motivés", "Motivés", "Pas motivés"],
    "assiduite": ["Très motivés", "Motivés", "Pas motivés"],
    "homogeneite": ["Oui", "Non"],
    "questions": ["Toutes les questions", "A peu près toutes", "Il y a quelques sujets sur lesquels je n'avais pas les réponses", "Je n'ai pas pu répondre à la majorité des questions"],
    "adaptation": ["Oui", "Non"],
    "suivi": ["Oui", "Non", "Non concerné"]
}

# Étape 1 : Importer les fichiers
with st.expander("Etape 1 : Importer les fichiers", expanded=True):
    excel_file = st.file_uploader("Fichier Excel des participants", type="xlsx")
    word_file = st.file_uploader("Modèle Word du compte rendu", type="docx")

# Traitement
if excel_file and word_file:
    df = pd.read_excel(excel_file)
    df.columns = df.columns.str.strip()

    if not set(["session", "formateur", "formation", "nb d'heure", "Nom", "Prénom"]).issubset(df.columns):
        st.error("Colonnes manquantes dans le fichier Excel.")
    else:
        sessions = df.groupby("session")
        reponses_figees = {}

        st.markdown("### Etape 2 : Choisir les réponses à figer (facultatif)")
        for groupe, options in CHECKBOX_GROUPS.items():
            figer = st.checkbox(f"Figer la réponse pour : {groupe}", key=f"figer_{groupe}")
            if figer:
                choix = st.selectbox(f"Choix figé pour {groupe}", options, key=f"choix_{groupe}")
                reponses_figees[groupe] = choix

        pistes = st.text_area("Avis & pistes d'amélioration :", key="pistes")
        observations = st.text_area("Autres observations :", key="obs")

        if st.button("🚀 Générer les comptes rendus"):
            with tempfile.TemporaryDirectory() as tmpdir:
                zip_path = os.path.join(tmpdir, "QCM_Sessions.zip")
                with ZipFile(zip_path, 'w') as zipf:
                    for session_id, participants in sessions:
                        doc = Document(word_file)
                        first = participants.iloc[0]
                        replacements = {
                            "{{formateur}}": str(first["formateur"]),
                            "{{ref_session}}": str(session_id),
                            "{{formation_dispensee}}": str(first["formation"]),
                            "{{duree_formation}}": str(first["nb d'heure"]),
                            "{{nb_participants}}": str(len(participants))
                        }
                        # Remplacement des placeholders dans les paragraphes
                        for para in doc.paragraphs:
                            remplacer_placeholders(para, replacements)
                        # Remplacement des placeholders dans les tables
                        for table in doc.tables:
                            for row in table.rows:
                                for cell in row.cells:
                                    for para in cell.paragraphs:
                                        remplacer_placeholders(para, replacements)

                        # Cocher les réponses (version corrigée) :
                        for para in doc.paragraphs:
                            texte = para.text.strip().replace(" ", " ")
                            for groupe, options in CHECKBOX_GROUPS.items():
                                for opt in options:
                                    # Correspondance mot-à-mot pour éviter les sous-chaînes
                                    if re.search(rf"\b{re.escape(opt)}\b", texte):
                                        # Si on a fixé ce groupe et que le choix correspond à opt
                                        if reponses_figees.get(groupe) == opt:
                                            para.text = texte.replace("{{checkbox}}", "☑")
                                        # Si le groupe est figé mais opt n'est pas celui choisi
                                        elif groupe in reponses_figees:
                                            para.text = texte.replace("{{checkbox}}", "☐")
                                        # Si pas de figé, on tire au sort
                                        elif groupe not in reponses_figees and random.choice([True, False]):
                                            para.text = texte.replace("{{checkbox}}", "☑")
                                        else:
                                            para.text = texte.replace("{{checkbox}}", "☐")

                                        # Mise à jour de texte après remplacement
                                        texte = para.text
                                        # On arrête de tester les autres options de ce groupe
                                        break
                                # Si une option de ce groupe a déjà été traitée, on sort de la boucle groupe
                                if re.search(rf"\b({'|'.join(map(re.escape, options))})\b", texte):
                                    break

                        # Ajout des sections "Avis & pistes d'amélioration" et "Autres observations"
                        doc.add_paragraph("\nAvis & pistes d'amélioration :\n" + pistes)
                        doc.add_paragraph("\nAutres observations :\n" + observations)

                        # Enregistrement du document pour chaque session
                        filename = f"Compte_Rendu_{session_id}.docx"
                        path = os.path.join(tmpdir, filename)
                        doc.save(path)
                        zipf.write(path, arcname=filename)

                # Bouton de téléchargement de l'archive ZIP
                with open(zip_path, "rb") as f:
                    st.success("Comptes rendus générés avec succès !")
                    st.download_button(
                        "📅 Télécharger l'archive ZIP",
                        data=f,
                        file_name="QCM_Sessions.zip",
                        mime="application/zip"
                    )
