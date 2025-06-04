import streamlit as st
import pandas as pd
from docx import Document
import os
import tempfile
from zipfile import ZipFile
import random
from collections import defaultdict

st.set_page_config(page_title="Générateur de QCM par session", layout="centered")
st.title("📄 Générateur de QCM par session (figé ou aléatoire)")

# Fonction de remplacement des balises
def remplacer_placeholders(paragraph, replacements):
    for key in replacements:
        if key in paragraph.text:
            full_text = paragraph.text
            while key in full_text:
                for run in paragraph.runs:
                    if key in run.text:
                        run.text = run.text.replace(key, replacements[key])
                full_text = full_text.replace(key, replacements[key], 1)

# Fonction pour itérer sur tous les paragraphes
def iter_all_paragraphs(doc):
    for para in doc.paragraphs:
        yield para
    for table in doc.tables:
        for row in table.rows:
            for cell in row.cells:
                for para in cell.paragraphs:
                    yield para

# Constantes
POSITIVE_OPTIONS = {
    "satisfaction": ["Très satisfait", "Satisfait"],
    "motivation": ["Très motivés", "Motivés"],
    "assiduite": ["Très motivés", "Motivés"],
    "homogeneite": ["Oui"],
    "questions": ["Toutes les questions", "A peu près toutes"],
}

CHECKBOX_GROUPS = {
    "motivation": ["Très motivés", "Motivés", "Pas motivés"],
    "assiduite": ["Très motivés", "Motivés", "Pas motivés"],
    "homogeneite": ["Oui", "Non"],
    "questions": ["Toutes les questions", "A peu près toutes", "Il y a quelques sujets", "Je n'ai pas pu répondre"],
    "adaptation": ["Oui", "Non"],
    "suivi": ["Oui", "Non", "Non concerné"],
    "satisfaction": ["Très satisfait", "Satisfait", "Moyennement satisfait", "Insatisfait", "Non satisfait"]
}

# Interface utilisateur
with st.expander("1. Importer les fichiers", expanded=True):
    excel_file = st.file_uploader("Fichier Excel des participants", type="xlsx")
    word_file = st.file_uploader("Modèle Word (.docx)", type="docx")

if excel_file and word_file:
    df = pd.read_excel(excel_file)
    df.columns = df.columns.str.strip()
    
    required_columns = ["session", "formateur", "formation", "nb d'heure", "Nom", "Prénom"]
    if not set(required_columns).issubset(df.columns):
        st.error(f"Colonnes manquantes. Requises : {required_columns}")
        st.info(f"Colonnes disponibles : {list(df.columns)}")
    else:
        reponses_figees = {}

        st.markdown("### 2. Réponses à figer (facultatif)")
        for groupe in CHECKBOX_GROUPS:
            if groupe in ["adaptation", "suivi"]:
                continue
            if st.checkbox(f"Figer : {groupe}", key=f"fig_{groupe}"):
                choix = st.selectbox(
                    f"Réponse pour {groupe}",
                    CHECKBOX_GROUPS[groupe],
                    key=f"sel_{groupe}"
                )
                reponses_figees[groupe] = choix

        pistes = st.text_area("Avis & pistes d'amélioration")
        obs = st.text_area("Autres observations")

        if st.button("🚀 Générer les comptes rendus"):
            with tempfile.TemporaryDirectory() as tmpdir:
                zip_path = os.path.join(tmpdir, "QCM.zip")
                with ZipFile(zip_path, 'w') as zipf:
                    for session_id, participants in df.groupby("session"):
                        doc = Document(word_file)
                        first = participants.iloc[0]

                        # Remplacements généraux
                        remplacements = {
                            "{{nom}}": first["Nom"],
                            "{{prénom}}": first["Prénom"],
                            "{{ref_session}}": str(session_id),
                            "{{formation_dispensee}}": first["formation"],
                            "{{duree_formation}}": str(first["nb d'heure"]),
                            "{{nb_participants}}": str(len(participants))
                        }

                        # Appliquer les remplacements
                        for para in iter_all_paragraphs(doc):
                            remplacer_placeholders(para, remplacements)

                        # Collecte des checkboxes
                        checkbox_paras = []
                        for para in iter_all_paragraphs(doc):
                            if "{{checkbox}}" in para.text:
                                texte_brut = para.text.lower().replace(" ", "").replace(" ", "")
                                for groupe, options in CHECKBOX_GROUPS.items():
                                    for opt in options:
                                        if opt.lower().replace(" ", "") in texte_brut:
                                            checkbox_paras.append((groupe, opt, para))
                                            break
                                    else:
                                        continue
                                    break

                        # Regrouper par groupe
                        group_to_paras = defaultdict(list)
                        for groupe, opt, para in checkbox_paras:
                            group_to_paras[groupe].append((opt, para))

                        # Appliquer les réponses
                        for groupe, paras in group_to_paras.items():
                            if groupe == "adaptation":
                                option_choisie = "Non"
                            elif groupe == "suivi":
                                option_choisie = "Non concerné"
                            elif groupe in reponses_figees:
                                option_choisie = reponses_figees[groupe]
                            else:
                                positives = [
                                    opt for opt, _ in paras
                                    if groupe in POSITIVE_OPTIONS and opt in POSITIVE_OPTIONS[groupe]
                                ]
                                option_choisie = random.choice(positives) if positives else random.choice([opt for opt, _ in paras])

                            # Appliquer le choix
                            for opt, para in paras:
                                for run in para.runs:
                                    if "{{checkbox}}" in run.text:
                                        run.text = run.text.replace("{{checkbox}}", "☑" if opt == option_choisie else "☐")

                        # Ajouter les commentaires
                        doc.add_paragraph("\nAvis & piste d'amélioration de la formation :\n" + pistes)
                        doc.add_paragraph("\nAutres observations (Exprimez-vous librement) :\n" + obs)

                        # Sauvegarder
                        path = os.path.join(tmpdir, f"CR_{session_id}.docx")
                        doc.save(path)
                        zipf.write(path, arcname=f"CR_{session_id}.docx")

                # Téléchargement
                with open(zip_path, "rb") as f:
                    st.success("✅ Fichiers générés !")
                    st.download_button("📥 Télécharger ZIP", f, "QCM.zip", "application/zip")
