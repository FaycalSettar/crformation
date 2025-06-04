import streamlit as st
import pandas as pd
from docx import Document
import os
import tempfile
from zipfile import ZipFile
import random
import re
from collections import defaultdict

st.set_page_config(page_title="Générateur de QCM par session", layout="centered")
st.title("📄 Générateur de QCM par session (figé ou aléatoire)")

# Fonction de remplacement des balises
def remplacer_placeholders(paragraph, replacements):
    for key, val in replacements.items():
        if key in paragraph.text:
            for run in paragraph.runs:
                if key in run.text:
                    run.text = run.text.replace(key, val)

# Fonction pour itérer sur tous les paragraphes (y compris ceux dans les tableaux)
def iter_all_paragraphs(doc):
    for para in doc.paragraphs:
        yield para
    for table in doc.tables:
        for row in table.rows:
            for cell in row.cells:
                for para in cell.paragraphs:
                    yield para

# Définition des réponses positives pour chaque groupe
POSITIVE_OPTIONS = {
    "satisfaction": ["Très satisfait", "Satisfait"],
    "motivation": ["Très motivés", "Motivés"],
    "assiduite": ["Très motivés", "Motivés"],
    "homogeneite": ["Oui"],
    "questions": ["Toutes les questions", "A peu près toutes"],
    "adaptation": ["Non"],            # on ne proposera plus que "Non"
    "suivi": ["Non concerné"]         # on ne proposera plus que "Non concerné"
}

# Détection des blocs de checkbox
CHECKBOX_GROUPS = {
    "satisfaction": ["Très satisfait", "Satisfait", "Moyennement satisfait", "Insatisfait", "Non satisfait"],
    "motivation": ["Très motivés", "Motivés", "Pas motivés"],
    "assiduite": ["Très motivés", "Motivés", "Pas motivés"],
    "homogeneite": ["Oui", "Non"],
    "questions": ["Toutes les questions", "A peu près toutes", "Il y a quelques sujets sur lesquels je n'avais pas les réponses", "Je n'ai pas pu répondre à la majorité des questions"],
    "adaptation": ["Oui", "Non"],            # le template peut rester inchangé, on cochera toujours "Non"
    "suivi": ["Oui", "Non", "Non concerné"]   # le template peut rester inchangé, on cochera toujours "Non concerné"
}

# Fonction pour appliquer la logique conditionnelle (reste inchangée)
def appliquer_logique_conditionnelle(reponses_figees):
    # Si adaptation n'est pas défini, le mettre à "Non" par défaut
    if "adaptation" not in reponses_figees:
        reponses_figees["adaptation"] = "Non"
    # Si adaptation est "Non", forcer suivi à "Non concerné"
    if reponses_figees["adaptation"] == "Non":
        reponses_figees["suivi"] = "Non concerné"
    return reponses_figees

# Étape 1 : Importer les fichiers
with st.expander("Etape 1 : Importer les fichiers", expanded=True):
    excel_file = st.file_uploader("Fichier Excel des participants", type="xlsx")
    word_file = st.file_uploader("Modèle Word du compte rendu", type="docx")

# Traitement
if excel_file and word_file:
    df = pd.read_excel(excel_file)
    df.columns = df.columns.str.strip()

    # Vérification des colonnes obligatoires
    required_columns = ["session", "formateur", "formation", "nb d'heure", "Nom", "Prénom"]
    if not set(required_columns).issubset(df.columns):
        st.error(f"Colonnes manquantes dans le fichier Excel. Colonnes requises : {required_columns}")
        st.info(f"Colonnes disponibles : {list(df.columns)}")
    else:
        sessions = df.groupby("session")
        reponses_figees = {}

        st.markdown("### Etape 2 : Figer systématiquement adaptation & suivi")

        # On fige directement adaptation et suivi sans interface
        reponses_figees["adaptation"] = "Non"
        reponses_figees["suivi"] = "Non concerné"

        # On garde les autres questions avec possibilité de figer ou non
        st.subheader("Autres questions")
        autres_groupes = [g for g in CHECKBOX_GROUPS.keys() if g not in ["adaptation", "suivi"]]
        for groupe in autres_groupes:
            figer = st.checkbox(f"Figer la réponse pour : {groupe}", key=f"figer_{groupe}")
            if figer:
                choix = st.selectbox(f"Choix figé pour {groupe}", CHECKBOX_GROUPS[groupe], key=f"choix_{groupe}")
                reponses_figees[groupe] = choix

        pistes = st.text_area("Avis & pistes d'amélioration :", key="pistes")
        observations = st.text_area("Autres observations :", key="obs")

        if st.button("🚀 Générer les comptes rendus"):
            # Appliquer la logique conditionnelle (même si adaptation/suivi sont déjà définis)
            reponses_figees = appliquer_logique_conditionnelle(reponses_figees)

            with tempfile.TemporaryDirectory() as tmpdir:
                zip_path = os.path.join(tmpdir, "QCM_Sessions.zip")
                with ZipFile(zip_path, 'w') as zipf:
                    for session_id, participants in sessions:
                        doc = Document(word_file)
                        first = participants.iloc[0]

                        replacements = {
                            "{{nom}}": str(first["Nom"]),
                            "{{prénom}}": str(first["Prénom"]),
                            "{{formateur}}": f"{first['Prénom']} {first['Nom']}",
                            "{{ref_session}}": str(session_id),
                            "{{formation_dispensee}}": str(first["formation"]),
                            "{{duree_formation}}": str(first["nb d'heure"]),
                            "{{nb_participants}}": str(len(participants))
                        }

                        # Remplacement des placeholders dans tout le document
                        for para in iter_all_paragraphs(doc):
                            remplacer_placeholders(para, replacements)

                        # Collecte des paragraphes contenant le placeholder "{{checkbox}}"
                        checkbox_paras = []
                        for para in iter_all_paragraphs(doc):
                            if "{{checkbox}}" in para.text:
                                texte = re.sub(r'\s+', ' ', para.text).strip()
                                for groupe, options in CHECKBOX_GROUPS.items():
                                    for opt in options:
                                        if re.search(rf"\b{re.escape(opt)}\b", texte):
                                            checkbox_paras.append((groupe, opt, para))
                                            break
                                    else:
                                        continue
                                    break

                        # Grouper les paragraphes par groupe
                        group_to_paras = defaultdict(list)
                        for groupe, opt, para in checkbox_paras:
                            group_to_paras[groupe].append((opt, para))

                        # Traitement des réponses : cocher (☑) ou décocher (☐)
                        for groupe, paras in group_to_paras.items():
                            options_presentes = [opt for opt, _ in paras]

                            # Déterminer l'option à cocher
                            if groupe in reponses_figees:
                                option_choisie = reponses_figees[groupe]
                            else:
                                positives_disponibles = [
                                    opt for opt in options_presentes
                                    if groupe in POSITIVE_OPTIONS and opt in POSITIVE_OPTIONS[groupe]
                                ]
                                if positives_disponibles:
                                    option_choisie = random.choice(positives_disponibles)
                                else:
                                    option_choisie = random.choice(options_presentes) if options_presentes else None

                            # Appliquer le choix
                            if option_choisie:
                                for opt, para in paras:
                                    for run in para.runs:
                                        if "{{checkbox}}" in run.text:
                                            run.text = run.text.replace(
                                                "{{checkbox}}",
                                                "☑" if opt == option_choisie else "☐"
                                            )

                        # Ajout des sections "Avis & pistes d'amélioration" et "Autres observations"
                        if pistes.strip():
                            doc.add_paragraph("\nAvis & pistes d'amélioration :\n" + pistes)
                        if observations.strip():
                            doc.add_paragraph("\nAutres observations :\n" + observations)

                        # Enregistrement du document pour chaque session
                        filename = f"Compte_Rendu_{session_id}.docx"
                        path = os.path.join(tmpdir, filename)
                        doc.save(path)
                        zipf.write(path, arcname=filename)

                # Téléchargement de l'archive ZIP
                with open(zip_path, "rb") as f:
                    st.success("Comptes rendus générés avec succès !")
                    st.download_button(
                        "📅 Télécharger l'archive ZIP",
                        data=f,
                        file_name="QCM_Sessions.zip",
                        mime="application/zip"
                    )
