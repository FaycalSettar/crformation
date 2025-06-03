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
    "deroulement": ["Satisfait"],                   # "Déroulé de la formation"
    "motivation": ["Très motivés", "Motivés"],
    "assiduite": ["Très motivés", "Motivés"],       # même options que motivation
    "homogeneite": ["Oui"],
    "questions": ["Toutes les questions", "A peu près toutes"],
    "adaptation": ["Oui"],
    "suivi": ["Oui"],
    "satisfaction_globale": ["Très satisfait", "Satisfait"]  # "Niveau global de satisfaction"
}

# Détection des blocs de checkbox
CHECKBOX_GROUPS = {
    "deroulement": ["Satisfait", "Moyennement satisfait", "Non satisfait"],
    "motivation": ["Très motivés", "Motivés", "Pas motivés"],
    "assiduite": ["Très motivés", "Motivés", "Pas motivés"],
    "homogeneite": ["Oui", "Non"],
    "questions": ["Toutes les questions", "A peu près toutes", "Il y a quelques sujets sur lesquels je n'avais pas les réponses", "Je n'ai pas pu répondre à la majorité des questions"],
    "adaptation": ["Oui", "Non"],
    "suivi": ["Oui", "Non", "Non concerné"],
    "satisfaction_globale": ["Très satisfait", "Satisfait", "Moyennement satisfait", "Insatisfait", "Non satisfait"]
}

# Étape 1 : Importer les fichiers
with st.expander("Etape 1 : Importer les fichiers", expanded=True):
    excel_file = st.file_uploader("Fichier Excel des participants", type="xlsx")
    word_file = st.file_uploader("Modèle Word du compte rendu", type="docx")

# Traitement principal
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
                        # Remplacements des balises (ex. {{nom}}, {{ref_session}}, etc.)
                        # Adaptez bien selon vos placeholders dans le document Word
                        replacements = {
                            "{{nom}}": str(first["formateur"]).split()[0] if len(str(first["formateur"]).split()) > 0 else str(first["formateur"]),
                            "{{prénom}}": str(first["formateur"]).split()[1] if len(str(first["formateur"]).split()) > 1 else "",
                            "{{ref_session}}": str(session_id),
                            "{{formation_dispensee}}": str(first["formation"]),
                            "{{duree_formation}}": str(first["nb d'heure"]),
                            "{{nb_participants}}": str(len(participants))
                        }
                        # Parcours de tout le document pour remplacer les balises
                        for para in iter_all_paragraphs(doc):
                            remplacer_placeholders(para, replacements)

                        # ----- DÉTECTION DES CHECKBOX « ☐ » ----- #
                        # On repère chaque paragraphe commençant par le symbole "☐"
                        all_paras = list(iter_all_paragraphs(doc))
                        checkbox_paras = []

                        for idx, para in enumerate(all_paras):
                            texte = para.text.strip()
                            # S’il y a un "☐" en tête, c’est une option de QCM
                            if texte.startswith("☐"):
                                # Récupérer le « texte d’option » sans le symbole
                                option_label = texte.lstrip("☐").strip()

                                # On tente de déterminer le groupe grâce au contexte
                                # On regarde quelques paragraphes précédents pour trouver une phrase-guide
                                context = ""
                                for j in range(max(0, idx - 5), idx):
                                    context += all_paras[j] + " "

                                groupe_nom = None
                                if "Déroulé de la formation" in context:
                                    groupe_nom = "deroulement"
                                elif "niveau global de satisfaction" in context:
                                    groupe_nom = "satisfaction_globale"
                                elif "étaient-ils motivés" in context or "motivés" in context:
                                    groupe_nom = "motivation"
                                elif "assidus" in context:
                                    groupe_nom = "assiduite"
                                elif "formation s’est avérée homogène" in context or "homogène" in context:
                                    groupe_nom = "homogeneite"
                                elif "répondre à toutes les questions" in context or "questions" in context:
                                    groupe_nom = "questions"
                                elif "adaptation du déroulé" in context:
                                    groupe_nom = "adaptation"
                                elif "tenir à jour le fichier" in context or "suivi" in context:
                                    groupe_nom = "suivi"
                                else:
                                    # Fallback : si on ne trouve pas de mot-clé, on essaie de reconnaître
                                    # le groupe par la simple présence d'une des options dans CHECKBOX_GROUPS
                                    for g, opts in CHECKBOX_GROUPS.items():
                                        if option_label in opts:
                                            groupe_nom = g
                                            break

                                if groupe_nom:
                                    checkbox_paras.append((groupe_nom, option_label, para))

                        # Regrouper les paragraphes par groupe
                        group_to_paras = defaultdict(list)
                        for groupe, opt_label, para in checkbox_paras:
                            group_to_paras[groupe].append((opt_label, para))

                        # ----- TRAITEMENT DES RÉPONSES PAR GROUPE ----- #
                        for groupe, paras in group_to_paras.items():
                            options_presentes = [opt_label for opt_label, _ in paras]

                            # Si l’utilisateur a figé ce groupe, on prend directement la valeur figée
                            if groupe in reponses_figees:
                                option_choisie = reponses_figees[groupe]
                            else:
                                # Sinon, on cherche les options « positives » pour ce groupe
                                positives_disponibles = [
                                    opt_label for opt_label in options_presentes
                                    if groupe in POSITIVE_OPTIONS and opt_label in POSITIVE_OPTIONS[groupe]
                                ]
                                if positives_disponibles:
                                    # Choix aléatoire parmi les positives
                                    option_choisie = random.choice(positives_disponibles)
                                else:
                                    # Si pas de positive dispo, on pioche aléatoirement parmi toutes
                                    option_choisie = random.choice(options_presentes) if options_presentes else None

                            # Appliquer le choix : remplacer le symbole "☐" par "☑" sur l'option choisie
                            if option_choisie:
                                for opt_label, para in paras:
                                    texte_actuel = para.text.strip()
                                    # Supprimer d’abord toute case en début de ligne (☐ ou ☑)
                                    bare = re.sub(r'^[☐☑]\s*', '', texte_actuel).strip()
                                    if opt_label == option_choisie:
                                        para.text = f"☑ {bare}"
                                    else:
                                        para.text = f"☐ {bare}"

                        # Les sections « pistes » et « observations » sont déjà présentes dans le template,
                        # donc on ne les rajoute pas ici.

                        # Enregistrement du document pour chaque session
                        filename = f"Compte_Rendu_{session_id}.docx"
                        path = os.path.join(tmpdir, filename)
                        doc.save(path)
                        zipf.write(path, arcname=filename)

                # Téléchargement de l’archive ZIP
                with open(zip_path, "rb") as f:
                    st.success("Comptes rendus générés avec succès !")
                    st.download_button(
                        "📅 Télécharger l’archive ZIP",
                        data=f,
                        file_name="QCM_Sessions.zip",
                        mime="application/zip"
                    )
