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

# Fonction pour itérer sur tous les paragraphes
def iter_all_paragraphs(doc):
    for para in doc.paragraphs:
        yield para
    for table in doc.tables:
        for row in table.rows:
            for cell in row.cells:
                for para in cell.paragraphs:
                    yield para

# Définition des réponses positives pour chaque groupe (CORRIGÉ selon le template)
POSITIVE_OPTIONS = {
    "deroulement": ["Satisfait"],  # Premier groupe satisfaction (Déroulé de la formation)
    "motivation": ["Très motivés", "Motivés"],
    "assiduite": ["Très motivés", "Motivés"],  # Assidus utilise les mêmes options que motivés
    "homogeneite": ["Oui"],
    "questions": ["Toutes les questions", "A peu près toutes"],
    "adaptation": ["Oui"],
    "suivi": ["Oui"],
    "satisfaction_globale": ["Très satisfait", "Satisfait"]  # Deuxième groupe satisfaction (niveau global)
}

# Détection des blocs de checkbox (CORRIGÉ selon le template)
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

# Traitement
if excel_file and word_file:
    df = pd.read_excel(excel_file)
    df.columns = df.columns.str.strip()

    # Colonnes requises mises à jour
    required_columns = ["session", "nom", "prénom", "formation", "nb d'heure", "Nom", "Prénom"]
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
                        
                        # Replacements corrigés selon le template
                        replacements = {
                            "{{nom}}": str(first["nom"]),
                            "{{prénom}}": str(first["prénom"]),
                            "{{ref_session}}": str(session_id),
                            "{{formation_dispensee}}": str(first["formation"]),
                            "{{duree_formation}}": str(first["nb d'heure"]),
                            "{{nb_participants}}": str(len(participants))
                        }
                       
                        # Remplacement des placeholders
                        for para in iter_all_paragraphs(doc):
                            remplacer_placeholders(para, replacements)

                        # Collecte des paragraphes avec checkbox
                        checkbox_paras = []
                        satisfaction_count = 0  # Pour différencier les deux groupes satisfaction
                        
                        for para in iter_all_paragraphs(doc):
                            if "{{checkbox}}" in para.text:
                                texte = re.sub(r'\s+', ' ', para.text).strip()
                                
                                # Logique spéciale pour distinguer les deux groupes de satisfaction
                                if any(opt in texte for opt in ["Satisfait", "Moyennement satisfait", "Non satisfait"]):
                                    if satisfaction_count == 0:
                                        groupe_nom = "deroulement"
                                        satisfaction_count += 1
                                    else:
                                        groupe_nom = "satisfaction_globale"
                                    
                                    for opt in CHECKBOX_GROUPS[groupe_nom]:
                                        if re.search(rf'\b{re.escape(opt)}\b', texte):
                                            checkbox_paras.append((groupe_nom, opt, para))
                                            break
                                else:
                                    # Pour les autres groupes
                                    for groupe, options in CHECKBOX_GROUPS.items():
                                        if groupe in ["deroulement", "satisfaction_globale"]:
                                            continue  # Déjà traité ci-dessus
                                        for opt in options:
                                            if re.search(rf'\b{re.escape(opt)}\b', texte):
                                                checkbox_paras.append((groupe, opt, para))
                                                break
                                        else:
                                            continue
                                        break

                        # Grouper les paragraphes par groupe
                        group_to_paras = defaultdict(list)
                        for groupe, opt, para in checkbox_paras:
                            group_to_paras[groupe].append((opt, para))

                        # Traitement des réponses
                        for groupe, paras in group_to_paras.items():
                            options_presentes = [opt for opt, _ in paras]
                           
                            # Déterminer l'option à cocher
                            if groupe in reponses_figees:
                                option_choisie = reponses_figees[groupe]
                            else:
                                # Sélection aléatoire uniquement parmi les réponses positives
                                positives_disponibles = [
                                    opt for opt in options_presentes
                                    if groupe in POSITIVE_OPTIONS and opt in POSITIVE_OPTIONS[groupe]
                                ]
                               
                                # Si des positives sont disponibles, choisir aléatoirement parmi elles
                                if positives_disponibles:
                                    option_choisie = random.choice(positives_disponibles)
                                else:
                                    # Si pas de positives disponibles, choisir aléatoirement parmi toutes
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

                        # Les sections pistes et observations sont déjà dans le template
                        # Pas besoin de les ajouter

                        # Enregistrement
                        filename = f"Compte_Rendu_{session_id}.docx"
                        path = os.path.join(tmpdir, filename)
                        doc.save(path)
                        zipf.write(path, arcname=filename)

                # Téléchargement
                with open(zip_path, "rb") as f:
                    st.success("Comptes rendus générés avec succès !")
                    st.download_button(
                        "📅 Télécharger l'archive ZIP",
                        data=f,
                        file_name="QCM_Sessions.zip",
                        mime="application/zip"
                    )
