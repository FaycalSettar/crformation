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
    "adaptation": ["Oui"],
    "suivi": ["Oui"]
}

# Détection des blocs de checkbox avec des identifiants plus robustes
CHECKBOX_GROUPS = {
    "satisfaction": ["Très satisfait", "Satisfait", "Moyennement satisfait", "Insatisfait", "Non satisfait"],
    "motivation": ["Très motivés", "Motivés", "Pas motivés"],
    "assiduite": ["Très motivés", "Motivés", "Pas motivés"],
    "homogeneite": ["Oui", "Non"],
    "questions": ["Toutes les questions", "A peu près toutes", "Il y a quelques sujets sur lesquels je n'avais pas les réponses", "Je n'ai pas pu répondre à la majorité des questions"],
    "adaptation": ["Oui", "Non"],  # Groupe pour la question d'adaptation
    "suivi": ["Oui", "Non", "Non concerné"]  # Groupe pour la question de suivi
}

# Textes clés pour les questions spécifiques
QUESTION_ADAPTATION = "Avez-vous effectué une quelconque adaptation du déroulé"
QUESTION_SUIVI = "Si oui, avez-vous pensé à tenir à jour le fichier de suivi"

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

        st.markdown("### Etape 2 : Choisir les réponses à figer (facultatif)")
        groupes_a_exclure = ["adaptation", "suivi"]  # Groupes à exclure de la sélection
        
        for groupe, options in CHECKBOX_GROUPS.items():
            if groupe in groupes_a_exclure:
                continue  # Sauter les groupes exclus
            
            figer = st.checkbox(f"Figer la réponse pour : {groupe}", key=f"figer_{groupe}")
            if figer:
                choix = st.selectbox(f"Choix figé pour {groupe}", options, key=f"choix_{groupe}")
                reponses_figees[groupe] = choix

        # Ajout des réponses figées pour les questions spécifiques
        reponses_figees["adaptation"] = "Non"  # Toujours figé à "Non"
        reponses_figees["suivi"] = "Non concerné"  # Toujours figé à "Non concerné"

        st.info("**Questions systématiquement figées :**")
        st.markdown("- Avez-vous effectué une quelconque adaptation : **Non**")
        st.markdown("- Mise à jour du fichier de suivi : **Non concerné**")

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

                        # Collecte de tous les paragraphes dans une liste pour un accès séquentiel
                        all_paragraphs = list(iter_all_paragraphs(doc))
                        
                        # Créer un mapping des paragraphes aux groupes pour les questions spécifiques
                        para_to_group = {}
                        current_group = None
                        
                        for para in all_paragraphs:
                            text = para.text.lower()
                            
                            # Détection des questions spécifiques
                            if QUESTION_ADAPTATION.lower() in text:
                                current_group = "adaptation"
                            elif QUESTION_SUIVI.lower() in text:
                                current_group = "suivi"
                            elif "{{checkbox}}" in para.text:
                                # Si nous sommes dans un groupe spécifique, assigner le groupe
                                if current_group:
                                    para_to_group[para] = current_group
                        
                        # Collecte des paragraphes contenant le placeholder "{{checkbox}}"
                        checkbox_paras = []
                        for para in all_paragraphs:
                            if "{{checkbox}}" in para.text:
                                # Vérifier si nous avons déjà associé ce paragraphe à un groupe spécifique
                                group = para_to_group.get(para)
                                
                                if not group:
                                    # Méthode de secours : détection par les options
                                    texte = re.sub(r'\s+', ' ', para.text).strip()
                                    for groupe, options in CHECKBOX_GROUPS.items():
                                        for opt in options:
                                            if re.search(rf"\b{re.escape(opt)}\b", texte):
                                                group = groupe
                                                break
                                        if group:
                                            break
                                
                                if group:
                                    checkbox_paras.append((group, para))

                        # Grouper les paragraphes par groupe
                        group_to_paras = defaultdict(list)
                        for group, para in checkbox_paras:
                            group_to_paras[group].append(para)

                        # Traitement des réponses : cocher (☑) ou décocher (☐)
                        for groupe, paras in group_to_paras.items():
                            # Déterminer l'option à cocher
                            if groupe in reponses_figees:
                                option_choisie = reponses_figees[groupe]
                            else:
                                # Choix aléatoire parmi les options positives
                                positives_disponibles = [
                                    opt for opt in CHECKBOX_GROUPS[groupe]
                                    if groupe in POSITIVE_OPTIONS and opt in POSITIVE_OPTIONS[groupe]
                                ]
                                if positives_disponibles:
                                    option_choisie = random.choice(positives_disponibles)
                                else:
                                    option_choisie = random.choice(CHECKBOX_GROUPS[groupe]) if CHECKBOX_GROUPS[groupe] else None

                            # Appliquer le choix
                            if option_choisie:
                                for para in paras:
                                    # Récupérer toutes les options de ce groupe
                                    options_presentes = []
                                    for run in para.runs:
                                        if "{{checkbox}}" in run.text:
                                            # Extraire le texte de l'option
                                            option_text = run.text.replace("{{checkbox}}", "").strip()
                                            options_presentes.append(option_text)
                                    
                                    # Appliquer le choix
                                    for run in para.runs:
                                        if "{{checkbox}}" in run.text:
                                            option_text = run.text.replace("{{checkbox}}", "").strip()
                                            # Remplacer "{{checkbox}}" par le symbole adéquat
                                            run.text = run.text.replace(
                                                "{{checkbox}}",
                                                "☑" if option_text == option_choisie else "☐"
                                            )

                        # Ajout des sections "Avis & pistes d'amélioration" et "Autres observations"
                        doc.add_paragraph("\nAvis & pistes d'amélioration :\n" + pistes)
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
