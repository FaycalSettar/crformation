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

# --------------------------------------------
# Fonctions utilitaires pour le traitement Word
# --------------------------------------------

def remplacer_placeholders(paragraph, replacements):
    """
    Parcourt chaque run d'un paragraphe et remplace les clés du dictionnaire `replacements`
    par leurs valeurs respectives, sans écraser le formatage local (gras, italique, etc.).
    """
    for key, val in replacements.items():
        if key in paragraph.text:
            for run in paragraph.runs:
                if key in run.text:
                    run.text = run.text.replace(key, val)


def iter_all_paragraphs(doc):
    """
    Génère tous les objets Paragraph d'un Document python‐docx, 
    qu'ils soient à la racine ou à l'intérieur de cellules de tableaux.
    """
    # Paragraphes à la racine
    for para in doc.paragraphs:
        yield para

    # Paragraphes à l'intérieur des cellules de chaque tableau
    for table in doc.tables:
        for row in table.rows:
            for cell in row.cells:
                for para in cell.paragraphs:
                    yield para


# -------------------------------
# Définitions des groupes et options
# -------------------------------

# (1) Réponses "positives" pour chaque groupe (prioritaires en mode aléatoire)
POSITIVE_OPTIONS = {
    "satisfaction": ["Très satisfait", "Satisfait"],
    "motivation": ["Très motivés", "Motivés"],
    "assiduite": ["Très motivés", "Motivés"],
    "homogeneite": ["Oui"],
    "questions": ["Toutes les questions", "A peu près toutes"],
    "adaptation": ["Oui"],   # on forcera toutefois "Non" en dur
    "suivi": ["Oui"]
}

# (2) Toutes les options possibles (pour vérification ou affichage)
CHECKBOX_GROUPS = {
    "satisfaction": ["Très satisfait", "Satisfait", "Moyennement satisfait", "Insatisfait", "Non satisfait"],
    "motivation": ["Très motivés", "Motivés", "Pas motivés"],
    "assiduite": ["Très motivés", "Motivés", "Pas motivés"],
    "homogeneite": ["Oui", "Non"],
    "questions": [
        "Toutes les questions",
        "A peu près toutes",
        "Il y a quelques sujets sur lesquels je n'avais pas les réponses",
        "Je n'ai pas pu répondre à la majorité des questions"
    ],
    "adaptation": ["Oui", "Non"],
    "suivi": ["Oui", "Non", "Non concerné"]
}


# -------------------------------
# Étape 1 : Chargement des fichiers
# -------------------------------
with st.expander("Etape 1 : Importer les fichiers", expanded=True):
    excel_file = st.file_uploader("Fichier Excel des participants", type="xlsx")
    word_file = st.file_uploader("Modèle Word du compte rendu", type="docx")


# -------------------------------
# Si les deux fichiers sont fournis, on peut traiter
# -------------------------------
if excel_file and word_file:
    # Lecture du fichier Excel
    df = pd.read_excel(excel_file)
    df.columns = df.columns.str.strip()

    # Vérification des colonnes obligatoires
    required_columns = ["session", "formateur", "formation", "nb d'heure", "Nom", "Prénom"]
    if not set(required_columns).issubset(df.columns):
        st.error(f"Colonnes manquantes dans le fichier Excel. Colonnes requises : {required_columns}")
        st.info(f"Colonnes disponibles : {list(df.columns)}")
    else:
        # On regroupe les participants par "session"
        sessions = df.groupby("session")
        # Dictionnaire pour stocker les réponses "figées" par l'utilisateur
        reponses_figees = {}

        # Étape 2 : Optionnellement, l’utilisateur peut "figer" certaines réponses
        st.markdown("### Etape 2 : Choisir les réponses à figer (facultatif)")
        for groupe, options in CHECKBOX_GROUPS.items():
            figer = st.checkbox(f"Figer la réponse pour : {groupe}", key=f"figer_{groupe}")
            if figer:
                choix = st.selectbox(f"Choix figé pour {groupe}", options, key=f"choix_{groupe}")
                reponses_figees[groupe] = choix

        # Champs libres pour « Avis & pistes d'amélioration » et « Autres observations »
        pistes = st.text_area("Avis & pistes d'amélioration :", key="pistes")
        observations = st.text_area("Autres observations :", key="obs")

        # Bouton pour lancer la génération
        if st.button("🚀 Générer les comptes rendus"):
            with tempfile.TemporaryDirectory() as tmpdir:
                zip_path = os.path.join(tmpdir, "QCM_Sessions.zip")
                with ZipFile(zip_path, 'w') as zipf:
                    # Pour chaque session (clé=session_id, DataFrame=participants)
                    for session_id, participants in sessions:
                        # On recharge le modèle Word pour chaque session
                        doc = Document(word_file)
                        first = participants.iloc[0]

                        # 1) Remplacements simples des placeholders standards
                        replacements = {
                            "{{nom}}": str(first["Nom"]),
                            "{{prénom}}": str(first["Prénom"]),
                            "{{formateur}}": f"{first['Prénom']} {first['Nom']}",
                            "{{ref_session}}": str(session_id),
                            "{{formation_dispensee}}": str(first["formation"]),
                            "{{duree_formation}}": str(first["nb d'heure"]),
                            "{{nb_participants}}": str(len(participants))
                        }
                        for para in iter_all_paragraphs(doc):
                            remplacer_placeholders(para, replacements)

                        # ------------------------------
                        # 2) On détecte et on regroupe les "{{checkbox}}"
                        #      en se basant sur la question située juste au-dessus
                        # ------------------------------
                        group_to_paras = defaultdict(list)

                        # --- On parcourt les paragraphes de la racine (doc.paragraphs) ---
                        for i, para in enumerate(doc.paragraphs):
                            if "{{checkbox}}" in para.text:
                                # On cherche le paragraphe-question le plus proche au-dessus
                                j = i - 1
                                while j >= 0:
                                    txt_j = doc.paragraphs[j].text.strip().lower()
                                    # On ignore les lignes vides ou contenant déjà "{{checkbox}}"
                                    if (not txt_j) or ("{{checkbox}}" in txt_j):
                                        j -= 1
                                    else:
                                        break

                                if j < 0:
                                    # Aucun header-question trouvé
                                    continue

                                header = doc.paragraphs[j].text.strip().lower()

                                # Détermination du groupe selon des mots-clés dans l'en-tête
                                groupe = None
                                # Ordre de priorité pour éviter les ambiguïtés "adaptation" vs "suivi"
                                if "motiv" in header:
                                    groupe = "motivation"
                                elif "assid" in header:
                                    groupe = "assiduite"
                                elif "niveau homogène" in header:
                                    groupe = "homogeneite"
                                elif "suivi" in header:
                                    groupe = "suivi"
                                elif "adaptation" in header:
                                    groupe = "adaptation"
                                elif "questions" in header:
                                    groupe = "questions"
                                elif "satisfaction" in header:
                                    groupe = "satisfaction"
                                # (On peut ajouter d’autres règles si besoin)

                                if groupe is None:
                                    # Si on ne reconnaît pas la question, on ignore cette case
                                    continue

                                # Récupération du libellé d’option (après "{{checkbox}}")
                                opt = para.text.replace("{{checkbox}}", "").strip()
                                group_to_paras[groupe].append((opt, para))

                        # ------------------------------
                        # 3) Pour chaque groupe, on décide quelle case cocher
                        #     en forçant "Non" pour adaptation et "Non concerné" pour suivi
                        # ------------------------------
                        for groupe, paras in group_to_paras.items():
                            options_presentes = [opt for opt, _ in paras]

                            # 3.1) Forcer "Non" pour "adaptation" si l'option existe
                            if groupe == "adaptation" and "Non" in options_presentes:
                                option_choisie = "Non"

                            # 3.2) Forcer "Non concerné" pour "suivi" si l'option existe
                            elif groupe == "suivi" and "Non concerné" in options_presentes:
                                option_choisie = "Non concerné"

                            # 3.3) Sinon, si l’utilisateur a VRAIMENT figé un choix via l’interface
                            elif groupe in reponses_figees:
                                option_choisie = reponses_figees[groupe]

                            # 3.4) Sinon, sélection aléatoire (priorité aux "positives")
                            else:
                                positives = [
                                    opt
                                    for opt in options_presentes
                                    if (groupe in POSITIVE_OPTIONS and opt in POSITIVE_OPTIONS[groupe])
                                ]
                                if positives:
                                    option_choisie = random.choice(positives)
                                else:
                                    option_choisie = random.choice(options_presentes) if options_presentes else None

                            # 3.5) On remplace "{{checkbox}}" par "☑" pour l’option choisie, "☐" sinon
                            if option_choisie:
                                for opt, para in paras:
                                    for run in para.runs:
                                        if "{{checkbox}}" in run.text:
                                            run.text = run.text.replace(
                                                "{{checkbox}}",
                                                "☑" if opt == option_choisie else "☐"
                                            )

                        # ------------------------------
                        # 4) On ajoute enfin les sections libres
                        # ------------------------------
                        doc.add_paragraph("\nAvis & pistes d'amélioration :\n" + pistes)
                        doc.add_paragraph("\nAutres observations :\n" + observations)

                        # ------------------------------
                        # 5) Enregistrement du document final pour cette session
                        # ------------------------------
                        filename = f"Compte_Rendu_{session_id}.docx"
                        path = os.path.join(tmpdir, filename)
                        doc.save(path)
                        zipf.write(path, arcname=filename)

                # ------------------------------
                # 6) Une fois toutes les sessions traitées, on propose le ZIP à télécharger
                # ------------------------------
                with open(zip_path, "rb") as f:
                    st.success("Comptes rendus générés avec succès !")
                    st.download_button(
                        "📅 Télécharger l'archive ZIP",
                        data=f,
                        file_name="QCM_Sessions.zip",
                        mime="application/zip"
                    )
