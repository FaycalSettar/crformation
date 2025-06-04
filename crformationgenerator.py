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

# 1. Fonction de remplacement des champs classiques ({{nom}}, {{prénom}}, etc.)
def remplacer_placeholders(paragraph, replacements):
    for key, val in replacements.items():
        if key in paragraph.text:
            for run in paragraph.runs:
                if key in run.text:
                    run.text = run.text.replace(key, val)

# 2. Itérateur sur tous les paragraphes (même ceux dans les tableaux)
def iter_all_paragraphs(doc):
    for para in doc.paragraphs:
        yield para
    for table in doc.tables:
        for row in table.rows:
            for cell in row.cells:
                for para in cell.paragraphs:
                    yield para

# 3. Options positives (pour tirage aléatoire si non figé)
POSITIVE_OPTIONS = {
    "satisfaction": ["Très satisfait", "Satisfait"],
    "motivation": ["Très motivés", "Motivés"],
    "assiduite": ["Très motivés", "Motivés"],
    "homogeneite": ["Oui"],
    "questions": ["Toutes les questions", "A peu près toutes"],
    # On ne laisse volontairement pas "adaptation" ni "suivi" ici,
    # car on les figera systématiquement plus bas.
}

# 4. Groupes de tous les libellés de cases à cocher (doivent correspondre à votre gabarit Word)
CHECKBOX_GROUPS = {
    "satisfaction": [
        "Très satisfait",
        "Satisfait",
        "Moyennement satisfait",
        "Insatisfait",
        "Non satisfait"
    ],
    "motivation": [
        "Très motivés",
        "Motivés",
        "Pas motivés"
    ],
    "assiduite": [
        "Très motivés",
        "Motivés",
        "Pas motivés"
    ],
    "homogeneite": [
        "Oui",
        "Non"
    ],
    "questions": [
        "Toutes les questions",
        "A peu près toutes",
        "Il y a quelques sujets sur lesquels je n'avais pas les réponses",
        "Je n'ai pas pu répondre à la majorité des questions"
    ],
    "adaptation": [
        "Non",
        "Oui"
    ],
    "suivi": [
        "Non concerné",
        "Non",
        "Oui"
    ]
}

# 5. Logique conditionnelle (au cas où adaptation n'existerait pas dans reponses_figees)
def appliquer_logique_conditionnelle(reponses_figees):
    # S'il manque adaptation, on force à "Non"
    if "adaptation" not in reponses_figees:
        reponses_figees["adaptation"] = "Non"
    # Si adaptation = "Non", on force suivi à "Non concerné"
    if reponses_figees["adaptation"] == "Non":
        reponses_figees["suivi"] = "Non concerné"
    return reponses_figees

# === Début de l'application Streamlit ===

# Étape 1 : import des fichiers Excel et Word
with st.expander("Etape 1 : Importer les fichiers", expanded=True):
    excel_file = st.file_uploader("Fichier Excel des participants", type="xlsx")
    word_file = st.file_uploader("Modèle Word du compte rendu", type="docx")

if excel_file and word_file:
    df = pd.read_excel(excel_file)
    df.columns = df.columns.str.strip()

    # Colonnes obligatoires dans l'Excel
    required_columns = ["session", "formateur", "formation", "nb d'heure", "Nom", "Prénom"]
    if not set(required_columns).issubset(df.columns):
        st.error(f"Colonnes manquantes dans le fichier Excel. Colonnes requises : {required_columns}")
        st.info(f"Colonnes disponibles : {list(df.columns)}")
    else:
        # On regroupe par session
        sessions = df.groupby("session")
        reponses_figees = {}

        # Étape 2 : on ne propose pas de radio pour adaptation/suivi, on fige directement
        st.markdown("### Etape 2 : Adaptation & suivi figés automatiquement")
        # on force ces deux clés
        reponses_figees["adaptation"] = "Non"
        reponses_figees["suivi"] = "Non concerné"

        # On laisse le choix de figer les autres groupes éventuels si besoin
        st.subheader("Autres questions (optionnel)")
        autres_groupes = [g for g in CHECKBOX_GROUPS.keys() if g not in ["adaptation", "suivi"]]
        for groupe in autres_groupes:
            figer = st.checkbox(f"Figer la réponse pour : {groupe}", key=f"figer_{groupe}")
            if figer:
                choix = st.selectbox(f"Choix figé pour {groupe}", CHECKBOX_GROUPS[groupe], key=f"choix_{groupe}")
                reponses_figees[groupe] = choix

        pistes = st.text_area("Avis & pistes d'amélioration :", key="pistes")
        observations = st.text_area("Autres observations :", key="obs")

        if st.button("🚀 Générer les comptes rendus"):
            # Re-validation de la logique pour adaptation / suivi
            reponses_figees = appliquer_logique_conditionnelle(reponses_figees)

            with tempfile.TemporaryDirectory() as tmpdir:
                zip_path = os.path.join(tmpdir, "QCM_Sessions.zip")
                with ZipFile(zip_path, "w") as zipf:
                    for session_id, participants in sessions:
                        doc = Document(word_file)
                        first = participants.iloc[0]

                        # 6. Remplacement des champs classiques dans tout le document
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

                        # 7. Détection des paragraphes contenant {{checkbox}}
                        checkbox_paras = []
                        for para in iter_all_paragraphs(doc):
                            if "{{checkbox}}" in para.text:
                                # On nettoie les espaces redondants
                                texte = re.sub(r"\s+", " ", para.text).strip()
                                for groupe, options in CHECKBOX_GROUPS.items():
                                    for opt in options:
                                        # Si l'option (texte) apparaît dans ce paragraphe
                                        # on considère que c'est du groupe ‘groupe’
                                        if re.search(rf"\b{re.escape(opt)}\b", texte):
                                            checkbox_paras.append((groupe, opt, para))
                                            break
                                    else:
                                        continue
                                    break

                        # 8. Grouper par groupe de question
                        group_to_paras = defaultdict(list)
                        for groupe, opt, para in checkbox_paras:
                            group_to_paras[groupe].append((opt, para))

                        # 9. Parcours des groupes pour cocher ou décocher
                        for groupe, paras in group_to_paras.items():
                            # Liste des libellés effectivement présents dans le doc pour ce groupe
                            options_presentes = [opt for opt, _ in paras]

                            # 9.a. Déterminer l’option à cocher
                            if groupe in reponses_figees:
                                option_choisie = reponses_figees[groupe]
                            else:
                                # Si on n'a pas figé ce groupe, on prend au hasard
                                positives_disponibles = [
                                    opt for opt in options_presentes
                                    if groupe in POSITIVE_OPTIONS and opt in POSITIVE_OPTIONS[groupe]
                                ]
                                if positives_disponibles:
                                    option_choisie = random.choice(positives_disponibles)
                                else:
                                    option_choisie = random.choice(options_presentes)

                            # 9.b. Appliquer le remplacement dans chacun des paragraphes
                            for opt, para in paras:
                                for run in para.runs:
                                    if "{{checkbox}}" in run.text:
                                        # On remplace le texte littéral "{{checkbox}}"
                                        # par "☑" si c'est l'option choisie, sinon "☐"
                                        symbole = "☑" if opt == option_choisie else "☐"
                                        run.text = run.text.replace("{{checkbox}}", symbole)

                        # 10. Ajout des zones “avis & pistes” et “observations”
                        if pistes.strip():
                            doc.add_paragraph("\nAvis & pistes d'amélioration :\n" + pistes)
                        if observations.strip():
                            doc.add_paragraph("\nAutres observations :\n" + observations)

                        # 11. Sauvegarde du fichier Word pour cette session
                        filename = f"Compte_Rendu_{session_id}.docx"
                        path = os.path.join(tmpdir, filename)
                        doc.save(path)
                        zipf.write(path, arcname=filename)

                # 12. Bouton de téléchargement de l’archive ZIP
                with open(zip_path, "rb") as f:
                    st.success("Comptes rendus générés avec succès !")
                    st.download_button(
                        "📅 Télécharger l'archive ZIP",
                        data=f,
                        file_name="QCM_Sessions.zip",
                        mime="application/zip"
                    )
