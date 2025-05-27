import streamlit as st
import pandas as pd
from docx import Document
import random
import os
import tempfile
from zipfile import ZipFile
import re

st.set_page_config(page_title="Générateur de QCM", layout="centered")
st.title("📄 Générateur de QCM (avec figement manuel des réponses)")

# === Fonctions utilitaires ===
def remplacer_placeholders(paragraph, replacements):
    if not paragraph.text:
        return
    for run in paragraph.runs:
        for key, val in replacements.items():
            if key in run.text:
                run.text = run.text.replace(key, val)

def detecter_questions(doc):
    questions = []
    current_question = None
    pattern = re.compile(r'^(\d+[\. ]*\d*)\s*[-–—)\s.]*\s*(.+?)\?$')
    reponse_pattern = re.compile(r'^([A-D])[\s\-–—).]+\s*(.*?)(\{\{checkbox\}\})?\s*$')
    for i, para in enumerate(doc.paragraphs):
        texte = para.text.strip()
        match_question = pattern.match(texte)
        if match_question:
            question_num = re.sub(r'\s+', '.', match_question.group(1)).strip().rstrip('.')
            current_question = {
                "index": i,
                "texte": f"{question_num} - {match_question.group(2)}?",
                "reponses": [],
                "original_text": texte
            }
            questions.append(current_question)
        elif current_question:
            match_reponse = reponse_pattern.match(texte)
            if match_reponse:
                lettre = match_reponse.group(1)
                texte_rep = match_reponse.group(2).strip()
                current_question["reponses"].append({
                    "index": i,
                    "lettre": lettre,
                    "texte": texte_rep,
                    "original_text": texte
                })
    return [q for q in questions if len(q["reponses"]) >= 2]

# === Import des fichiers ===
with st.expander("📁 Étape 1 : Importation", expanded=True):
    excel_file = st.file_uploader("Fichier Excel des participants", type="xlsx")
    word_file = st.file_uploader("Modèle Word du QCM", type="docx")

# === Lecture Word une seule fois ===
if word_file and ('questions' not in st.session_state or st.session_state.get('current_template') != word_file.name):
    doc = Document(word_file)
    st.session_state.questions = detecter_questions(doc)
    st.session_state.figees = {}
    st.session_state.reponses_figees = {}
    st.session_state.current_template = word_file.name

# === Configuration manuelle ===
if 'questions' in st.session_state:
    st.markdown("### ⚙️ Étape 2 : Choix des réponses figées")

    for q in st.session_state['questions']:
        q_id = q['index']
        q_num = q['texte'].split()[0]
        col1, col2 = st.columns([1, 4])
        with col1:
            figer = st.checkbox(f"Figer Q{q_num}", value=st.session_state.figees.get(q_id, False), key=f"figer_{q_id}")
        with col2:
            if figer:
                options = [f"{r['lettre']} - {r['texte']}" for r in q['reponses']]
                bonne = st.selectbox(f"Bonne réponse pour Q{q_num}", options=options, key=f"bonne_{q_id}")
                st.session_state.figees[q_id] = True
                st.session_state.reponses_figees[q_id] = options.index(bonne)

# === Génération QCM ===
def generer_qcm(doc_model, row):
    doc = Document(doc_model)
    replacements = {
        '{{prenom}}': str(row.get('Prénom', '')),
        '{{nom}}': str(row.get('Nom', '')),
        '{{email}}': str(row.get('Email', ''))
    }
    for para in doc.paragraphs:
        remplacer_placeholders(para, replacements)
    for table in doc.tables:
        for row_cell in table.rows:
            for cell in row_cell.cells:
                for para in cell.paragraphs:
                    remplacer_placeholders(para, replacements)

    for q in st.session_state.questions:
        reponses = q['reponses'].copy()
        is_figee = st.session_state.figees.get(q['index'], False)

        if is_figee:
            bonne_idx = st.session_state.reponses_figees.get(q['index'], 0)
            reponse_correcte = reponses.pop(bonne_idx)
            reponses.insert(0, reponse_correcte)
        else:
            random.shuffle(reponses)

        for idx, rep in enumerate(reponses):
            checkbox = "☑" if idx == 0 else "☐"
            base = rep['original_text'].split(' ', 1)[0]
            nouvelle_ligne = f"{base} - {rep['texte']} {checkbox}"
            doc.paragraphs[rep['index']].text = nouvelle_ligne
    return doc

# === Étape 3 : Génération ===
if excel_file and word_file and 'questions' in st.session_state:
    if st.button("🚀 Générer les QCM"):
        df = pd.read_excel(excel_file)
        df.columns = df.columns.str.strip()
        with tempfile.TemporaryDirectory() as tmpdir:
            zip_path = os.path.join(tmpdir, "QCM_Generes.zip")
            with ZipFile(zip_path, 'w') as zipf:
                for _, row in df.iterrows():
                    doc = generer_qcm(word_file, row)
                    prenom = re.sub(r'\W+', '_', str(row.get('Prénom', '')))
                    nom = re.sub(r'\W+', '_', str(row.get('Nom', '')))
                    filename = f"QCM_{prenom}_{nom}.docx"
                    path = os.path.join(tmpdir, filename)
                    doc.save(path)
                    zipf.write(path, arcname=filename)
            with open(zip_path, "rb") as f:
                st.success("✅ QCM générés avec succès !")
                st.download_button("📦 Télécharger l'archive ZIP", data=f, file_name="QCM_Personnalises.zip", mime="application/zip")
