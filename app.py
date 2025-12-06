import streamlit as st
import google.generativeai as genai
from PIL import Image
import pandas as pd
import io
import fitz  # PyMuPDF
from openpyxl import Workbook
from openpyxl.utils.dataframe import dataframe_to_rows
from openpyxl.worksheet.table import Table, TableStyleInfo
from openpyxl.utils import get_column_letter
from openpyxl.styles import Alignment

# --- CONFIGURATION ---
try:
    api_key = st.secrets["GEMINI_API_KEY"]
    genai.configure(api_key=api_key)
except FileNotFoundError:
    st.error("Erreur : La clé API n'est pas configurée dans les Secrets.")
    st.stop()

# --- SESSION STATE (MÉMOIRE) ---
if 'df_result' not in st.session_state:
    st.session_state.df_result = None
if 'chat_history' not in st.session_state:
    st.session_state.chat_history = []
if 'current_image' not in st.session_state:
    st.session_state.current_image = None

# --- DESIGN ---
st.set_page_config(page_title="Pic to Excel AI", page_icon="🤖", layout="wide")

st.markdown("""
<style>
    @import url('https://fonts.googleapis.com/css2?family=Inter:wght@400;600;700&display=swap');
    html, body, [class*="css"]  { font-family: 'Inter', sans-serif; }
    .stButton > button { width: 100%; background-color: #0F172A; color: white; border-radius: 8px; }
    #MainMenu {visibility: hidden;} footer {visibility: hidden;}
    
    /* Zone de Chat plus propre */
    .stChatMessage { background-color: #F1F5F9; border-radius: 10px; padding: 10px; }
</style>
""", unsafe_allow_html=True)

# --- FONCTIONS ---
def load_image_from_upload(uploaded_file):
    if uploaded_file.type == "application/pdf":
        doc = fitz.open(stream=uploaded_file.read(), filetype="pdf")
        page = doc.load_page(0)
        pix = page.get_pixmap(dpi=300)
        img = Image.frombytes("RGB", [pix.width, pix.height], pix.samples)
        return img
    else:
        return Image.open(uploaded_file)

def create_styled_excel(df):
    buffer = io.BytesIO()
    wb = Workbook()
    ws = wb.active
    ws.title = "Données"
    
    new_columns = [str(col).strip() if "Unnamed" not in str(col) and str(col).strip() != "" else f"Col_{i+1}" for i, col in enumerate(df.columns)]
    df.columns = new_columns

    for r in dataframe_to_rows(df, index=False, header=True):
        ws.append(r)

    max_row = ws.max_row
    max_col = ws.max_column
    
    if max_row >= 2 and max_col >= 1:
        last_col_letter = get_column_letter(max_col)
        ref = f"A1:{last_col_letter}{max_row}"
        tab = Table(displayName="TabDonnees", ref=ref)
        style = TableStyleInfo(name="TableStyleMedium9", showRowStripes=True)
        tab.tableStyleInfo = style
        ws.add_table(tab)

        for col_idx in range(1, max_col + 1):
            col_letter = get_column_letter(col_idx)
            ws.column_dimensions[col_letter].width = 20
            for cell in ws[col_letter]:
                cell.alignment = Alignment(wrap_text=True, vertical='center')

    wb.save(buffer)
    return buffer.getvalue()

# --- SIDEBAR ---
with st.sidebar:
    st.header("🧠 Cerveau IA")
    choix_modele = st.radio("Modèle :", ["Flash ⚡", "Pro 🎯"], index=1)
    MODEL_NAME = "gemini-2.5-flash" if "Flash" in choix_modele else "gemini-2.5-pro"
    
    try:
        model = genai.GenerativeModel(MODEL_NAME)
        st.success(f"Connecté : {MODEL_NAME}")
    except:
        st.error("Erreur API")
        
    if st.button("🗑️ Reset tout"):
        st.session_state.df_result = None
        st.session_state.chat_history = []
        st.session_state.current_image = None
        st.rerun()

# --- HEADER ---
st.title("🤖 Pic-to-Excel : Assistant Intelligent")
st.markdown("Extrais des données et **parle avec l'IA** pour corriger ou modifier le tableau.")

col_gauche, col_droite = st.columns([1, 2], gap="large")

# --- GAUCHE : IMPORTATION ---
with col_gauche:
    st.subheader("1. Document")
    uploaded_file = st.file_uploader("PDF ou Image", type=["pdf", "jpg", "png", "jpeg", "webp"])
    
    if uploaded_file:
        try:
            # On charge l'image une seule fois
            if st.session_state.current_image is None:
                st.session_state.current_image = load_image_from_upload(uploaded_file)
            
            st.image(st.session_state.current_image, caption="Source", use_container_width=True)
            
            if st.button("⚡ Extraire (Premier Jet)", type="primary"):
                with st.spinner("Analyse initiale..."):
                    prompt = """
                    Tu es un expert Data. Extrais ce tableau en CSV (séparateur point-virgule).
                    Règles : Précision 100%, pas de blabla, respecte la structure.
                    """
                    response = model.generate_content([prompt, st.session_state.current_image])
                    raw_csv = response.text.replace("```csv", "").replace("```", "").strip()
                    
                    st.session_state.df_result = pd.read_csv(io.StringIO(raw_csv), sep=";", engine="python", on_bad_lines='skip')
                    st.session_state.chat_history.append({"role": "assistant", "content": "J'ai extrait le tableau. Tu peux me demander des modifications ci-dessous !"})
                    st.rerun() # On recharge pour afficher le tableau à droite
                    
        except Exception as e:
            st.error(f"Erreur : {e}")

# --- DROITE : RÉSULTAT & CHAT ---
with col_droite:
    if st.session_state.df_result is not None:
        st.subheader("2. Tableau Actuel")
        
        # TABLEAU ÉDITABLE
        edited_df = st.data_editor(st.session_state.df_result, num_rows="dynamic", use_container_width=True, height=400, key="editor")
        
        # EXPORT
        excel_data = create_styled_excel(edited_df)
        c1, c2 = st.columns(2)
        c1.download_button("📥 Excel (.xlsx)", excel_data, "Export.xlsx", "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet", type="primary")
        c2.download_button("📄 CSV", edited_df.to_csv(index=False, sep=";").encode('utf-8'), "data.csv", "text/csv")
        
        st.divider()
        st.subheader("💬 Modifier avec l'IA")
        
        # Historique du chat
        for msg in st.session_state.chat_history:
            with st.chat_message(msg["role"]):
                st.write(msg["content"])
        
        # INPUT CHAT
        if user_input := st.chat_input("Ex: 'Supprime la colonne Prix', 'Mets les noms en majuscule'..."):
            # 1. On affiche le message utilisateur
            st.session_state.chat_history.append({"role": "user", "content": user_input})
            with st.chat_message("user"):
                st.write(user_input)
            
            # 2. On prépare le contexte pour l'IA
            # On lui donne le CSV ACTUEL (format texte) + l'image d'origine + la demande
            csv_current = edited_df.to_csv(index=False, sep=";")
            
            prompt_modif = f"""
            Tu es un assistant Data.
            Voici le tableau CSV actuel :
            {csv_current}
            
            L'utilisateur veut cette modification : "{user_input}"
            
            TACHE : Renvoie le NOUVEAU CSV complet modifié.
            RÈGLES :
            1. Renvoie UNIQUEMENT le CSV (point-virgule).
            2. Pas de texte avant/après.
            3. Si la demande est impossible, renvoie le CSV tel quel.
            """
            
            with st.spinner("L'IA modifie le tableau..."):
                try:
                    # On envoie (Prompt + Image pour contexte visuel si besoin)
                    response = model.generate_content([prompt_modif, st.session_state.current_image])
                    new_csv = response.text.replace("```csv", "").replace("```", "").strip()
                    
                    # Mise à jour du DataFrame
                    st.session_state.df_result = pd.read_csv(io.StringIO(new_csv), sep=";", engine="python", on_bad_lines='skip')
                    
                    # Réponse de l'IA
                    bot_reply = "C'est fait ! Le tableau a été mis à jour."
                    st.session_state.chat_history.append({"role": "assistant", "content": bot_reply})
                    
                    st.rerun() # On recharge pour montrer le nouveau tableau
                    
                except Exception as e:
                    st.error(f"Erreur IA : {e}")

    elif not uploaded_file:
        st.info("👈 Charge un fichier pour commencer.")