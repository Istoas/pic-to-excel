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

# --- CONFIGURATION API ---
try:
    api_key = st.secrets["GEMINI_API_KEY"]
    genai.configure(api_key=api_key)
except FileNotFoundError:
    st.error("Erreur : La clé API n'est pas configurée dans les Secrets.")
    st.stop()


# --- SESSION STATE ---
if 'df_result' not in st.session_state:
    st.session_state.df_result = None
if 'chat_history' not in st.session_state:
    st.session_state.chat_history = []
if 'current_image' not in st.session_state:
    st.session_state.current_image = None

# --- DESIGN PAGE ---
st.set_page_config(page_title="Pic to Excel", layout="wide", initial_sidebar_state="expanded")

st.markdown("""
<style>
    @import url('https://fonts.googleapis.com/css2?family=Inter:wght@400;600;700&display=swap');
    html, body, [class*="css"]  { font-family: 'Inter', sans-serif; }
    h1 { text-align: center; font-weight: 700; padding-bottom: 20px; }
    .subtitle { text-align: center; opacity: 0.7; margin-top: -20px; margin-bottom: 40px; }
    .stButton > button { width: 100%; border-radius: 6px; height: 3em; font-weight: 600; }
    #MainMenu {visibility: hidden;} footer {visibility: hidden;}
    [data-testid="stFileUploader"] { border: 1px dashed opacity: 0.5; border-radius: 6px; }
</style>
""", unsafe_allow_html=True)

# --- FONCTIONS UTILITAIRES ---

def clean_dataframe_columns(df):
    """Nettoie les colonnes AVANT l'affichage pour éviter le crash Streamlit."""
    new_cols = []
    seen_cols = {}
    
    for col in df.columns:
        # Convertir en string et nettoyer
        col_str = str(col).strip()
        
        # Si vide ou "Unnamed", on donne un nom générique
        if not col_str or "Unnamed" in col_str:
            col_str = "Colonne"
            
        # Gestion des doublons (Ex: Prix, Prix_2, Prix_3)
        if col_str in seen_cols:
            seen_cols[col_str] += 1
            col_str = f"{col_str}_{seen_cols[col_str]}"
        else:
            seen_cols[col_str] = 1
            
        new_cols.append(col_str)
        
    df.columns = new_cols
    return df

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
    
    # Écriture
    for r in dataframe_to_rows(df, index=False, header=True):
        ws.append(r)

    max_row = ws.max_row
    max_col = ws.max_column
    
    if max_row >= 2 and max_col >= 1:
        last_col_letter = get_column_letter(max_col)
        ref = f"A1:{last_col_letter}{max_row}"
        tab = Table(displayName="TableauDonnees", ref=ref)
        style = TableStyleInfo(name="TableStyleMedium2", showRowStripes=True)
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
    st.markdown("### Configuration")
    choix_modele = st.radio("Moteur d'analyse :", ["Standard", "Avancé"], index=1)
    MODEL_NAME = "gemini-2.5-flash" if "Standard" in choix_modele else "gemini-2.5-pro"
    
    st.divider()
    try:
        model = genai.GenerativeModel(MODEL_NAME)
        st.caption(f"Status : Connecté")
    except:
        st.error("Erreur API : Vérifiez les secrets.")
        
    if st.button("Réinitialiser l'application"):
        st.session_state.df_result = None
        st.session_state.chat_history = []
        st.session_state.current_image = None
        st.rerun()

# --- HEADER ---
st.title("Pic to Excel")
st.markdown('<p class="subtitle">Convertisseur universel (Documents & Images)</p>', unsafe_allow_html=True)

col_gauche, col_droite = st.columns([1, 2], gap="large")

# --- GAUCHE : IMPORTATION ---
with col_gauche:
    st.markdown("##### 1. Document Source")
    uploaded_file = st.file_uploader("Format : PDF, PNG, JPG", type=["pdf", "jpg", "png", "jpeg", "webp"])
    
    if uploaded_file:
        try:
            if st.session_state.current_image is None:
                st.session_state.current_image = load_image_from_upload(uploaded_file)
            
            st.image(st.session_state.current_image, caption="Aperçu", use_container_width=True)
            
            st.write("") 
            if st.button("Lancer l'extraction", type="primary", use_container_width=True):
                with st.spinner("Analyse en cours..."):
                    
                    prompt = """
                    Tu es un expert Data & Vision.
                    TACHE : Transforme cette image en données structurées CSV (séparateur point-virgule).
                    
                    RÈGLES :
                    - Uniquement du CSV brut.
                    - Pas de texte avant/après.
                    - Si une case est vide, laisse ;;
                    """
                    
                    response = model.generate_content([prompt, st.session_state.current_image])
                    raw_csv = response.text.replace("```csv", "").replace("```", "").strip()
                    
                    # Lecture
                    df_temp = pd.read_csv(io.StringIO(raw_csv), sep=";", engine="python", on_bad_lines='skip')
                    
                    # NETTOYAGE IMMÉDIAT (Le Fix)
                    st.session_state.df_result = clean_dataframe_columns(df_temp)
                    
                    st.session_state.chat_history.append({"role": "assistant", "content": "Analyse terminée."})
                    st.rerun()
                    
        except Exception as e:
            st.error(f"Erreur technique : {e}")

# --- DROITE : RÉSULTAT & CHAT ---
with col_droite:
    if st.session_state.df_result is not None:
        st.markdown("##### 2. Données Extraites")
        
        # TABLEAU (C'est ici que ça plantait)
        try:
            edited_df = st.data_editor(
                st.session_state.df_result, 
                num_rows="dynamic", 
                use_container_width=True, 
                height=400, 
                key="editor"
            )
            
            # EXPORT
            excel_data = create_styled_excel(edited_df)
            c1, c2 = st.columns(2)
            c1.download_button("Télécharger Excel (.xlsx)", excel_data, "Export.xlsx", "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet", type="primary")
            c2.download_button("Télécharger CSV (.csv)", edited_df.to_csv(index=False, sep=";").encode('utf-8'), "data.csv", "text/csv")
            
        except Exception as e:
            st.error(f"Erreur d'affichage du tableau : {e}")
            st.warning("Essayez de réinitialiser ou de recharger l'image.")
        
        st.divider()
        st.markdown("##### Assistant IA")
        
        # HISTORIQUE CHAT
        for msg in st.session_state.chat_history:
            with st.chat_message(msg["role"]):
                st.write(msg["content"])
        
        # INPUT
        if user_input := st.chat_input("Ex: Ajoute une colonne 'Total'..."):
            st.session_state.chat_history.append({"role": "user", "content": user_input})
            with st.chat_message("user"):
                st.write(user_input)
            
            csv_current = edited_df.to_csv(index=False, sep=";")
            
            prompt_modif = f"""
            CSV actuel :
            {csv_current}
            
            Modif : "{user_input}"
            
            Renvoie le NOUVEAU CSV complet. Rien d'autre.
            """
            
            with st.spinner("Traitement..."):
                try:
                    response = model.generate_content([prompt_modif, st.session_state.current_image])
                    new_csv = response.text.replace("```csv", "").replace("```", "").strip()
                    
                    df_temp = pd.read_csv(io.StringIO(new_csv), sep=";", engine="python", on_bad_lines='skip')
                    
                    # NETTOYAGE AUSSI APRÈS MODIF
                    st.session_state.df_result = clean_dataframe_columns(df_temp)
                    
                    st.session_state.chat_history.append({"role": "assistant", "content": "C'est modifié."})
                    st.rerun()
                    
                except Exception as e:
                    st.error(f"Erreur IA : {e}")

    elif not uploaded_file:
        st.info("Veuillez importer un fichier pour commencer.")