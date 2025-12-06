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
from openpyxl.styles import Alignment # Pour gérer le retour à la ligne

# --- CONFIGURATION API ---
try:
    api_key = st.secrets["GEMINI_API_KEY"]
    genai.configure(api_key=api_key)
except FileNotFoundError:
    st.error("Erreur : La clé API n'est pas configurée dans les Secrets.")
    st.stop()

# --- CONFIGURATION PAGE ---
st.set_page_config(page_title="Pic to Excel", page_icon="📊", layout="wide")

# --- CSS ---
st.markdown("""
<style>
    @import url('https://fonts.googleapis.com/css2?family=Inter:wght@400;600;700&display=swap');
    html, body, [class*="css"]  { font-family: 'Inter', sans-serif; }
    h1 { text-align: center; font-weight: 700; color: #0F172A; padding-bottom: 20px; }
    .subtitle { text-align: center; font-size: 1.1rem; color: #64748B; margin-top: -20px; margin-bottom: 40px; }
    .stButton > button { width: 100%; background-color: #0F172A; color: white; border-radius: 8px; height: 3em; font-weight: 600; border: none; }
    .stButton > button:hover { background-color: #334155; color: white; }
    #MainMenu {visibility: hidden;} footer {visibility: hidden;}
    [data-testid="stFileUploader"] { background-color: #F8FAFC; border: 1px dashed #CBD5E1; border-radius: 8px; }
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
    ws.title = "Export"
    
    # Nettoyage Colonnes
    new_columns = []
    for i, col in enumerate(df.columns):
        col_str = str(col).strip()
        if "Unnamed" in col_str or col_str == "":
            new_columns.append(f"Col_{i+1}")
        else:
            new_columns.append(col_str)
    df.columns = new_columns

    # Écriture des données
    for r in dataframe_to_rows(df, index=False, header=True):
        ws.append(r)

    # Création Tableau + Ajustement Largeurs
    max_row = ws.max_row
    max_col = ws.max_column
    
    if max_row >= 2 and max_col >= 1:
        last_col_letter = get_column_letter(max_col)
        ref = f"A1:{last_col_letter}{max_row}"
        tab = Table(displayName="TabDonnees", ref=ref)
        style = TableStyleInfo(name="TableStyleMedium9", showRowStripes=True)
        tab.tableStyleInfo = style
        ws.add_table(tab)

        # BOUCLE D'AJUSTEMENT INTELLIGENT
        for col_idx in range(1, max_col + 1):
            col_letter = get_column_letter(col_idx)
            max_len = 0
            
            # On cherche la cellule la plus longue de la colonne
            for cell in ws[col_letter]:
                try:
                    if cell.value:
                        cell_len = len(str(cell.value))
                        if cell_len > max_len: max_len = cell_len
                except: pass
            
            # Calcul de la largeur idéale
            adjusted_width = (max_len + 2) * 1.2
            
            # SÉCURITÉ : Si c'est trop large (>50 caractères), on bloque et on active le retour à la ligne
            if adjusted_width > 50:
                ws.column_dimensions[col_letter].width = 50
                # On applique le retour à la ligne pour toutes les cellules de cette colonne
                for cell in ws[col_letter]:
                    cell.alignment = Alignment(wrap_text=True, vertical='center')
            else:
                # Sinon on applique la largeur calculée
                ws.column_dimensions[col_letter].width = adjusted_width

    wb.save(buffer)
    return buffer.getvalue()

# --- SIDEBAR ---
with st.sidebar:
    st.markdown("### Configuration")
    choix_modele = st.radio("Moteur :", ["Standard", "Avancé"], index=1)
    MODEL_NAME = "gemini-2.5-flash" if "Standard" in choix_modele else "gemini-2.5-pro"
    st.divider()
    try:
        model = genai.GenerativeModel(MODEL_NAME)
        st.caption(f"Status: Prêt ({MODEL_NAME})")
    except: st.error("Erreur API")

# --- INTERFACE ---
st.title("Pic to Excel")
st.markdown('<p class="subtitle">Convertisseur universel : Images & PDF vers Excel</p>', unsafe_allow_html=True)

col_gauche, col_droite = st.columns([1, 2], gap="large")

if 'df_result' not in st.session_state:
    st.session_state.df_result = None

with col_gauche:
    st.markdown("##### 1. Importation")
    uploaded_file = st.file_uploader("Format : PDF, PNG, JPG", type=["pdf", "jpg", "png", "jpeg", "webp"])
    
    if uploaded_file:
        try:
            image = load_image_from_upload(uploaded_file)
            st.image(image, caption="Aperçu", use_container_width=True)
            st.write("") 
            if st.button("Lancer la conversion", type="primary", use_container_width=True):
                with st.spinner("Analyse du document..."):
                    try:
                        prompt = """
                        Agis comme un expert en saisie de données.
                        ANALYSE : Transforme ce document en données structurées CSV.
                        RÈGLES :
                        1. Respect strict de la structure (lignes/colonnes).
                        2. Si une cellule est vide, laisse-la vide (;;).
                        3. Format numérique précis (virgules/points).
                        4. Sortie : Uniquement le CSV brut (séparateur point-virgule).
                        """
                        response = model.generate_content([prompt, image])
                        raw_csv = response.text.replace("```csv", "").replace("```", "").strip()
                        st.session_state.df_result = pd.read_csv(io.StringIO(raw_csv), sep=";", engine="python", on_bad_lines='skip')
                    except Exception as e: st.error(f"Erreur IA : {e}")
        except Exception as e: st.error(f"Erreur lecture : {e}")

with col_droite:
    if st.session_state.df_result is not None:
        st.markdown("##### 2. Vérification & Export")
        edited_df = st.data_editor(st.session_state.df_result, num_rows="dynamic", use_container_width=True, height=600)
        excel_data = create_styled_excel(edited_df)
        c1, c2 = st.columns([1, 1])
        with c1: st.download_button("Télécharger Excel", excel_data, "Export.xlsx", "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet", type="primary")
        with c2: st.download_button("Télécharger CSV", edited_df.to_csv(index=False, sep=";").encode('utf-8'), "export.csv", "text/csv")
    elif not uploaded_file:
        st.info("Importez un fichier PDF ou Image pour commencer.")