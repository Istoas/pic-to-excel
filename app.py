import streamlit as st
import google.generativeai as genai
from PIL import Image
import pandas as pd
import io
from openpyxl import Workbook
from openpyxl.utils.dataframe import dataframe_to_rows
from openpyxl.worksheet.table import Table, TableStyleInfo
from openpyxl.utils import get_column_letter

# --- CONFIGURATION ---
# Remplace par ta vraie clé API
genai.configure(api_key="AIzaSyCPj5fuce3c9sS6r7Dd7LI3sCEGFDCM7eA")

# --- DESIGN PAGE ---
st.set_page_config(page_title="Pic-to-Excel", page_icon="⚡", layout="wide")

# CSS pour cacher les menus Streamlit et styliser
st.markdown("""
<style>
    #MainMenu {visibility: hidden;}
    footer {visibility: hidden;}
    header {visibility: hidden;}
    .main-title {
        font-size: 2.5rem;
        font-weight: 700;
        margin-bottom: 0px;
    }
</style>
""", unsafe_allow_html=True)

# --- SIDEBAR (CHOIX DU CERVEAU) ---
with st.sidebar:
    st.header("Vitesse / Précision")
    choix_modele = st.radio(
        "Modèle IA :",
        ["Flash ⚡", "Pro 🎯"],
        index=1,
        help="Flash est rapide. Pro est minutieux."
    )
    
    # Mapping des modèles (Tes noms personnalisés)
    if "Flash" in choix_modele:
        MODEL_NAME = "gemini-2.5-flash"
    else:
        MODEL_NAME = "gemini-2.5-pro"

    # Initialisation silencieuse
    try:
        model = genai.GenerativeModel(MODEL_NAME)
    except Exception as e:
        st.error(f"Erreur modèle : {e}")

# --- TITRE ---
st.markdown(f'<p class="main-title">Pic-to-Excel</p>', unsafe_allow_html=True)
st.write("Extraction de données haute fidélité.")

# --- FONCTION EXCEL ---
def create_styled_excel(df):
    buffer = io.BytesIO()
    wb = Workbook()
    ws = wb.active
    ws.title = "Export"
    
    new_columns = []
    for i, col in enumerate(df.columns):
        col_str = str(col).strip()
        if "Unnamed" in col_str or col_str == "":
            new_columns.append(f"Col {i+1}")
        else:
            new_columns.append(col_str)
    df.columns = new_columns

    for r in dataframe_to_rows(df, index=False, header=True):
        ws.append(r)

    max_row = ws.max_row
    max_col = ws.max_column
    if max_row >= 2 and max_col >= 1:
        last_col_letter = get_column_letter(max_col)
        ref = f"A1:{last_col_letter}{max_row}"
        tab = Table(displayName="TabData", ref=ref)
        style = TableStyleInfo(name="TableStyleMedium9", showRowStripes=True)
        tab.tableStyleInfo = style
        ws.add_table(tab)

        for col_idx in range(1, max_col + 1):
            col_letter = get_column_letter(col_idx)
            max_len = 0
            for cell in ws[col_letter]:
                try:
                    if cell.value: max_len = max(max_len, len(str(cell.value)))
                except: pass
            ws.column_dimensions[col_letter].width = min((max_len + 2) * 1.2, 60)
            
    wb.save(buffer)
    return buffer.getvalue()

# --- INTERFACE PRINCIPALE ---
col1, col2 = st.columns([1, 1])

with col1:
    uploaded_file = st.file_uploader("Dépose ton image ici", type=["jpg", "png", "jpeg", "webp"])
    if uploaded_file:
        image = Image.open(uploaded_file)
        st.image(image, use_column_width=True)

if 'df_result' not in st.session_state:
    st.session_state.df_result = None

with col2:
    if uploaded_file and st.button('Extraire les données', type="primary"):
        with st.spinner('Analyse en cours...'):
            try:
                prompt = """
                Tu es un Expert Digitalisation.
                Transforme cette image en CSV (séparateur point-virgule).
                RÈGLES :
                1. Précision absolue sur les chiffres et textes.
                2. Si graphique : relève tous les points visibles.
                3. Pas de texte avant/après.
                """
                
                response = model.generate_content([prompt, image])
                raw_csv = response.text.replace("```csv", "").replace("```", "").strip()
                
                st.session_state.df_result = pd.read_csv(
                    io.StringIO(raw_csv), 
                    sep=";", 
                    engine="python",
                    on_bad_lines='skip' 
                )
                st.success("Terminé !")
                
            except Exception as e:
                st.error(f"Erreur : {e}")

# --- EXPORT ---
if st.session_state.df_result is not None:
    st.divider()
    edited_df = st.data_editor(st.session_state.df_result, num_rows="dynamic", use_container_width=True)
    excel_data = create_styled_excel(edited_df)
    
    c1, c2 = st.columns(2)
    with c1:
        st.download_button("Télécharger Excel", excel_data, "Export.xlsx", "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet", type="primary")
    with c2:
        st.download_button("Télécharger CSV", edited_df.to_csv(index=False, sep=";").encode('utf-8'), "export.csv", "text/csv")