import streamlit as st
import pdfplumber
import pandas as pd
import io
import pytesseract
import re
import fitz  # PyMuPDF
from PIL import Image
from openpyxl.styles import PatternFill, Font, Alignment
from openpyxl.utils import get_column_letter

# ============================================================
# CONFIGURACIÓN Y ESTILO UI ESTABLE
# ============================================================
st.set_page_config(page_title="INVIMA Data Engine", layout="wide", page_icon="🏛️")

# Estilo simplificado para evitar errores de renderizado en el navegador
st.markdown("""
    <style>
    .main-title {
        color: #1a365d;
        font-weight: 800;
        text-align: center;
        font-size: 2.5rem;
        margin-bottom: 2rem;
    }
    .stButton>button {
        width: 100%;
        font-weight: 600;
    }
    </style>
    """, unsafe_allow_html=True)

# ============================================================
# MOTOR DE EXTRACCIÓN (HÍBRIDO + OCR REAL)
# ============================================================

def extraer_con_ocr(file):
    """Convierte PDF a imagen usando PyMuPDF y aplica OCR con Tesseract."""
    filas = []
    file.seek(0)
    pdf_bytes = file.read()
    
    # Abrir PDF desde memoria
    doc = fitz.open(stream=pdf_bytes, filetype="pdf")
    
    for page in doc:
        # Alta resolución para OCR (300 DPI)
        zoom = 300 / 72
        mat = fitz.Matrix(zoom, zoom)
        pix = page.get_pixmap(matrix=mat)
        
        img = Image.open(io.BytesIO(pix.tobytes("png")))
        
        # OCR optimizado para tablas
        text = pytesseract.image_to_string(img, lang='spa', config='--psm 6')
        
        for line in text.split('\n'):
            # Dividir por múltiples espacios para detectar columnas
            partes = [p.strip() for p in re.split(r'\t| {2,}', line) if p.strip()]
            if len(partes) > 1:
                filas.append(partes)
    
    doc.close()
    return filas

def extraer_con_limpieza(file, modo_ocr=False):
    if modo_ocr:
        return extraer_con_ocr(file)
    
    filas = []
    file.seek(0)
    try:
        with pdfplumber.open(file) as pdf:
            for page in pdf.pages:
                # Estrategia de líneas para INVIMA
                table = page.extract_table({"vertical_strategy": "lines", "horizontal_strategy": "lines"})
                if not table:
                    table = page.extract_table()
                
                if table:
                    for row in table:
                        clean_row = [" ".join(str(c).split()) if c else "" for c in row]
                        if any(clean_row): filas.append(clean_row)
    except:
        return extraer_con_ocr(file)
    
    if not filas:
        return extraer_con_ocr(file)
        
    return filas

# ============================================================
# LÓGICA DE EXPORTACIÓN EXCEL
# ============================================================

def generar_excel_profesional(final_data):
    output = io.BytesIO()
    with pd.ExcelWriter(output, engine='openpyxl') as writer:
        for name, df in final_data.items():
            # Limitar nombre de pestaña a 30 caracteres
            sheet_name = name[:30].replace("[", "").replace("]", "").replace("*", "").replace("?", "").replace("/", "").replace("\\", "")
            df.to_excel(writer, index=False, sheet_name=sheet_name)
            
            ws = writer.sheets[sheet_name]
            azul_invima = PatternFill(start_color="002F6C", end_color="002F6C", fill_type="solid")
            fuente_blanca = Font(color="FFFFFF", bold=True)
            
            ws.sheet_view.showGridLines = False
            
            for col_num, value in enumerate(df.columns, 1):
                cell = ws.cell(row=1, column=col_num)
                cell.fill = azul_invima
                cell.font = fuente_blanca
                cell.alignment = Alignment(horizontal="center")
                
                # Ajuste automático de ancho
                ws.column_dimensions[get_column_letter(col_num)].width = 25

    return output.getvalue()

# ============================================================
# INTERFAZ DE USUARIO (SIN HTML CONFLICTIVO)
# ============================================================

st.markdown("<h1 class='main-title'>🏛️ INVIMA DATA ENGINE</h1>", unsafe_allow_html=True)

with st.sidebar:
    st.header("🛠️ Configuración")
    modo_scan = st.toggle("🔦 Activar Motor OCR", help="Activa esto para documentos escaneados.")
    if st.button("🔄 Reiniciar Aplicación"):
        st.session_state.clear()
        st.rerun()

# Inicializar estado
if 'master' not in st.session_state:
    st.session_state.master = {}

uploaded_files = st.file_uploader("Cargar Reportes PDF", type=["pdf"], accept_multiple_files=True)

if uploaded_files:
    for i, f in enumerate(uploaded_files):
        # Contenedor nativo con borde, mucho más estable que HTML manual
        with st.container(border=True):
            col_info, col_btn = st.columns([3, 1])
            col_info.subheader(f"📄 {f.name}")
            
            if col_btn.button("Analizar PDF", key=f"btn_{i}"):
                with st.spinner("Analizando estructura..."):
                    datos = extraer_con_limpieza(f, modo_scan)
                    if datos:
                        st.session_state.master[f.name] = {"df": pd.DataFrame(datos), "clean": None}
                    else:
                        st.error("No se detectaron tablas.")

            # Si ya fue analizado, mostrar previsualización
            if f.name in st.session_state.master:
                res = st.session_state.master[f.name]
                df_raw = res["df"]
                
                st.write("---")
                h_idx = st.number_input(f"Fila de encabezados para {f.name}:", 0, max(0, len(df_raw)-1), 0, key=f"h_{i}")
                st.dataframe(df_raw.head(10), use_container_width=True)
                
                if st.button("💎 Validar y Limpiar Tabla", key=f"fix_{i}"):
                    try:
                        df_final = df_raw.iloc[h_idx:].copy()
                        df_final.columns = [str(c).strip() for c in df_final.iloc[0]]
                        df_final = df_final[1:].reset_index(drop=True)
                        st.session_state.master[f.name]["clean"] = df_final
                        st.success(f"✅ {f.name} listo para exportar.")
                    except Exception as e:
                        st.error(f"Error al limpiar: {e}")

    # Sección de Descarga
    ready_data = {k: v["clean"] for k, v in st.session_state.master.items() if v.get("clean") is not None}
    
    if ready_data:
        st.markdown("### 🚀 Procesamiento Completo")
        excel_file = generar_excel_profesional(ready_data)
        st.download_button(
            label="📥 DESCARGAR TODOS LOS ARCHIVOS EN EXCEL",
            data=excel_file,
            file_name="CONSOLIDADO_INVIMA.xlsx",
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
            use_container_width=True
        )
