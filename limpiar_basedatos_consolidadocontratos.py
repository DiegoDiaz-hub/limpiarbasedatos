import streamlit as st
import pandas as pd
import openpyxl
from openpyxl.styles import Font, PatternFill, Alignment, Border, Side
from openpyxl.utils import get_column_letter
import tempfile
import os
import warnings

warnings.filterwarnings('ignore')

# ─────────────────────────────────────────────────────────────
# ✅ LISTA MAESTRA ESTRICTA (SOLO ESTOS NOMBRES SOBREVIVEN)
# ─────────────────────────────────────────────────────────────
STRATEGIC_BUYERS = {
    'Patricio Espinoza', 'Jorge Urrutia', 'Bárbara García', 'Claudio Berrios',
    'Martina Cifuentes', 'Joseph España',
    'Michelle Palma', 'Juan Figueroa', 'Magdalena Farias', 'Denisse Andrea Gonzalez Terrile'
}

TACTICAL_BUYERS = {
    'Leonardo Nacarate', 'Martina Cifuentes', 'Scarlette Lucero',
    'Margarita Lineros', 'Erika Silva', 'Karina Satelo', 'Pablo Labs', 'Dayana Dávila', 'BPO'
}

ALL_VALID_BUYERS = STRATEGIC_BUYERS | TACTICAL_BUYERS

# Mapeo de variaciones comunes encontradas en el Pivot → Nombre Oficial
NAME_VARIATIONS = {
    'juan daniel figueroa': 'Juan Figueroa',
    'jorgue urrutia': 'Jorge Urrutia',
    'jorge uturria': 'Jorge Urrutia',
    'joseph eduardo españa escalona': 'Joseph España',
    'denisse andrea gonzales terrile': 'Denisse Andrea Gonzalez Terrile',
    'denisse andrea gonzalez': 'Denisse Andrea Gonzalez Terrile',
    'leonardo nacarete': 'Leonardo Nacarate',
    'margarita': 'Margarita Lineros',
    'erika': 'Erika Silva',
    'karina': 'Karina Satelo',
    'pablo labs': 'Pablo Labs',
    'bpo': 'BPO'
}

def normalize_name(name: str) -> str:
    if pd.isna(name) or str(name).strip() == '':
        return ''
    clean = str(name).strip().lower()
    # Eliminar acentos para matching robusto
    clean = ''.join(c for c in clean if c not in 'áéíóúüñ')
    return clean

def classify_buyer_strict(raw_name: str) -> tuple:
    clean_raw = normalize_name(raw_name)
    if not clean_raw: return '', ''
    
    # 1. Buscar en variaciones conocidas
    official_name = NAME_VARIATIONS.get(clean_raw, None)
    
    # 2. Si no hay variación, buscar match exacto en listas oficiales
    if official_name is None:
        # Intentar match directo con listas (sin acentos)
        for official in ALL_VALID_BUYERS:
            if clean_raw == normalize_name(official):
                official_name = official
                break
                
    # 3. Si no coincide con NADA oficial, retornar vacío (la fila será eliminada después)
    if official_name is None:
        return '', ''
        
    # 4. Asignar a columna según categoría
    if official_name in STRATEGIC_BUYERS:
        return official_name, ''
    elif official_name in TACTICAL_BUYERS:
        return '', official_name
    return '', ''

# ─────────────────────────────────────────────────────────────
# 📐 ESTRUCTURA EXACTA DEL CONSOLIDADO (30 Columnas)
# ─────────────────────────────────────────────────────────────
TARGET_HEADERS = [
    'Contrato Sap', 'Contrato Legado', 'Comprador Estratégico', 'Comprador Táctico',
    'Estado Contrato Ariba', 'Rut', 'Cód SAP', 'Proveedor', 'Fecha Inicio',
    'Fecha Término Contrato', 'Estado Contrato', 'Descripción', 'Área', 'Gerencia',
    'Planta', 'Ingresa a Planta', 'Aplica Boleta de Garantía (Ariba)',
    'Aplica Boleta de Garantía (Contrato firmado)', 'Tipo Garantía', 'N° Garantia',
    'Moneda Garantía', 'Monto Garantía', 'Vencimiento Garantía', 'Estado Garantía',
    'Administrador de Contrato', 'Correo Electrónico', 'Observación contrato Control Contratista',
    'Observación Control Contratistas Boleta de Garantía', 'Contratos Indefinidos', 'Observación Interna'
]

STYLES = {
    "header": {
        "fill": PatternFill(start_color="E7E6E6", end_color="E7E6E6", fill_type="solid"),
        "font": Font(name='Arial', size=8, bold=True),
        "alignment": Alignment(horizontal='center', vertical='center', wrap_text=True),
        "border": Border(left=Side('thin'), right=Side('thin'), top=Side('thin'), bottom=Side('thin'))
    },
    "data": {
        "font": Font(name='Arial', size=8),
        "border": Border(left=Side('thin'), right=Side('thin'), top=Side('thin'), bottom=Side('thin')),
        "alignment": Alignment(horizontal='left', vertical='center', wrap_text=False)
    }
}

COLUMN_WIDTHS = [14, 16, 20, 16, 20, 16, 12, 38, 14, 18, 16, 45, 16, 16, 12, 14, 26, 32, 14, 14, 14, 16, 16, 14, 20, 22, 35, 45, 18, 45]

def load_pivot(file_path: str) -> pd.DataFrame:
    try: df_scan = pd.read_excel(file_path, sheet_name='Data', header=None, nrows=50)
    except: df_scan = pd.read_excel(file_path, header=None, nrows=50)
    
    header_row = None
    for i, row in df_scan.iterrows():
        if any('ID de contrato' in str(v) for v in row if pd.notna(v)):
            header_row = i; break
    if header_row is None: raise ValueError("No se encontró 'ID de contrato' en la hoja Data.")
    
    df = pd.read_excel(file_path, sheet_name='Data' if 'Data' in pd.ExcelFile(file_path).sheet_names else 0, header=header_row)
    df.columns = [str(c).strip() for c in df.columns]
    return df

def transform_data(df_pivot: pd.DataFrame) -> tuple:
    df_out = pd.DataFrame(columns=TARGET_HEADERS)
    
    # Mapeo directo
    df_out['Contrato Sap'] = df_pivot.get('ID de contrato', pd.Series(dtype='object'))
    df_out['Estado Contrato Ariba'] = df_pivot.get('Estado del contrato', pd.Series(dtype='object'))
    df_out['Estado Contrato'] = df_out['Estado Contrato Ariba']
    df_out['Rut'] = df_pivot.get('Rut empresa proveedor', pd.Series(dtype='object'))
    df_out['Cód SAP'] = df_pivot.get('Código acreedor SAP', pd.Series(dtype='object'))
    df_out['Proveedor'] = df_pivot.get('Partes afectadas - Proveedor común', pd.Series(dtype='object'))
    df_out['Descripción'] = df_pivot.get('Descripción', pd.Series(dtype='object'))
    df_out['Contratos Indefinidos'] = df_pivot.get('Es Indefinido', pd.Series(dtype='object'))
    
    # Fechas
    for src, tgt in [('Fecha de entrada en vigor - Fecha', 'Fecha Inicio'), ('Fecha de expiración - Fecha', 'Fecha Término Contrato')]:
        if src in df_pivot.columns: df_out[tgt] = pd.to_datetime(df_pivot[src], errors='coerce').dt.strftime('%d/%m/%Y')
        else: df_out[tgt] = ''
        
    # 🔥 CLASIFICACIÓN ESTRICTA + FILTRADO
    raw_owners = df_pivot.get('Nombre del propietario', pd.Series(dtype='object')).fillna('').astype(str)
    classified = raw_owners.apply(classify_buyer_strict)
    
    df_out['Comprador Estratégico'] = [x[0] for x in classified]
    df_out['Comprador Táctico'] = [x[1] for x in classified]
    
    # 🗑️ ELIMINAR FILAS CON COMPRADORES NO OFICIALES
    mask_valid_buyer = (df_out['Comprador Estratégico'] != '') | (df_out['Comprador Táctico'] != '')
    dropped_invalid = (~mask_valid_buyer).sum()
    df_out = df_out[mask_valid_buyer].reset_index(drop=True)
    
    # 🚫 FILTRAR CONTRATOS CERRADOS
    mask_no_cerrado = ~df_out['Estado Contrato Ariba'].astype(str).str.strip().str.lower().isin(['cerrado', 'cerrados'])
    dropped_cerrados = (~mask_no_cerrado).sum()
    df_out = df_out[mask_no_cerrado].reset_index(drop=True)
    
    # Si táctico está vacío pero estratégico tiene valor, copiar estratégico a táctico
    mask_tactical_empty = (df_out['Comprador Táctico'] == '') & (df_out['Comprador Estratégico'] != '')
    df_out.loc[mask_tactical_empty, 'Comprador Táctico'] = df_out.loc[mask_tactical_empty, 'Comprador Estratégico']
    
    # Administrador y Correo
    df_out['Administrador de Contrato'] = df_out['Comprador Estratégico']
    df_out['Correo Electrónico'] = ''
    
    # Limpieza final
    df_out = df_out.fillna('')
    df_out = df_out.replace(['null', 'None', 'Unclassified', 'nan'], '')
    
    total_dropped = dropped_invalid + dropped_cerrados
    return df_out, dropped_invalid, dropped_cerrados

def apply_formatting(df: pd.DataFrame, output_path: str):
    wb = openpyxl.Workbook()
    ws = wb.active
    ws.title = "Consolidado de Contratos"
    
    for col_idx, header in enumerate(TARGET_HEADERS, start=1):
        cell = ws.cell(row=1, column=col_idx, value=header)
        cell.fill = STYLES["header"]["fill"]; cell.font = STYLES["header"]["font"]
        cell.alignment = STYLES["header"]["alignment"]; cell.border = STYLES["header"]["border"]
    ws.row_dimensions[1].height = 45
    
    for r_idx, row in df.iterrows():
        for c_idx, val in enumerate(row, start=1):
            cell = ws.cell(row=r_idx + 2, column=c_idx, value=val)
            cell.font = STYLES["data"]["font"]; cell.border = STYLES["data"]["border"]; cell.alignment = STYLES["data"]["alignment"]
            if c_idx in [9, 10]: cell.alignment = Alignment(horizontal='center', vertical='center')
            elif isinstance(val, (int, float)): cell.alignment = Alignment(horizontal='center', vertical='center'); cell.number_format = '#,##0.00'
                
    for i, w in enumerate(COLUMN_WIDTHS, start=1): ws.column_dimensions[get_column_letter(i)].width = w
    ws.auto_filter.ref = f"A1:{get_column_letter(len(TARGET_HEADERS))}{len(df) + 1}"
    ws.freeze_panes = "C2"
    wb.save(output_path)

# ─────────────────────────────────────────────────────────────
# 🌐 STREAMLIT UI
# ─────────────────────────────────────────────────────────────
st.set_page_config(page_title="Generador Consolidado", layout="centered")
st.title("📑 Generador de Consolidado de Contratos")
st.caption("Sube el Pivot. Se aplicará **filtro estricto de compradores oficiales** y se eliminarán contratos cerrados.")

uploaded_file = st.file_uploader("📥 Archivo Pivot (.xlsx)", type=["xlsx"])

if uploaded_file:
    with st.spinner("Validando lista maestra, filtrando cerrados y limpiando datos..."):
        try:
            with tempfile.NamedTemporaryFile(delete=False, suffix="_pivot.xlsx") as tmp:
                tmp.write(uploaded_file.getvalue()); pivot_path = tmp.name
            
            df_pivot = load_pivot(pivot_path)
            df_final, dropped_invalid, dropped_cerrados = transform_data(df_pivot)
            
            out_path = pivot_path.replace("_pivot.xlsx", "_Consolidado_Limpio.xlsx")
            apply_formatting(df_final, out_path)
            
            with open(out_path, "rb") as f:
                st.download_button(
                    label="📥 Descargar Consolidado Limpio",
                    data=f,
                    file_name="Consolidado de Contratos.xlsx",
                    mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                    use_container_width=True
                )
            
            st.success(f"✅ **Archivo generado y depurado.**\n• 🗑️ **{dropped_invalid} contratos eliminados** por compradores no oficiales.\n• 🚫 **{dropped_cerrados} contratos eliminados** por estado 'Cerrado'.\n• 📄 **{len(df_final)} contratos válidos** en el archivo final.\n• Formato 100% idéntico al original.")
            
            os.unlink(pivot_path); os.unlink(out_path)
        except Exception as e:
            st.error(f"❌ Error: {str(e)}"); st.exception(e)
