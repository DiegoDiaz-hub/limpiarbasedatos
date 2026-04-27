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
# ✅ LISTA MAESTRA DE COMPRADORES (CLASIFICACIÓN ESTRICTA)
# ─────────────────────────────────────────────────────────────
STRATEGIC_CHILE = {'Patricio Espinoza', 'Jorge Urrutia', 'Bárbara García', 'Claudio Berrios'}
STRATEGIC_PROYECTOS = {'Martina Cifuentes', 'Joseph España'}
STRATEGIC_CORPORATIVO = {'Michelle Palma', 'Juan Figueroa', 'Magdalena Farias', 'Denisse Andrea Gonzalez Terrile'}
ALL_STRATEGIC = STRATEGIC_CHILE | STRATEGIC_PROYECTOS | STRATEGIC_CORPORATIVO

TACTICAL_BUYERS = {'Leonardo Nacarate', 'Martina Cifuentes', 'Scarlette Lucero', 'Margarita Lineros', 'Erika Silva', 'Karina Satelo', 'Pablo Labs', 'Dayana Dávila', 'BPO'}

def normalize_name(name):
    if pd.isna(name) or str(name).strip() == '': return ''
    clean = str(name).strip().title()
    return clean.replace('Jorgue', 'Jorge').replace('Uturria', 'Urrutia')

def classify_buyer_robust(raw_name: str) -> tuple:
    clean_name = normalize_name(raw_name)
    if not clean_name: return '', ''
    
    name_parts = clean_name.split()
    short_name = ' '.join(name_parts[:2]) if len(name_parts) >= 2 else clean_name
    
    for strat in ALL_STRATEGIC:
        if clean_name == strat or strat in clean_name or clean_name in strat or short_name in strat or strat in short_name:
            return strat, ''
    for tact in TACTICAL_BUYERS:
        if clean_name == tact or tact in clean_name or clean_name in tact or short_name in tact or tact in short_name:
            return '', tact
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

def transform_data(df_pivot: pd.DataFrame) -> pd.DataFrame:
    df_out = pd.DataFrame(columns=TARGET_HEADERS)
    
    # 1. Mapeo directo
    df_out['Contrato Sap'] = df_pivot.get('ID de contrato', pd.Series(dtype='object'))
    df_out['Estado Contrato Ariba'] = df_pivot.get('Estado del contrato', pd.Series(dtype='object'))
    df_out['Estado Contrato'] = df_out['Estado Contrato Ariba']
    df_out['Rut'] = df_pivot.get('Rut empresa proveedor', pd.Series(dtype='object'))
    df_out['Cód SAP'] = df_pivot.get('Código acreedor SAP', pd.Series(dtype='object'))
    df_out['Proveedor'] = df_pivot.get('Partes afectadas - Proveedor común', pd.Series(dtype='object'))
    df_out['Descripción'] = df_pivot.get('Descripción', pd.Series(dtype='object'))
    df_out['Contratos Indefinidos'] = df_pivot.get('Es Indefinido', pd.Series(dtype='object'))
    
    # 2. Fechas
    for src, tgt in [('Fecha de entrada en vigor - Fecha', 'Fecha Inicio'), ('Fecha de expiración - Fecha', 'Fecha Término Contrato')]:
        if src in df_pivot.columns: df_out[tgt] = pd.to_datetime(df_pivot[src], errors='coerce').dt.strftime('%d/%m/%Y')
        else: df_out[tgt] = ''
    
    # 3. Clasificación de compradores
    raw_owners = df_pivot.get('Nombre del propietario', pd.Series(dtype='object')).fillna('').astype(str)
    classified = raw_owners.apply(classify_buyer_robust)
    df_out['Comprador Estratégico'] = [x[0] for x in classified]
    df_out['Comprador Táctico'] = [x[1] for x in classified]
    
    # 4. Regla: Nunca dejar sin encargado
    mask_empty_both = (df_out['Comprador Estratégico'] == '') & (df_out['Comprador Táctico'] == '')
    if mask_empty_both.any():
        df_out.loc[mask_empty_both, 'Comprador Estratégico'] = raw_owners[mask_empty_both]
        df_out.loc[mask_empty_both, 'Observación Interna'] = '⚠️ No está en lista oficial: ' + raw_owners[mask_empty_both].astype(str)
    
    mask_tactical_empty = (df_out['Comprador Táctico'] == '') & (df_out['Comprador Estratégico'] != '')
    df_out.loc[mask_tactical_empty, 'Comprador Táctico'] = df_out.loc[mask_tactical_empty, 'Comprador Estratégico']
    
    df_out['Administrador de Contrato'] = df_out['Comprador Estratégico'].where(df_out['Comprador Estratégico'] != '', df_out['Comprador Táctico'])
    df_out['Correo Electrónico'] = ''
    
    # 5. 🚫 FILTRAR CONTRATOS CERRADOS
    status_col = 'Estado Contrato Ariba'
    if status_col in df_out.columns:
        # Elimina filas donde el estado sea "Cerrado" o "Cerrados" (ignora mayúsculas/minúsculas y espacios)
        mask_no_cerrado = ~df_out[status_col].astype(str).str.strip().str.lower().isin(['cerrado', 'cerrados'])
        removed_count = (~mask_no_cerrado).sum()
        df_out = df_out[mask_no_cerrado].reset_index(drop=True)
    else:
        removed_count = 0

    # 6. Limpieza final
    df_out = df_out.fillna('')
    df_out = df_out.replace(['null', 'None', 'Unclassified', 'nan'], '')
    
    return df_out, removed_count

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
st.caption("Sube el Pivot. Se filtrarán contratos **Cerrados** y se asignarán compradores sin espacios en blanco.")

uploaded_file = st.file_uploader("📥 Archivo Pivot (.xlsx)", type=["xlsx"])

if uploaded_file:
    with st.spinner("Procesando, validando compradores y filtrando cerrados..."):
        try:
            with tempfile.NamedTemporaryFile(delete=False, suffix="_pivot.xlsx") as tmp:
                tmp.write(uploaded_file.getvalue()); pivot_path = tmp.name
            
            df_pivot = load_pivot(pivot_path)
            df_final, removed_cerrados = transform_data(df_pivot)
            
            out_path = pivot_path.replace("_pivot.xlsx", "_Consolidado_Final.xlsx")
            apply_formatting(df_final, out_path)
            
            with open(out_path, "rb") as f:
                st.download_button(
                    label="📥 Descargar Consolidado",
                    data=f,
                    file_name="Consolidado de Contratos.xlsx",
                    mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                    use_container_width=True
                )
            
            st.success(f"✅ **Archivo generado.**\n• **{removed_cerrados} contratos 'Cerrados' eliminados.**\n• {len(df_final)} contratos activos/proceso en el archivo.\n• Formato y compradores validados.")
            
            os.unlink(pivot_path); os.unlink(out_path)
        except Exception as e:
            st.error(f"❌ Error: {str(e)}"); st.exception(e)
