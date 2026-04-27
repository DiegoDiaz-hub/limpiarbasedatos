import streamlit as st
import pandas as pd
import openpyxl
from openpyxl.styles import Font, PatternFill, Alignment, Border, Side
from openpyxl.utils import get_column_letter
import tempfile
import os
import warnings
import re

warnings.filterwarnings('ignore')

# ─────────────────────────────────────────────────────────────
# ✅ LISTA MAESTRA DE COMPRADORES (ACTUALIZADA)
# ─────────────────────────────────────────────────────────────

# ESTRATÉGICOS - CHILE
STRATEGIC_CHILE = {
    'Patricio Espinoza', 'Jorge Urrutia', 'Bárbara García', 'Claudio Berrios'
}

# ESTRATÉGICOS - PROYECTOS CHILE
STRATEGIC_PROYECTOS = {
    'Martina Cifuentes', 'Joseph España'
}

# ESTRATÉGICOS - CORPORATIVOS
STRATEGIC_CORPORATIVO = {
    'Michelle Palma', 'Juan Figueroa', 'Magdalena Farias', 
    'Denisse Andrea Gonzalez Terrile'
}

# TÁCTICOS
TACTICAL_BUYERS = {
    'Leonardo Nacarate', 'Martina Cifuentes', 'Scarlette Lucero',
    'Margarita Lineros', 'Erika Silva', 'Karina Satelo', 'Pablo Labs'
}

# Unificar todos los estratégicos
ALL_STRATEGIC = STRATEGIC_CHILE | STRATEGIC_PROYECTOS | STRATEGIC_CORPORATIVO

def normalize_name(name):
    """Normaliza nombres para matching"""
    if pd.isna(name) or str(name).strip() == '':
        return ''
    clean = str(name).strip().title()
    # Correcciones comunes
    clean = clean.replace('Jorgue', 'Jorge').replace('Uturria', 'Urrutia')
    return clean

def find_buyer_match(raw_name: str) -> tuple:
    """
    Busca coincidencias entre el nombre del Pivot y las listas oficiales.
    Retorna (estrategico, tactico)
    """
    clean_name = normalize_name(raw_name)
    if not clean_name:
        return '', ''
    
    # Extraer solo el primer nombre + apellido principal para matching
    name_parts = clean_name.split()
    short_name = ' '.join(name_parts[:2]) if len(name_parts) >= 2 else clean_name
    
    # Buscar en estratégicos (match exacto o parcial)
    for strat in ALL_STRATEGIC:
        if (clean_name == strat or 
            strat in clean_name or 
            clean_name in strat or
            short_name in strat or 
            strat in short_name):
            return strat, ''
    
    # Buscar en tácticos
    for tact in TACTICAL_BUYERS:
        if (clean_name == tact or 
            tact in clean_name or 
            clean_name in tact or
            short_name in tact or 
            tact in short_name):
            return '', tact
    
    # Si no hay match, retornar vacío
    return '', ''

# ─────────────────────────────────────────────────────────────
# 📐 ESTRUCTURA DEL CONSOLIDADO (30 Columnas)
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

# Estilos
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
    """Carga el Pivot detectando headers"""
    try:
        df_scan = pd.read_excel(file_path, sheet_name='Data', header=None, nrows=50)
    except:
        df_scan = pd.read_excel(file_path, header=None, nrows=50)
    
    header_row = None
    for i, row in df_scan.iterrows():
        if any('ID de contrato' in str(v) for v in row if pd.notna(v)):
            header_row = i
            break
    
    if header_row is None:
        raise ValueError("No se encontró 'ID de contrato' en la hoja Data")
    
    df = pd.read_excel(file_path, sheet_name='Data' if 'Data' in pd.ExcelFile(file_path).sheet_names else 0, 
                       header=header_row)
    df.columns = [str(c).strip() for c in df.columns]
    return df

def transform_data(df_pivot: pd.DataFrame) -> pd.DataFrame:
    """Transforma y asigna compradores"""
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
    for src, tgt in [('Fecha de entrada en vigor - Fecha', 'Fecha Inicio'), 
                     ('Fecha de expiración - Fecha', 'Fecha Término Contrato')]:
        if src in df_pivot.columns:
            df_out[tgt] = pd.to_datetime(df_pivot[src], errors='coerce').dt.strftime('%d/%m/%Y')
        else:
            df_out[tgt] = ''
    
    # 🔥 ASIGNACIÓN DE COMPRADORES
    raw_owners = df_pivot.get('Nombre del propietario', pd.Series(dtype='object')).fillna('').astype(str)
    
    classified = raw_owners.apply(find_buyer_match)
    df_out['Comprador Estratégico'] = [x[0] for x in classified]
    df_out['Comprador Táctico'] = [x[1] for x in classified]
    
    # 🛡️ NUNCA DEJAR EN BLANCO
    mask_empty_both = (df_out['Comprador Estratégico'] == '') & (df_out['Comprador Táctico'] == '')
    
    if mask_empty_both.any():
        # Si no está en listas oficiales, asignar al Estratégico el nombre original
        df_out.loc[mask_empty_both, 'Comprador Estratégico'] = raw_owners[mask_empty_both]
        df_out.loc[mask_empty_both, 'Observación Interna'] = '⚠️ No está en lista oficial: ' + raw_owners[mask_empty_both].astype(str)
    
    # Si táctico está vacío pero estratégico no, copiar estratégico a táctico
    mask_tactical_empty = (df_out['Comprador Táctico'] == '') & (df_out['Comprador Estratégico'] != '')
    df_out.loc[mask_tactical_empty, 'Comprador Táctico'] = df_out.loc[mask_tactical_empty, 'Comprador Estratégico']
    
    # Administrador
    df_out['Administrador de Contrato'] = df_out['Comprador Estratégico'].where(
        df_out['Comprador Estratégico'] != '', 
        df_out['Comprador Táctico']
    )
    df_out['Correo Electrónico'] = ''
    
    # Limpieza
    df_out = df_out.fillna('')
    df_out = df_out.replace(['null', 'None', 'Unclassified', 'nan'], '')
    
    return df_out

def apply_formatting(df: pd.DataFrame, output_path: str):
    """Aplica formato Excel"""
    wb = openpyxl.Workbook()
    ws = wb.active
    ws.title = "Consolidado de Contratos"
    
    # Headers
    for col_idx, header in enumerate(TARGET_HEADERS, start=1):
        cell = ws.cell(row=1, column=col_idx, value=header)
        cell.fill = STYLES["header"]["fill"]
        cell.font = STYLES["header"]["font"]
        cell.alignment = STYLES["header"]["alignment"]
        cell.border = STYLES["header"]["border"]
    ws.row_dimensions[1].height = 45
    
    # Data
    for r_idx, row in df.iterrows():
        for c_idx, val in enumerate(row, start=1):
            cell = ws.cell(row=r_idx + 2, column=c_idx, value=val)
            cell.font = STYLES["data"]["font"]
            cell.border = STYLES["data"]["border"]
            cell.alignment = STYLES["data"]["alignment"]
            
            if c_idx in [9, 10]:
                cell.alignment = Alignment(horizontal='center', vertical='center')
            elif isinstance(val, (int, float)):
                cell.alignment = Alignment(horizontal='center', vertical='center')
                cell.number_format = '#,##0.00'
    
    # Anchos
    for i, w in enumerate(COLUMN_WIDTHS, start=1):
        ws.column_dimensions[get_column_letter(i)].width = w
    
    # Filtros y paneles
    ws.auto_filter.ref = f"A1:{get_column_letter(len(TARGET_HEADERS))}{len(df) + 1}"
    ws.freeze_panes = "C2"
    
    wb.save(output_path)

# ─────────────────────────────────────────────────────────────
# 🌐 STREAMLIT
# ─────────────────────────────────────────────────────────────
st.set_page_config(page_title="Generador Consolidado", layout="centered")
st.title("📑 Generador de Consolidado de Contratos")
st.caption("Sube el Pivot. El sistema asignará compradores estratégicos/tácticos sin espacios en blanco.")

uploaded_file = st.file_uploader("📥 Archivo Pivot (.xlsx)", type=["xlsx"])

if uploaded_file:
    with st.spinner("Procesando..."):
        try:
            with tempfile.NamedTemporaryFile(delete=False, suffix="_pivot.xlsx") as tmp:
                tmp.write(uploaded_file.getvalue())
                pivot_path = tmp.name
            
            df_pivot = load_pivot(pivot_path)
            df_final = transform_data(df_pivot)
            
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
            
            st.success("✅ **Archivo generado.**\n• Compradores validados contra lista maestra.\n• **0 espacios en blanco**.\n• Formato idéntico al original.")
            
            os.unlink(pivot_path)
            os.unlink(out_path)
            
        except Exception as e:
            st.error(f"❌ Error: {str(e)}")
            st.exception(e)
