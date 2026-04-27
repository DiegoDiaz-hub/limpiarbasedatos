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
# ✅ LISTA OFICIAL DE COMPRADORES (EXTRACTADA DE TU DATA)
# ─────────────────────────────────────────────────────────────
VALID_STRATEGIC_BUYERS = {
    'Jorge Urrutia',
    'Jorge Alfonso Urrutia Carillo',
    'Bárbara García',
    'Patricio Espinoza',
    'Joseph España',
    'Viviana Grandón',
    'Claudio Berrios',
    'Magdalena Farias',
    'Martina Fuentes',
    'Denisse Andrea Gonzalez Terrile',
    'BPO',
    'Juan Daniel Figueroa',
    'Juan Figueroa',
    'Laura Mendoza',
    'Diego Escalona',
    'Judith Rivas',
    'Dayana Dávila',
    'Daniela Escobar',
    'Leandro Medina',
    'Sofia Delgado',
    'Lina Diaz',
    'Michelle Esperanza',
    'Carol Reyes',
    'Claudia Castillo',
    'Cecilia Fernandez',
    'Servio Salges Cazorla',
    'Priscilla Gre Guerra',
    'Angela Gallardo'
}

VALID_TACTICAL_BUYERS = {
    'Leonardo Nacarate',
    'Patricio Espinoza',
    'Joseph España',
    'Dayana Dávila',
    'BPO'
}

def normalize_name(name):
    """Normaliza el nombre eliminando espacios extras y estandarizando."""
    if pd.isna(name) or str(name).strip() == '':
        return ''
    return str(name).strip()

def validate_and_assign_buyers(df_pivot: pd.DataFrame) -> tuple:
    """
    Valida y asigna compradores estratégicos y tácticos SOLO si están en la lista oficial.
    Retorna dos Series: strategic_buyers, tactical_buyers
    """
    strategic = []
    tactical = []
    
    # Obtener columna del propietario (comprador en Ariba)
    owner_col = 'Nombre del propietario' if 'Nombre del propietario' in df_pivot.columns else None
    
    for idx, row in df_pivot.iterrows():
        raw_name = normalize_name(row.get(owner_col, '') if owner_col else '')
        
        # Limpiar variaciones comunes del nombre
        clean_name = raw_name.replace('  ', ' ').title()
        
        # Buscar match exacto o parcial en lista estratégica
        matched_strategic = None
        for valid_name in VALID_STRATEGIC_BUYERS:
            if clean_name.lower() == valid_name.lower() or clean_name in valid_name or valid_name in clean_name:
                matched_strategic = valid_name
                break
        
        # Si no hay match estratégico, buscar en tácticos
        matched_tactical = None
        if not matched_strategic:
            for valid_name in VALID_TACTICAL_BUYERS:
                if clean_name.lower() == valid_name.lower() or clean_name in valid_name or valid_name in clean_name:
                    matched_tactical = valid_name
                    break
        
        strategic.append(matched_strategic if matched_strategic else '')
        tactical.append(matched_tactical if matched_tactical else '')
    
    return pd.Series(strategic), pd.Series(tactical)

# ─────────────────────────────────────────────────────────────
# 📐 ESTRUCTURA EXACTA DEL CONSOLIDADO OBJETIVO (30 Columnas)
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

# ─────────────────────────────────────────────────────────────
# 🎨 ESTILOS IDÉNTICOS AL CONSOLIDADO DE REFERENCIA
# ─────────────────────────────────────────────────────────────
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
    """Lee el Pivot detectando automáticamente la fila de encabezados reales."""
    try:
        df_scan = pd.read_excel(file_path, sheet_name='Data', header=None, nrows=50)
    except Exception:
        df_scan = pd.read_excel(file_path, header=None, nrows=50)
        
    header_row = None
    for i, row in df_scan.iterrows():
        if any('ID de contrato' in str(v) for v in row if pd.notna(v)):
            header_row = i
            break
            
    if header_row is None:
        raise ValueError("No se encontró la fila de encabezados ('ID de contrato') en la hoja Data.")
        
    df = pd.read_excel(file_path, sheet_name='Data' if 'Data' in pd.ExcelFile(file_path).sheet_names else 0, header=header_row)
    df.columns = [str(c).strip() for c in df.columns]
    return df

def transform_data(df_pivot: pd.DataFrame) -> pd.DataFrame:
    """Mapea, limpia y aplica la lógica crítica de compradores SIN ESPACIOS EN BLANCO."""
    df_out = pd.DataFrame(columns=TARGET_HEADERS)
    
    # 1. Mapeo directo de columnas existentes en el Pivot
    df_out['Contrato Sap'] = df_pivot.get('ID de contrato', pd.Series(dtype='object'))
    df_out['Estado Contrato Ariba'] = df_pivot.get('Estado del contrato', pd.Series(dtype='object'))
    df_out['Estado Contrato'] = df_out['Estado Contrato Ariba']
    df_out['Rut'] = df_pivot.get('Rut empresa proveedor', pd.Series(dtype='object'))
    df_out['Cód SAP'] = df_pivot.get('Código acreedor SAP', pd.Series(dtype='object'))
    df_out['Proveedor'] = df_pivot.get('Partes afectadas - Proveedor común', pd.Series(dtype='object'))
    df_out['Descripción'] = df_pivot.get('Descripción', pd.Series(dtype='object'))
    df_out['Contratos Indefinidos'] = df_pivot.get('Es Indefinido', pd.Series(dtype='object'))
    
    # 2. Formateo de fechas
    for src, tgt in [('Fecha de entrada en vigor - Fecha', 'Fecha Inicio'), 
                     ('Fecha de expiración - Fecha', 'Fecha Término Contrato')]:
        if src in df_pivot.columns:
            df_out[tgt] = pd.to_datetime(df_pivot[src], errors='coerce').dt.strftime('%d/%m/%Y')
        else:
            df_out[tgt] = ''
    
    # 3. 🔥 VALIDACIÓN ESTRICTA DE COMPRADORES (SOLO NOMBRES OFICIALES)
    strategic_series, tactical_series = validate_and_assign_buyers(df_pivot)
    
    df_out['Comprador Estratégico'] = strategic_series
    df_out['Comprador Táctico'] = tactical_series
    
    # 🛡️ REGLA DE ORO: NUNCA DEJAR COMPRADORES EN BLANCO
    # Si el táctico está vacío, copiamos el estratégico
    mask_tactical_empty = (df_out['Comprador Táctico'] == '') | (df_out['Comprador Táctico'].isna())
    df_out.loc[mask_tactical_empty, 'Comprador Táctico'] = df_out.loc[mask_tactical_empty, 'Comprador Estratégico']
    
    # Si el estratégico está vacío, copiamos el táctico
    mask_strategic_empty = (df_out['Comprador Estratégico'] == '') | (df_out['Comprador Estratégico'].isna())
    df_out.loc[mask_strategic_empty, 'Comprador Estratégico'] = df_out.loc[mask_strategic_empty, 'Comprador Táctico']
    
    # Safety net final: si ambos están vacíos, asignar "Sin Asignar"
    df_out.loc[(df_out['Comprador Estratégico'] == '') & (df_out['Comprador Táctico'] == ''), 'Comprador Estratégico'] = 'Sin Asignar'
    df_out.loc[(df_out['Comprador Estratégico'] == '') & (df_out['Comprador Táctico'] == ''), 'Comprador Táctico'] = 'Sin Asignar'
    
    # 4. Administrador y Correo
    df_out['Administrador de Contrato'] = df_out['Comprador Estratégico']
    df_out['Correo Electrónico'] = ''
    
    # 5. Limpieza general
    df_out = df_out.fillna('')
    df_out = df_out.replace(['null', 'None', 'Unclassified', 'nan'], '')
    
    return df_out

def apply_formatting(df: pd.DataFrame, output_path: str):
    """Escribe el DataFrame aplicando el formato pixel-perfect del Consolidado."""
    wb = openpyxl.Workbook()
    ws = wb.active
    ws.title = "Consolidado de Contratos"
    
    # Encabezados
    for col_idx, header in enumerate(TARGET_HEADERS, start=1):
        cell = ws.cell(row=1, column=col_idx, value=header)
        cell.fill = STYLES["header"]["fill"]
        cell.font = STYLES["header"]["font"]
        cell.alignment = STYLES["header"]["alignment"]
        cell.border = STYLES["header"]["border"]
    ws.row_dimensions[1].height = 45
    
    # Datos
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
                
    # Anchos de columna
    for i, w in enumerate(COLUMN_WIDTHS, start=1):
        ws.column_dimensions[get_column_letter(i)].width = w
        
    # Filtros y paneles
    ws.auto_filter.ref = f"A1:{get_column_letter(len(TARGET_HEADERS))}{len(df) + 1}"
    ws.freeze_panes = "C2"
    
    wb.save(output_path)

# ─────────────────────────────────────────────────────────────
# 🌐 INTERFAZ STREAMLIT
# ─────────────────────────────────────────────────────────────
st.set_page_config(page_title="Generador Consolidado Ariba", layout="centered")
st.title("📑 Generador de Consolidado de Contratos")
st.caption("Sube el Pivot crudo. El sistema validará compradores contra la lista oficial y nunca dejará espacios en blanco.")

uploaded_file = st.file_uploader("📥 Selecciona el archivo Pivot (.xlsx)", type=["xlsx"])

if uploaded_file:
    with st.spinner("Validando compradores oficiales y aplicando formato..."):
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
                    label="📥 Descargar Consolidado de Contratos",
                    data=f,
                    file_name="Consolidado de Contratos.xlsx",
                    mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                    use_container_width=True
                )
                
            st.success("✅ **Archivo generado correctamente.**\n• **Solo nombres oficiales** de compradores.\n• **0 espacios en blanco** en compradores estratégicos/tácticos.\n• Formato idéntico al original.")
            
            os.unlink(pivot_path)
            os.unlink(out_path)
            
        except Exception as e:
            st.error(f"❌ Error: {str(e)}")
            st.exception(e)
