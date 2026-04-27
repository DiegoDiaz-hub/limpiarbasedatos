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
# ✅ LISTA MAESTRA ESTRICTA DE COMPRADORES
# ─────────────────────────────────────────────────────────────
STRATEGIC_BUYERS = {
    'Patricio Espinoza', 'Jorge Urrutia', 'Bárbara García', 'Claudio Berrios',
    'Martina Fuentes', 'Joseph España', 'Michelle Palma', 'Juan Figueroa',
    'Magdalena Farias', 'Denisse Andrea Gonzalez Terrile'
}

TACTICAL_BUYERS = {
    'Leonardo Nacarate', 'Martina Fuentes', 'Scarlette Lucero',
    'Margarita Lineros', 'Erika Silva', 'Karina Satelo', 'Pablo Labs'
}

TYPO_CORRECTIONS = {
    'jorge uturria': 'Jorge Urrutia', 'jorgue urrutia': 'Jorge Urrutia',
    'dennis andrea gonzales': 'Denisse Andrea Gonzalez Terrile',
    'denisse andrea gonzalez terrile': 'Denisse Andrea Gonzalez Terrile',
    'juan daniel figueroa': 'Juan Figueroa',
    'joseph eduardo españa escalona': 'Joseph España',
    'michelle esperanza': 'Michelle Palma',
    'leonardo nacarete': 'Leonardo Nacarate'
}

def normalize_name(name: str) -> str:
    if pd.isna(name) or str(name).strip() == '': return ''
    clean = str(name).strip().lower()
    clean = ''.join(c for c in clean if c not in 'áéíóúüñ')
    return clean

def classify_buyer_strict(raw_name: str) -> tuple:
    clean_raw = normalize_name(raw_name)
    if not clean_raw: return None, None
    for typo, correct in TYPO_CORRECTIONS.items():
        if typo in clean_raw or clean_raw in typo:
            clean_raw = normalize_name(correct); break
    for official in STRATEGIC_BUYERS:
        if clean_raw == normalize_name(official) or clean_raw in normalize_name(official) or normalize_name(official) in clean_raw:
            return 'strategic', official
    for official in TACTICAL_BUYERS:
        if clean_raw == normalize_name(official) or clean_raw in normalize_name(official) or normalize_name(official) in clean_raw:
            return 'tactical', official
    return None, None

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

# 🔄 MAPEO REAL DE CAMPOS ARIBA → CONSOLIDADO (Nombres técnicos del Pivot)
ARIBA_FIELD_MAP = {
    'ContractId': 'Contrato Sap',
    'Contract.ContractName': 'Contrato Legado',
    'Owner.UserName': 'Comprador Estratégico',
    'ContractStatus': 'Estado Contrato Ariba',
    'UF_string11': 'Rut',  # Rut empresa proveedor
    'UF_string10': 'Cód SAP',  # Código acreedor SAP
    'AffectedParties.CommonSupplierName': 'Proveedor',
    'EffectiveDate.Day': 'Fecha Inicio',
    'ExpirationDate.Day': 'Fecha Término Contrato',
    'Description': 'Descripción',
    'Region.RegionNameL2': 'Área',  # Puede ser Área o Región
    'IsEvergreen': 'Contratos Indefinidos',
    'UF_boolean1': 'Aplica Boleta de Garantía (Ariba)',
    'UF_string23': 'Tipo Garantía',
    'UF_time6.Day': 'Vencimiento Garantía',
    'BeginDate.Day': 'Fecha de entrada en vigor - Fecha',  # Backup
}

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
    """Carga el Pivot de Ariba detectando la hoja Data y saltando metadatos."""
    try:
        # Intentar leer hoja 'Data' primero
        if 'Data' in pd.ExcelFile(file_path).sheet_names:
            df_scan = pd.read_excel(file_path, sheet_name='Data', header=None, nrows=50)
        else:
            df_scan = pd.read_excel(file_path, header=None, nrows=50)
    except:
        df_scan = pd.read_excel(file_path, header=None, nrows=50)
    
    # Buscar fila de encabezados reales (donde aparece ContractId o ID de contrato)
    header_row = None
    for i, row in df_scan.iterrows():
        row_str = ' '.join(str(v).lower() for v in row if pd.notna(v))
        if 'contractid' in row_str or 'id de contrato' in row_str:
            header_row = i
            break
    
    if header_row is None:
        # Fallback: usar fila 12 típica de exports de Ariba
        header_row = 12
    
    # Leer con el header detectado
    try:
        df = pd.read_excel(file_path, sheet_name='Data', header=header_row)
    except:
        df = pd.read_excel(file_path, header=header_row)
    
    # Limpiar nombres de columnas
    df.columns = [str(c).strip() for c in df.columns]
    return df

def transform_data(df_pivot: pd.DataFrame) -> tuple:
    """Transforma datos del Pivot Ariba al formato Consolidado."""
    df_out = pd.DataFrame(columns=TARGET_HEADERS)
    
    # 1. Mapeo de campos Ariba → Consolidado
    for ariba_col, target_col in ARIBA_FIELD_MAP.items():
        if ariba_col in df_pivot.columns and target_col in df_out.columns:
            df_out[target_col] = df_pivot[ariba_col].copy()
    
    # Duplicar estado si existe
    if 'Estado Contrato Ariba' in df_out.columns and 'Estado Contrato' in df_out.columns:
        df_out['Estado Contrato'] = df_out['Estado Contrato Ariba']
    
    # 2. 🔥 VALIDACIÓN ESTRICTA DE COMPRADORES
    if 'Owner.UserName' in df_pivot.columns:
        raw_owners = df_pivot['Owner.UserName'].fillna('').astype(str)
        classified = raw_owners.apply(classify_buyer_strict)
        df_out['Comprador Estratégico'] = [x[1] if x[0] == 'strategic' else '' for x in classified]
        df_out['Comprador Táctico'] = [x[1] if x[0] == 'tactical' else '' for x in classified]
    else:
        # Si no hay Owner.UserName, intentar con columna alternativa
        alt_owner = [c for c in df_pivot.columns if 'owner' in c.lower() or 'propietario' in c.lower()]
        if alt_owner:
            raw_owners = df_pivot[alt_owner[0]].fillna('').astype(str)
            classified = raw_owners.apply(classify_buyer_strict)
            df_out['Comprador Estratégico'] = [x[1] if x[0] == 'strategic' else '' for x in classified]
            df_out['Comprador Táctico'] = [x[1] if x[0] == 'tactical' else '' for x in classified]
    
    # 🗑️ Eliminar filas sin compradores válidos
    mask_valid_buyer = (df_out['Comprador Estratégico'] != '') | (df_out['Comprador Táctico'] != '')
    dropped_invalid = (~mask_valid_buyer).sum()
    df_out = df_out[mask_valid_buyer].reset_index(drop=True)
    
    # 🚫 Eliminar contratos Cerrados
    if 'Estado Contrato Ariba' in df_out.columns:
        mask_no_cerrado = ~df_out['Estado Contrato Ariba'].astype(str).str.strip().str.lower().isin(['cerrado', 'cerrados'])
        dropped_cerrados = (~mask_no_cerrado).sum()
        df_out = df_out[mask_no_cerrado].reset_index(drop=True)
    else:
        dropped_cerrados = 0
    
    # 🛡️ Copiar Estratégico a Táctico si está vacío
    mask_tactical_empty = (df_out['Comprador Táctico'] == '') & (df_out['Comprador Estratégico'] != '')
    df_out.loc[mask_tactical_empty, 'Comprador Táctico'] = df_out.loc[mask_tactical_empty, 'Comprador Estratégico']
    
    # Administrador = Estratégico por defecto
    df_out['Administrador de Contrato'] = df_out['Comprador Estratégico'].where(
        df_out['Comprador Estratégico'] != '', df_out['Comprador Táctico']
    )
    
    # 3. Formateo de fechas
    for date_col in ['Fecha Inicio', 'Fecha Término Contrato', 'Vencimiento Garantía']:
        if date_col in df_out.columns:
            df_out[date_col] = pd.to_datetime(df_out[date_col], errors='coerce').dt.strftime('%d-%m-%Y')
            df_out.loc[df_out[date_col] == 'NaT', date_col] = ''
    
    # 4. Normalizar Sí/No
    for col in ['Ingresa a Planta', 'Aplica Boleta de Garantía (Ariba)', 'Aplica Boleta de Garantía (Contrato firmado)', 'Contratos Indefinidos']:
        if col in df_out.columns:
            df_out[col] = df_out[col].astype(str).str.strip().str.title()
            df_out[col] = df_out[col].replace(['Si', 'Sí', 'Yes', 'Y', 'True'], 'Sí')
            df_out[col] = df_out[col].replace(['No', 'N', 'False', 'Nan', ''], 'No')
    
    # 5. Limpieza final
    df_out = df_out.fillna('')
    df_out = df_out.replace(['null', 'None', 'Unclassified', 'nan'], '')
    
    return df_out, dropped_invalid, dropped_cerrados

def apply_formatting(df: pd.DataFrame, output_path: str):
    """Aplica formato Excel idéntico al Consolidado de referencia."""
    wb = openpyxl.Workbook()
    ws = wb.active
    ws.title = "Consolidado de Contratos"
    
    for col_idx, header in enumerate(TARGET_HEADERS, start=1):
        cell = ws.cell(row=1, column=col_idx, value=header)
        cell.fill = STYLES["header"]["fill"]
        cell.font = STYLES["header"]["font"]
        cell.alignment = STYLES["header"]["alignment"]
        cell.border = STYLES["header"]["border"]
    ws.row_dimensions[1].height = 45
    
    for r_idx, row in df.iterrows():
        for c_idx, val in enumerate(row, start=1):
            cell = ws.cell(row=r_idx + 2, column=c_idx, value=val)
            cell.font = STYLES["data"]["font"]
            cell.border = STYLES["data"]["border"]
            cell.alignment = STYLES["data"]["alignment"]
            if c_idx in [9, 10, 23]:
                cell.alignment = Alignment(horizontal='center', vertical='center')
            elif isinstance(val, (int, float)):
                cell.alignment = Alignment(horizontal='center', vertical='center')
                cell.number_format = '#,##0.00'
                
    for i, w in enumerate(COLUMN_WIDTHS, start=1):
        ws.column_dimensions[get_column_letter(i)].width = w
    
    ws.auto_filter.ref = f"A1:{get_column_letter(len(TARGET_HEADERS))}{len(df) + 1}"
    ws.freeze_panes = "C2"
    wb.save(output_path)

# ─────────────────────────────────────────────────────────────
# 🌐 STREAMLIT UI
# ─────────────────────────────────────────────────────────────
st.set_page_config(page_title="Generador Consolidado Ariba", layout="centered")
st.title("📑 Generador de Consolidado de Contratos")
st.caption("Sube el Pivot de Ariba. Se mapearán los campos técnicos correctamente.")

uploaded_file = st.file_uploader("📥 Archivo Pivot (.xlsx)", type=["xlsx"])

if uploaded_file:
    with st.spinner("Leyendo Pivot, mapeando campos y validando compradores..."):
        try:
            with tempfile.NamedTemporaryFile(delete=False, suffix="_pivot.xlsx") as tmp:
                tmp.write(uploaded_file.getvalue())
                pivot_path = tmp.name
            
            df_pivot = load_pivot(pivot_path)
            st.info(f"📊 Pivot cargado: {len(df_pivot)} filas, columnas: {list(df_pivot.columns)[:10]}...")
            
            df_final, dropped_invalid, dropped_cerrados = transform_data(df_pivot)
            
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
            
            st.success(f"✅ **Archivo generado.**\n• 📄 {len(df_final)} contratos procesados.\n• 🗑️ {dropped_invalid} eliminados por compradores no oficiales.\n• 🚫 {dropped_cerrados} eliminados por estado 'Cerrado'.")
            
            os.unlink(pivot_path)
            os.unlink(out_path)
            
        except Exception as e:
            st.error(f"❌ Error: {str(e)}")
            st.exception(e)
