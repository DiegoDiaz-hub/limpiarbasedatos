import streamlit as st
import pandas as pd
import openpyxl
from openpyxl.styles import Font, PatternFill, Alignment, Border, Side
from openpyxl.utils import get_column_letter
import tempfile
import os
import re

# ─────────────────────────────────────────────────────────────
# 📋 LISTA OFICIAL DE COMPRADORES ESTRATÉGICOS
# Actualiza esta lista si se incorporan nuevos compradores reales.
# ─────────────────────────────────────────────────────────────
VALID_STRATEGIC_BUYERS = {
    "jorge urrutia", "juan figueroa", "bárbara garcía", "claudio berrios",
    "viviana grandón", "joseph españa", "patricio espinoza", "dayana dávila",
    "michelle esperanza", "diego escalona", "magdalena farias", "leandro medina",
    "judith rivas", "daniela escobar", "lina diaz", "victor camilla",
    "laura mendoza", "sofia delgado", "priscilla gre guerra", "valeria silva",
    "denisse andrea gonzalez terrile", "carol reyes", "servio salges cazorla",
    "martina fuentes", "claudia castillo", "angela gallardo", "cecilia fernandez",
    "denisse lopez", "felipe pavez correa"
}

def is_valid_strategic_buyer(name: str) -> bool:
    """Valida si el nombre pertenece a la lista oficial de compradores estratégicos."""
    if pd.isna(name): return False
    return str(name).strip().lower() in VALID_STRATEGIC_BUYERS

# ─────────────────────────────────────────────────────────────
# 🎨 ESTILOS DEL CONSOLIDADO
# ─────────────────────────────────────────────────────────────
STYLES = {
    "header": {
        "fill": PatternFill(start_color="FFE7E6E6", end_color="FFE7E6E6", fill_type="solid"),
        "font": Font(name='Arial', size=8, bold=True),
        "alignment": Alignment(horizontal='center', vertical='center', wrap_text=True),
        "border": Border(left=Side('thin'), right=Side('thin'), top=Side('thin'), bottom=Side('thin'))
    },
    "data": {
        "font": Font(name='Arial', size=8),
        "border": Border(left=Side('thin'), right=Side('thin'), top=Side('thin'), bottom=Side('thin')),
        "alignment": Alignment(horizontal='left', vertical='center')
    }
}

# ─────────────────────────────────────────────────────────────
# 🔄 PROCESAMIENTO PRINCIPAL
# ─────────────────────────────────────────────────────────────
def process_pivot(pivot_df: pd.DataFrame) -> pd.DataFrame:
    """Limpia, valida compradores y mapea columnas al formato Consolidado."""
    # 1. Limpieza básica
    df = pivot_df.copy()
    df = df.dropna(how='all').dropna(axis=1, how='all')
    
    # Normalizar nombres de columnas
    df.columns = [col.strip() for col in df.columns]
    
    # 2. Mapeo al formato Consolidado
    consolidado_cols = [
        'Contrato Sap', 'Contrato Legado', 'Comprador Estratégico', 'Comprador Táctico',
        'Estado Contrato Ariba', 'Rut', 'Cód SAP', 'Proveedor', 'Fecha Inicio',
        'Fecha Término Contrato', 'Estado Contrato', 'Descripción', 'Área', 'Gerencia',
        'Planta', 'Ingresa a Planta', 'Aplica Boleta de Garantía (Ariba)',
        'Aplica Boleta de Garantía (Contrato firmado)', 'Tipo Garantía', 'N° Garantia',
        'Moneda Garantía', 'Monto Garantía', 'Vencimiento Garantía', 'Estado Garantía',
        'Administrador de Contrato', 'Correo Electrónico',
        'Observación contrato Control Contratista',
        'Observación Control Contratistas Boleta de Garantía',
        'Contratos Indefinidos', 'Observación Interna'
    ]
    
    res = pd.DataFrame(columns=consolidado_cols)
    
    # Mapeo directo
    res['Contrato Sap'] = df.get('ID de contrato', pd.Series(dtype='object'))
    res['Estado Contrato Ariba'] = df.get('Estado del contrato', pd.Series(dtype='object'))
    res['Rut'] = df.get('Rut empresa proveedor', pd.Series(dtype='object'))
    res['Cód SAP'] = df.get('Código acreedor SAP', pd.Series(dtype='object'))
    res['Proveedor'] = df.get('Partes afectadas - Proveedor común', pd.Series(dtype='object'))
    res['Descripción'] = df.get('Descripción', pd.Series(dtype='object'))
    res['Contratos Indefinidos'] = df.get('Es Indefinido', pd.Series(dtype='object'))
    
    # Fechas (formateo básico)
    for col_in, col_out in [('Fecha de entrada en vigor - Fecha', 'Fecha Inicio'), 
                            ('Fecha de expiración - Fecha', 'Fecha Término Contrato')]:
        if col_in in df.columns:
            res[col_out] = pd.to_datetime(df[col_in], errors='coerce').dt.strftime('%d/%m/%Y')
        else:
            res[col_out] = pd.Series(dtype='object')
            
    # 🛡️ VALIDACIÓN DE COMPRADOR ESTRATÉGICO
    owners = df.get('Nombre del propietario', pd.Series(dtype='object'))
    valid_buyers = owners.apply(is_valid_strategic_buyer)
    
    res['Comprador Estratégico'] = owners.where(valid_buyers, None)
    
    # Si no es estratégico, dejar rastro en observación para auditoría
    res['Observación Interna'] = res['Observación Interna'].astype(str) + " | " + owners[~valid_buyers].fillna("N/A")
    res.loc[res['Observación Interna'].str.len() < 4, 'Observación Interna'] = None
    
    # Estados y garantías
    res['Aplica Boleta de Garantía (Ariba)'] = df.get('Aplica Garantía', pd.Series(dtype='object'))
    res['Administrador de Contrato'] = owners  # Propietario Ariba va como Administrador
    
    # Limpiar NaN/NaT finales
    res = res.fillna('')
    return res

def create_excel_output(df: pd.DataFrame, output_path: str):
    """Escribe el DataFrame en Excel con el formato exacto del Consolidado."""
    wb = openpyxl.Workbook()
    ws = wb.active
    ws.title = "Consolidado de Contratos"
    
    # Headers
    for col_idx, header in enumerate(df.columns, start=1):
        cell = ws.cell(row=1, column=col_idx, value=header)
        for k, v in STYLES["header"].items():
            setattr(cell, k, v)
    ws.row_dimensions[1].height = 40
    
    # Data
    for row_idx, row in df.iterrows():
        for col_idx, value in enumerate(row, start=1):
            cell = ws.cell(row=row_idx + 1, column=col_idx, value=value)
            cell.font = STYLES["data"]["font"]
            cell.border = STYLES["data"]["border"]
            if isinstance(value, (int, float)):
                cell.alignment = Alignment(horizontal='center', vertical='center')
                cell.number_format = '#,##0.00' if value > 100 else 'General'
                
    # Filtros y paneles
    ws.auto_filter.ref = f"A1:{get_column_letter(len(df.columns))}{len(df) + 1}"
    ws.freeze_panes = "C2"
    
    # Anchos automáticos (limitados)
    for col_cells in ws.columns:
        max_len = max((len(str(cell.value)) for cell in col_cells), default=10)
        ws.column_dimensions[get_column_letter(col_cells[0].column)].width = min(max_len + 2, 35)
        
    wb.save(output_path)

# ─────────────────────────────────────────────────────────────
# 🌐 INTERFAZ STREAMLIT
# ─────────────────────────────────────────────────────────────
st.set_page_config(page_title="Consolidador Ariba → Trabajo", layout="centered")
st.title("📊 Generador de Consolidado de Trabajo")
st.caption("Sube solo el Pivot de Ariba. El sistema validará compradores estratégicos y aplicará el formato oficial.")

uploaded = st.file_uploader("📂 Archivo Pivot (.xlsx)", type=["xlsx"])

if uploaded:
    with st.spinner("Procesando y validando compradores estratégicos..."):
        try:
            # Leer Pivot
            pivot_df = pd.read_excel(uploaded, sheet_name='Data', header=None, nrows=50, engine='openpyxl')
            # Detectar fila de headers dinámicamente
            header_idx = 0
            for i, row in pivot_df.iterrows():
                if any('ID de contrato' in str(v) for v in row if pd.notna(v)):
                    header_idx = i
                    break
                    
            pivot_df = pd.read_excel(uploaded, sheet_name='Data', header=header_idx, engine='openpyxl')
            pivot_df = pivot_df.dropna(how='all').dropna(axis=1, how='all')
            
            # Validar y mapear
            final_df = process_pivot(pivot_df)
            
            # Guardar temporal
            with tempfile.NamedTemporaryFile(delete=False, suffix="_Consolidado_Trabajo.xlsx") as tmp:
                create_excel_output(final_df, tmp.name)
                tmp_path = tmp.name
                
            st.success(f"✅ Procesado {len(final_df)} contratos. Compradores estratégicos validados correctamente.")
            
            # Botón descarga
            with open(tmp_path, "rb") as f:
                st.download_button(
                    label="📥 Descargar Consolidado de trabajo.xlsx",
                    data=f,
                    file_name="Consolidado de trabajo.xlsx",
                    mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
                )
                
            # Limpieza
            os.unlink(tmp_path)
            
        except Exception as e:
            st.error(f"❌ Error: {str(e)}")
