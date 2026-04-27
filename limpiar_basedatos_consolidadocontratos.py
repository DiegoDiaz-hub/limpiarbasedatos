import streamlit as st
import pandas as pd
import openpyxl
from openpyxl.styles import Font, PatternFill, Alignment, Border, Side
from openpyxl.utils import get_column_letter
import tempfile
import os

# ─────────────────────────────────────────────────────────────
# 📐 ESTRUCTURA EXACTA DEL CONSOLIDADO OBJETIVO
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

# 🔄 Mapeo Pivot Ariba → Columnas del Consolidado
# Se corrigió la duplicidad de 'Nombre del propietario'
PIVOT_MAPPING = {
    'ID de contrato': 'Contrato Sap',
    'Nombre del propietario': 'Comprador Estratégico',
    'Estado del contrato': 'Estado Contrato Ariba',
    'Rut empresa proveedor': 'Rut',
    'Código acreedor SAP': 'Cód SAP',
    'Partes afectadas - Proveedor común': 'Proveedor',
    'Fecha de entrada en vigor - Fecha': 'Fecha Inicio',
    'Fecha de expiración - Fecha': 'Fecha Término Contrato',
    'Descripción': 'Descripción',
    'Aplica Garantía': 'Aplica Boleta de Garantía (Ariba)',
    'Es Indefinido': 'Contratos Indefinidos'
}

def read_pivot_safely(file_path):
    """Lee el Pivot detectando automáticamente la fila de encabezados."""
    df_scan = pd.read_excel(file_path, sheet_name='Data', header=None, nrows=50)
    header_row = None
    for i, row in df_scan.iterrows():
        if any('ID de contrato' in str(v) for v in row if pd.notna(v)):
            header_row = i
            break
    if header_row is None:
        raise ValueError("No se encontró la fila de encabezados ('ID de contrato') en la hoja Data.")
        
    df = pd.read_excel(file_path, sheet_name='Data', header=header_row, engine='openpyxl')
    df.columns = [str(c).strip() for c in df.columns]
    return df

def generate_identical_consolidado(pivot_path, output_path):
    """Genera un Excel idéntico en estructura y formato al Consolidado de Contratos."""
    df_pivot = read_pivot_safely(pivot_path)
    
    # Crear DataFrame con estructura vacía exacta
    df_out = pd.DataFrame(columns=TARGET_HEADERS)
    
    # Mapear datos según configuración
    for p_col, t_col in PIVOT_MAPPING.items():
        if p_col in df_pivot.columns and t_col in df_out.columns:
            df_out[t_col] = df_pivot[p_col].copy()
            
    # Mapear Propietario también a Administrador de Contrato
    # (Suelen ser la misma persona en Ariba)
    if 'Nombre del propietario' in df_pivot.columns and 'Administrador de Contrato' in df_out.columns:
        df_out['Administrador de Contrato'] = df_pivot['Nombre del propietario'].copy()
            
    # Duplicar estado en la columna objetivo
    if 'Estado Contrato Ariba' in df_out.columns and 'Estado Contrato' in df_out.columns:
        df_out['Estado Contrato'] = df_out['Estado Contrato Ariba'].copy()
        
    # Limpieza
    df_out = df_out.replace(['null', 'None', 'Unclassified', 'nan'], '')
    df_out = df_out.fillna('')
    
    # Formatear fechas
    for date_col in ['Fecha Inicio', 'Fecha Término Contrato']:
        if date_col in df_out.columns:
            df_out[date_col] = pd.to_datetime(df_out[date_col], errors='coerce').dt.strftime('%d/%m/%Y')
            df_out.loc[df_out[date_col] == 'NaT', date_col] = ''

    # ─────────────────────────────────────────────────────────────
    # 🎨 ESCRITURA EXCEL CON FORMATO IDÉNTICO
    # ─────────────────────────────────────────────────────────────
    wb = openpyxl.Workbook()
    ws = wb.active
    ws.title = "Consolidado de Contratos"
    
    # Estilos
    hdr_fill = PatternFill(start_color="E7E6E6", end_color="E7E6E6", fill_type="solid")
    hdr_font = Font(name='Arial', size=8, bold=True)
    hdr_align = Alignment(horizontal='center', vertical='center', wrap_text=True)
    thin_border = Border(left=Side('thin'), right=Side('thin'), top=Side('thin'), bottom=Side('thin'))
    data_font = Font(name='Arial', size=8)
    data_align = Alignment(horizontal='left', vertical='center', wrap_text=False)
    
    # Escribir encabezados
    for col_idx, header in enumerate(TARGET_HEADERS, start=1):
        cell = ws.cell(row=1, column=col_idx, value=header)
        cell.fill = hdr_fill
        cell.font = hdr_font
        cell.alignment = hdr_align
        cell.border = thin_border
    ws.row_dimensions[1].height = 45
    
    # Escribir datos
    for r_idx, row in df_out.iterrows():
        for c_idx, val in enumerate(row, start=1):
            cell = ws.cell(row=r_idx + 2, column=c_idx, value=val)
            cell.font = data_font
            cell.border = thin_border
            cell.alignment = data_align
            
            # Ajuste especial para fechas y montos
            if c_idx in [9, 10]: # Columnas Fecha
                cell.alignment = Alignment(horizontal='center', vertical='center')
            if isinstance(val, (int, float)):
                cell.alignment = Alignment(horizontal='center', vertical='center')
                cell.number_format = '#,##0.00'

    # Anchos de columna (ajustados para igualar visualmente tu archivo)
    widths = [14, 16, 20, 16, 20, 16, 12, 38, 14, 18, 16, 45, 16, 16, 12, 14, 26, 32, 14, 14, 14, 16, 16, 14, 20, 22, 35, 45, 18, 45]
    for i, w in enumerate(widths, 1):
        ws.column_dimensions[get_column_letter(i)].width = w
        
    # Autoajuste de filtros y paneles
    ws.auto_filter.ref = f"A1:{get_column_letter(len(TARGET_HEADERS))}{len(df_out) + 1}"
    ws.freeze_panes = "C2"
    
    wb.save(output_path)
    return output_path

# ─────────────────────────────────────────────────────────────
# 🌐 INTERFAZ STREAMLIT
# ─────────────────────────────────────────────────────────────
st.set_page_config(page_title="Generador Consolidado", layout="centered")
st.title("📑 Generador de Consolidado de Contratos")
st.caption("Sube el Pivot de Ariba y descarga el Consolidado de Contratos con los compradores correctos.")

uploaded_file = st.file_uploader("📥 Selecciona el archivo Pivot (.xlsx)", type=["xlsx"])

if uploaded_file:
    with st.spinner("Procesando y aplicando formato idéntico..."):
        try:
            # Guardar temporalmente
            with tempfile.NamedTemporaryFile(delete=False, suffix="_pivot.xlsx") as tmp:
                tmp.write(uploaded_file.getvalue())
                pivot_path = tmp.name
                
            out_path = pivot_path.replace("_pivot.xlsx", "_Consolidado_Identico.xlsx")
            generate_identical_consolidado(pivot_path, out_path)
            
            # Descargar
            with open(out_path, "rb") as f:
                st.download_button(
                    label="📥 Descargar Consolidado (Formato Idéntico)",
                    data=f,
                    file_name="Consolidado de Contratos.xlsx",
                    mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
                )
            st.success("✅ Archivo generado con éxito. Los nombres de los compradores se han cargado correctamente.")
            
            # Limpieza
            os.unlink(pivot_path)
            os.unlink(out_path)
            
        except Exception as e:
            st.error(f"❌ Error: {str(e)}")
            st.exception(e)
