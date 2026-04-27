"""
pivot_to_consolidado.py
========================
Toma el archivo Pivot descargado desde SAP Ariba y genera desde cero un archivo
llamado "Consolidado de trabajo.xlsx", replicando la estructura y formato
de la hoja "Info Ariba" del Consolidado original.
 
Uso Streamlit:
    streamlit run app.py
"""

import sys
import os
import tempfile
import pandas as pd
import openpyxl
from openpyxl.styles import Font, PatternFill, Alignment, Border, Side
from openpyxl.utils import get_column_letter
import streamlit as st
from pathlib import Path

# ─────────────────────────────────────────────────────────────
# CONFIGURACIÓN DE COLUMNAS Y ESTILOS
# ─────────────────────────────────────────────────────────────

# Estructura de columnas esperada (Info Ariba)
INFO_ARIBA_HEADERS = [
    None,                                               # col A – siempre vacía
    'ID de contrato',
    'Proyecto - Nombre del proyecto',
    'Fecha de inicio',
    'Nombre del propietario',
    'Código acreedor SAP',
    'Es Indefinido',
    'Región - Región (L2)',
    'Rut empresa proveedor',
    'Partes afectadas - Proveedor común',
    'Contrato - Contrato',
    'Fecha de entrada en vigor - Fecha',
    'Fecha de finalización - Año',
    'Estado del contrato',
    'Fecha de expiración - Fecha',
    'Es un proyecto de prueba',
    'Descripción',
    'Aplica Garantía',
    'Fecha de presentación Garantía N°1 - Fecha',
    'Fecha de termino de notificaciones de garantía - Fecha',
    'N° de Tipos de Garantías',
    'Fecha de termino de notificaciones de garantía - Año',
    'sum(Importe del contrato)',
    'sum(Importe Monto Total Contrato Original)',
    'sum(Importe Monto total Contrato)',
    'Sample',
]

# Estilos
STYLE_HEADER = {
    'fill': PatternFill(start_color="FFE7E6E6", end_color="FFE7E6E6", fill_type="solid"),
    'font': Font(name='Arial', size=8, bold=True),
    'alignment': Alignment(horizontal='center', vertical='center', wrap_text=True),
    'border': Border(
        left=Side(style='thin'),
        right=Side(style='thin'),
        top=Side(style='thin'),
        bottom=Side(style='thin')
    )
}

STYLE_DATA = {
    'font': Font(name='Arial', size=8),
    'border': Border(
        left=Side(style='thin'),
        right=Side(style='thin'),
        top=Side(style='thin'),
        bottom=Side(style='thin')
    )
}


def find_header_row(xl_path: Path) -> int:
    """Encuentra la fila (0-indexed para pandas) donde está 'ID de contrato'."""
    df = pd.read_excel(xl_path, sheet_name='Data', header=None, nrows=30, engine='openpyxl')
    for i, row in df.iterrows():
        # Busca en toda la fila por si hay espacios o variaciones
        if any('ID de contrato' in str(val) for val in row if pd.notna(val)):
            return i
    raise ValueError("No se encontró la fila de encabezados 'ID de contrato' en la hoja Data.")


def read_and_clean_pivot(pivot_path: Path) -> pd.DataFrame:
    """Lee el Pivot, limpia datos y reordena columnas."""
    header_idx = find_header_row(pivot_path)
    
    # Leemos el archivo usando la fila de encabezado encontrada
    df = pd.read_excel(pivot_path, sheet_name='Data', header=header_idx, engine='openpyxl')
    
    # Eliminamos filas y columnas completamente vacías
    df = df.dropna(how='all').dropna(axis=1, how='all')
    
    # Normalizamos los nombres de las columnas para asegurar el match
    # A veces Ariba pone espacios extra
    df.columns = [col.strip() if isinstance(col, str) else col for col in df.columns]
    
    # Reordenamos columnas para que coincidan con INFO_ARIBA_HEADERS
    # Las columnas que no están en la lista se ignoran, las que faltan se dejan vacías
    final_columns = [col for col in INFO_ARIBA_HEADERS if col is not None]
    df_final = pd.DataFrame()
    
    # Aseguramos que la columna vacía (A) exista
    df_final[None] = None 
    
    for col_name in final_columns:
        if col_name in df.columns:
            df_final[col_name] = df[col_name]
        else:
            df_final[col_name] = None # Columna vacía si falta en el pivot

    return df_final


def create_consolidado_workbook(df: pd.DataFrame) -> str:
    """Crea un archivo Excel nuevo con formato de Consolidado."""
    wb = openpyxl.Workbook()
    ws = wb.active
    ws.title = "Info Ariba"

    # 1. Escribir Encabezados
    for col_idx, header in enumerate(INFO_ARIBA_HEADERS, start=1):
        cell = ws.cell(row=1, column=col_idx)
        cell.value = header
        cell.fill = STYLE_HEADER['fill']
        cell.font = STYLE_HEADER['font']
        cell.alignment = STYLE_HEADER['alignment']
        cell.border = STYLE_HEADER['border']

    # Altura de la fila de encabezado para que se vea el texto
    ws.row_dimensions[1].height = 45

    # 2. Escribir Datos
    # pivot_cols son los nombres de las columnas (excluyendo None)
    pivot_cols = [h for h in INFO_ARIBA_HEADERS if h is not None]

    for r_idx, (_, row_data) in enumerate(df.iterrows(), start=2):
        # Columna A (vacía)
        ws.cell(row=r_idx, column=1).value = None
        ws.cell(row=r_idx, column=1).border = STYLE_DATA['border']
        
        # Resto de columnas
        for c_idx, col_name in enumerate(pivot_cols, start=2):
            val = row_data.get(col_name, None)
            
            # Limpieza de NaN / NaT
            if pd.isna(val):
                val = None
            
            cell = ws.cell(row=r_idx, column=c_idx)
            cell.value = val
            cell.font = STYLE_DATA['font']
            cell.border = STYLE_DATA['border']
            # Centramos texto numérico y fechas, alineamos a la izquierda texto largo
            if isinstance(val, (int, float)):
                cell.alignment = Alignment(horizontal='center', vertical='center')
            elif isinstance(val, pd.Timestamp):
                cell.alignment = Alignment(horizontal='center', vertical='center')
                cell.number_format = 'DD/MM/YYYY'

    # 3. Filtros y Paneles
    max_row = ws.max_row
    max_col = ws.max_column
    
    # AutoFilter en toda la tabla
    ws.auto_filter.ref = f"A1:{get_column_letter(max_col)}{max_row}"
    
    # Inmovilizar paneles (similar a 'E3' en el original, ajustado a la estructura)
    ws.freeze_panes = "B2"

    # 4. Ajustar ancho de columnas (Opcional, pero ayuda a que se vea bien)
    for column_cells in ws.columns:
        max_length = 0
        column = column_cells[0].column_letter
        for cell in column_cells:
            try:
                if cell.value:
                    max_length = max(max_length, len(str(cell.value)))
            except:
                pass
        adjusted_width = min(max_length + 2, 40) # Tope de 40
        ws.column_dimensions[column].width = adjusted_width

    # Guardar temporalmente para retornar la ruta
    with tempfile.NamedTemporaryFile(delete=False, suffix=".xlsx", prefix="Consolidado_Trabajo_") as tmp_file:
        wb.save(tmp_file.name)
        return tmp_file.name


# ─────────────────────────────────────────────────────────────
# INTERFAZ STREAMLIT
# ─────────────────────────────────────────────────────────────

st.set_page_config(page_title="Generador Consolidado Ariba", layout="centered")
st.title("📑 Generador de Consolidado de Trabajo")
st.caption("Sube el archivo Pivot de Ariba y descarga el archivo formateado.")

uploaded_file = st.file_uploader("Selecciona el archivo Pivot (.xlsx)", type=["xlsx"])

if uploaded_file:
    st.divider()
    col1, col2 = st.columns([2, 1])
    
    with col1:
        st.info(f"Archivo cargado: **{uploaded_file.name}**")
    
    with col2:
        process_btn = st.button("🚀 Generar Archivo", type="primary", use_container_width=True)

    if process_btn:
        try:
            with st.spinner("Procesando datos y aplicando formato..."):
                # 1. Guardar upload temporal
                with tempfile.NamedTemporaryFile(delete=False, suffix=".xlsx") as tmp_in:
                    tmp_in.write(uploaded_file.getvalue())
                    pivot_path = Path(tmp_in.name)

                # 2. Leer y limpiar
                df_clean = read_and_clean_pivot(pivot_path)
                st.success(f"✅ Se encontraron {len(df_clean)} registros válidos.")

                # 3. Generar archivo final
                output_path = create_consolidado_workbook(df_clean)

                # 4. Descargar
                with open(output_path, "rb") as f:
                    st.download_button(
                        label="📥 Descargar Consolidado de trabajo.xlsx",
                        data=f,
                        file_name="Consolidado de trabajo.xlsx",
                        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
                    )
                
                # Limpieza de temporales
                os.unlink(pivot_path)
                os.unlink(output_path)

        except Exception as e:
            st.error(f"❌ Ocurrió un error: {str(e)}")
            st.exception(e)
