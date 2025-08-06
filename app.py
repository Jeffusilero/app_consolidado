import streamlit as st
import pandas as pd
from datetime import datetime
from fpdf import FPDF
import io
import locale

# Configuración robusta de locale para nombres de meses en español
def configure_locale():
    locale_options = [
        'es_ES.UTF-8', 
        'es_ES.utf8',
        'es_ES',
        'spanish',
        'es_CO.UTF-8',
        'es_MX.UTF-8',
        'es_AR.UTF-8'
    ]
    
    for loc in locale_options:
        try:
            locale.setlocale(locale.LC_TIME, loc)
            return True
        except (locale.Error, Exception):
            continue
    
    try:
        # Último intento con locale neutral
        locale.setlocale(locale.LC_TIME, 'C')
        st.warning("No se pudo configurar el locale en español. Se usará formato neutral.")
        return False
    except:
        st.error("Error crítico al configurar locale")
        return False

# Ejecutar la configuración de locale
configure_locale()

class PDF(FPDF):
    def __init__(self, orientation='P', unit='mm', format='A4'):
        super().__init__(orientation=orientation, unit=unit, format=format)
        self._has_header_been_shown = False

    def header(self):
        if not self._has_header_been_shown:
            self.set_font('Arial', 'B', 17)
            self.cell(0, 5, 'LDC LOGISTICA ECUADOR S.A.S.', 0, 1, 'C')
            self.ln(3)
            self._has_header_been_shown = True
        
    def footer(self):
        self.set_y(-10)
        self.set_font('Arial', 'I', 6)
        self.cell(0, 4, f'Página {self.page_no()}', 0, 0, 'C')

def generar_pdf(fecha, consolidado, hora, df, total_paquetes):
    pdf = PDF(orientation='P', unit='mm', format='A4')
    pdf.add_page()
    
    # --- ENCABEZADO SUPERIOR ---
    pdf.set_font('Arial', 'B', 10)
    fecha_x, fecha_y = 11, 17
    consolidado_x, consolidado_y = 11, 22
    hora_x, hora_y = 110, 17
    paquetes_x, paquetes_y = 110, 22
    
    # Formatear fecha según locale configurado
    try:
        fecha_str = fecha.strftime("%d de %B de %Y")
    except:
        # Fallback si hay problemas con el locale
        fecha_str = fecha.strftime("%d/%m/%Y")
    
    pdf.set_xy(fecha_x, fecha_y)
    pdf.cell(40, 5, f'Fecha: {fecha_str}', 0, 0, 'L')
    pdf.set_xy(consolidado_x, consolidado_y)
    pdf.cell(40, 5, f'Consolidado: {consolidado}', 0, 0, 'L')
    pdf.set_xy(hora_x, hora_y)
    pdf.cell(40, 5, f'Hora: {hora}', 0, 0, 'L')
    pdf.set_xy(paquetes_x, paquetes_y)
    pdf.cell(40, 5, f'Paquetes entregados: {total_paquetes}', 0, 0, 'L')
    pdf.ln(10)

    # --- TABLAS ---
    col_width = 23
    row_height = 3
    max_rows_per_page = 88
    start_y = 28
    
    total_rows = len(df)
    current_row = 0
    last_left_table_bottom = 0

    while current_row < total_rows:
        if current_row > 0:
            pdf.add_page()
            start_y = 20
        
        remaining_rows = total_rows - current_row
        is_last_page_with_few_rows = (remaining_rows <= max_rows_per_page) and (remaining_rows > max_rows_per_page//2)
        
        # --- TABLA IZQUIERDA ---
        if is_last_page_with_few_rows:
            left_rows = remaining_rows // 2
            right_rows = remaining_rows - left_rows
        else:
            available_rows = min(max_rows_per_page, remaining_rows)
            left_rows = int(min(available_rows, (pdf.h - start_y - 40) // row_height - 1))
        
        if left_rows > 0:
            pdf.set_xy(11, start_y)
            pdf.set_font('Arial', 'B', 7)
            pdf.set_fill_color(200, 200, 200)
            pdf.set_text_color(255, 255, 255)
            pdf.cell(col_width, row_height, 'GUIA DORADOS', 1, 0, 'C', 1)
            pdf.cell(col_width, row_height, 'GUIAS TRAMACO', 1, 0, 'C', 1)
            pdf.cell(col_width, row_height, 'MASTER', 1, 0, 'C', 1)
            pdf.cell(col_width, row_height, 'SACO', 1, 1, 'C', 1)
            pdf.set_text_color(0, 0, 0)
            
            pdf.set_font('Arial', '', 5)
            for i in range(current_row, current_row + left_rows):
                pdf.set_x(11)
                row = df.iloc[i]
                for col in range(3):  # Primeras 3 columnas normales
                    pdf.cell(col_width, row_height, str(row.iloc[col]) if pd.notna(row.iloc[col]) else '', 1, 0, 'C')
                # Cuarta columna (SACO) con formato especial
                saco_value = str(row.iloc[3]) if pd.notna(row.iloc[3]) and str(row.iloc[3]) != '' else 'AMBATO'
                if saco_value.replace('.', '').isdigit():  # Si es numérico (incluye decimales)
                    saco_value = str(int(float(saco_value)))  # Convertir a entero
                pdf.cell(col_width, row_height, saco_value, 1, 0, 'C')
                pdf.ln(row_height)
            
            last_left_table_bottom = pdf.get_y()
            current_row += left_rows
        
        # --- TABLA DERECHA ---
        if is_last_page_with_few_rows:
            pass
        else:
            right_rows = 0
            if current_row < total_rows and pdf.get_y() < pdf.h - 40:
                right_rows = int(min(total_rows - current_row, (pdf.h - start_y - 40) // row_height - 1))
        
        if right_rows > 0:
            pdf.set_xy(110, start_y)
            pdf.set_font('Arial', 'B', 7)
            pdf.set_fill_color(200, 200, 200)
            pdf.set_text_color(255, 255, 255)
            pdf.cell(col_width, row_height, 'GUIA DORADOS', 1, 0, 'C', 1)
            pdf.cell(col_width, row_height, 'GUIAS TRAMACO', 1, 0, 'C', 1)
            pdf.cell(col_width, row_height, 'MASTER', 1, 0, 'C', 1)
            pdf.cell(col_width, row_height, 'SACO', 1, 1, 'C', 1)
            pdf.set_text_color(0, 0, 0)
            
            pdf.set_font('Arial', '', 5)
            for i in range(current_row, current_row + right_rows):
                pdf.set_x(110)
                row = df.iloc[i]
                for col in range(3):  # Primeras 3 columnas normales
                    pdf.cell(col_width, row_height, str(row.iloc[col]) if pd.notna(row.iloc[col]) else '', 1, 0, 'C')
                # Cuarta columna (SACO) con formato especial
                saco_value = str(row.iloc[3]) if pd.notna(row.iloc[3]) and str(row.iloc[3]) != '' else 'AMBATO'
                if saco_value.replace('.', '').isdigit():  # Si es numérico (incluye decimales)
                    saco_value = str(int(float(saco_value)))  # Convertir a entero
                pdf.cell(col_width, row_height, saco_value, 1, 0, 'C')
                pdf.ln(row_height)
            
            current_row += right_rows
    
    # --- FIRMAS ---
    firmas_y = last_left_table_bottom + 50
    if firmas_y > pdf.h - 30:
        pdf.add_page()
        firmas_y = 20
    
    line_length = 50
    left_x = 40
    right_x = 120
    
    # Centrar texto de firma con la línea
    pdf.set_font('Arial', '', 10)
    text_width = pdf.get_string_width('Yixon Gonzalez Diaz')
    centered_x = left_x + (line_length - text_width) / 2
    pdf.set_xy(centered_x, firmas_y)
    pdf.cell(text_width, 5, 'Yixon Diaz', 0, 0, 'L')
    pdf.line(left_x, firmas_y + 7, left_x + line_length, firmas_y + 7)
    
    pdf.set_font('Arial', 'B', 12)
    text_width = pdf.get_string_width('DESPACHADO POR')
    centered_x = left_x + (line_length - text_width) / 2
    pdf.set_xy(centered_x, firmas_y + 9)
    pdf.cell(text_width, 5, 'DESPACHADO POR', 0, 0, 'L')
    
    pdf.line(right_x, firmas_y + 7, right_x + line_length, firmas_y + 7)
    pdf.set_font('Arial', 'B', 12)
    text_width = pdf.get_string_width('TRANSPORTISTA')
    centered_x = right_x + (line_length - text_width) / 2
    pdf.set_xy(centered_x, firmas_y + 9)
    pdf.cell(text_width, 5, 'TRANSPORTISTA', 0, 0, 'L')
    
    return pdf

def main():
    st.set_page_config(page_title="Consolidado de Guías", layout="wide")
    st.title("LDC LOGISTICA ECUADOR S.A.S. - Consolidado de Guías")
    
    with st.container():
        col1, col2 = st.columns(2)
        with col1:
            fecha = st.date_input("Fecha:", value=datetime.now().date())
        with col2:
            consolidado = st.text_input("Consolidado No:", value="")
    
    hora = st.text_input("Hora:", value="17:00")
    
    st.subheader("Archivos Base (puede subir varios)")
    base_files = st.file_uploader("Subir archivos Excel base (con columnas A, B, C)", 
                                type=["xlsx"], 
                                accept_multiple_files=True,
                                key="base_files")
    
    st.subheader("Archivo de Comparación (solo uno)")
    comp_file = st.file_uploader("Subir archivo Excel de comparación (con columnas H y M)", 
                               type=["xlsx"], 
                               accept_multiple_files=False,
                               key="comp_file")
    
    if base_files and comp_file:
    # Procesar archivo de comparación
    try:
        df_comp = pd.read_excel(comp_file, header=None, dtype=str)
        
        # Crear diccionario de coincidencias M -> H
        saco_mapping = {}
        for _, row in df_comp.iterrows():
            if len(row) >= 13:  # Verificar que tenga columna M (índice 12)
                if pd.notna(row[12]):  # Columna M
                    saco_mapping[row[12]] = row[7] if (len(row) > 7 and pd.notna(row[7])) else ''
        
        # Combinar todos los archivos base en un solo DataFrame
        combined_df = pd.DataFrame()
        
        for base_file in base_files:
            df_base = pd.read_excel(base_file, header=None, dtype=str)
            df_base = df_base.iloc[:, :3]  # Tomar primeras 3 columnas
            df_base[3] = ''  # Columna SACO vacía
            combined_df = pd.concat([combined_df, df_base], ignore_index=True)
        
        # Aplicar mapeo SACO al dataframe combinado
        for idx, row in combined_df.iterrows():
            guia = str(row[0]) if pd.notna(row[0]) else ''
            
            # Caso 1: Guías que son exactamente "SACO" -> "AMBATO"
            if guia == 'SACO':
                combined_df.at[idx, 3] = 'AMBATO'
            # Caso 2: Guías que comienzan con "QU" -> "QUITO" (si no están en el mapeo)
            elif guia.startswith('QU'):
                combined_df.at[idx, 3] = saco_mapping.get(guia, 'QUITO')
            # Caso 3: Cualquier otro caso -> "AMBATO"
            else:
                combined_df.at[idx, 3] = saco_mapping.get(guia, 'AMBATO')
        
        # ORDENAR POR COLUMNA SACO (columna 3)
        # Convertir a numérico (los no numéricos se convierten en NaN)
        combined_df[3] = pd.to_numeric(combined_df[3], errors='coerce')
        # Ordenar (los NaN van primero)
        combined_df = combined_df.sort_values(by=3, na_position='first')
        # Volver a string y limpiar NaN
        combined_df[3] = combined_df[3].astype(str)
        combined_df = combined_df.replace('nan', '')
            
            total_paquetes = len(combined_df)
            
            # Mostrar vista previa del DataFrame combinado
            with st.expander("Vista previa de todos los datos combinados (ordenados por SACO)"):
                st.dataframe(combined_df.head())
            
            # Generar un solo PDF con todos los datos
            pdf = generar_pdf(fecha, consolidado, hora, combined_df, total_paquetes)
            pdf_bytes = pdf.output(dest='S').encode('latin-1')
            
            # Botón de descarga directa del PDF único
            st.download_button(
                label="📥 Descargar PDF Consolidado",
                data=pdf_bytes,
                file_name=f"Consolidado_{consolidado}_{fecha.strftime('%d-%m-%Y')}.pdf",
                mime="application/pdf",
                use_container_width=True
            )
        
        except Exception as e:
            st.error(f"Error al procesar los archivos: {str(e)}")

if __name__ == "__main__":
    main()


