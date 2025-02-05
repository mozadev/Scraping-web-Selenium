import datetime
import os
from typing import List
import pandas as pd
from docxtpl import DocxTemplate
from datetime import datetime

def validate_required_colums(df: pd.DataFrame, required_columns: List[str]):
    """
    Validates that all required columns exist in the DataFrame.
    Raises a ValueError with detailed information if any columns are missing.
    
    Args:
        df: pandas DataFrame to validate
        required_columns: List of column names that must exist in the DataFrame
    """
    missing_columns = [col for col in required_columns if col not in df.columns]
    if missing_columns:
        error_message = (
            f"Error 400: Missing required columns in Excel file. \n"
            f"Missing columns: {missing_columns}\n"
            f"Please check column names in your Excel file match exactly"
        )
        raise ValueError(error_message)
    

def limpiar_texto(texto):
    """Limpia caracteres especiales como _x000D_ y espacios innecesarios"""
    if isinstance(texto, str): 
        return texto.replace("\r", "").replace("_x000D_", "").strip()
    return texto

def fill_word_template(excel_path, word_template_path, word_output_path):
    
    """
    Llena un plantilla de word con datos extraidos de un archivo excel, utilizando jinja2.
    Incluye validacion de columnas de excel requeridas.
    Args:

    excel_path (str) : ruta del archivo de Excel
    word_template_path (str) : Ruta de la plantilla de word
    word_outh_path (str) : Ruta donde se guardara el documento word final.

    """

    required_columns = [
        'nro_incidencia',
        'fecha_generacion',
        'tipo_servicio',
        'cid',
        'it_determinacion_de_la_causa',
        'it_medidas_tomadas',
        'it_conclusiones',
        'tipo_generacion_ticket',
        'fecha_hora_interrupcion',
        'fecha_hora_solicitud',
        'fecha_hora_llegada_personal',
        'tiempo_llegada_personal',
        'fecha_hora_subsanacion',
        'tiempo_indisponibilidad_hr',
        'tiempo_subsanacion_efectivo_hr',
        'horas_excedidas_plazo_reparacion',
        'fecha_hora_inicio_averia',
        'fecha_hora_fin_averia',
        'descripcion_problema',
        'mejoras'
    ]
    
    try:
        df_reporte_detalle = pd.read_excel(excel_path, sheet_name="reporte_detalle", dtype=str, engine="openpyxl")
        df_reporte_detalle.columns = df_reporte_detalle.columns.str.strip()
        df_reporte_detalle = df_reporte_detalle.map(limpiar_texto)
    except Exception as e:
        raise ValueError(f"Error 400: Could not read Excel file: {str(e)}")
    
    validate_required_colums(df_reporte_detalle,required_columns)

    df_reporte_detalle.fillna(" ", inplace=True)
    reportes_detallados = []
    for _, row in df_reporte_detalle.iterrows():
        reporte = {
            'nro_incidencia': row.get('nro_incidencia',''), #ticket
            'fecha_generacion': row.get('fecha_generacion',''),
            'tipo_servicio': row.get('tipo_servicio', ''),
            'cid': row.get('cid', ''),
            'tipo_caso': row.get('tipo_caso', ''),
            'tipificacion_problema': row.get('tipificacion_problema',''), #averia
            'it_determinacion_de_la_causa': row.get('it_determinacion_de_la_causa', ''),
            'it_medidas_tomadas': row.get('it_medidas_tomadas', ''),
            'it_conclusiones': row.get('it_conclusiones', ''),#recomendaciones
            'tipo_generacion_ticket':row.get('tipo_generacion_ticket', ''),
            'fecha_hora_interrupcion':row.get('fecha_hora_interrupcion', ''),
            'fecha_hora_solicitud':row.get('fecha_hora_solicitud', 'no data'),
            'fecha_hora_llegada_personal':row.get('fecha_hora_llegada_personal', ''),
            'tiempo_llegada_personal':row.get('tiempo_llegada_personal', ''),
            'fecha_hora_subsanacion': row.get('fecha_hora_subsanacion', ''),
            'tiempo_indisponibilidad_hr':row.get('tiempo_indisponibilidad_hr', ''),
            'tiempo_subsanacion_efectivo_hr':row.get('tiempo_subsanacion_efectivo_hr', ''),
            'horas_excedidas_plazo_reparacion':row.get('horas_excedidas_plazo_reparacion', ''),
            'fecha_hora_inicio_averia': row.get('fecha_hora_inicio_averia', ''),
            'fecha_hora_fin_averia': row.get('fecha_hora_fin_averia', ''),
            'descripcion_problema': row.get('descripcion_problema', ''),
            'mejoras': row.get('mejoras', ''),
        }
        reportes_detallados.append(reporte)

    doc = DocxTemplate(word_template_path)
    context  = {
        'fecha_ini': "01/01/2025",
        'fecha_fin': "31/01/2025",
        'reportes': reportes_detallados
        }
    output_file = os.path.join(word_output_path, f"reporte_{datetime.now().strftime('%Y%m%d_%H%M%S')}.docx")
    doc.render(context)
    doc.save(output_file)
    print(f"Reporte mensual generado exitosamente: {output_file}")
    return output_file
