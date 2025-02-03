import datetime
import os
import pandas as pd
from docxtpl import DocxTemplate

def fill_word_template(excel_path, word_template_path, word_output_path):
  
    """
    Llena un plantilla de word con datos extraidos de un archivo excel, utilizando jinja2.

    Args:

    excel_path (str) : ruta del archivo de Excel
    word_template_path (str) : Ruta de la plantilla de word
    word_outh_path (str) : Ruta donde se guardara el documento word final.

    """
    df_tabla_resumen = pd.read_excel(excel_path, sheet_name="tabla_resumen", dtype=str, engine="openpyxl")
    df_reporte_detalle = pd.read_excel(excel_path, sheet_name="reporte_detalle", dtype=str, engine="openpyxl")

    df_tabla_resumen.columns = df_tabla_resumen.columns.str.strip()
    df_reporte_detalle.columns = df_reporte_detalle.columns.str.strip()

    tabla_resumen = []
    for _, row in df_tabla_resumen.iterrows():
    
        resumen_info = {
            'ticket': row.get('Ticket', ''),
            'tipo_generacion': row.get('Tipo de generacion de ticket', ''),
            'fecha_interrupcion': row.get('Fecha/Hora Inturrepcion', ''),
            'fecha_solicitud': row.get('Fecha/Hora Solicitud', ''),
            'fecha_generacion': row.get('Fecha/Hora Generación', ''),
            'fecha_llegada': row.get('Fecha/Hora Llegada de Personal', ''),
            'tiempo_llegada': row.get('Tiempo de llegada del Personal (Hrs)', ''),
            'fecha_subsanacion': row.get('Fecha/Hora Subsanación', ''),
            'cid': row.get('CID', ''),
            'tipo_caso': row.get('Tipo Caso', ''),
            'averia': row.get('Avería', ''),
            'tiempo_indisponibilidad': row.get('Tiempo de Indisponibilidad - Ver Nota 1 (Hrs)', ''),
            'tiempo_subsanacion': row.get('Tiempo de subsanación efectivo - Ver Nota 2 (Hrs)', ''),
            'horas_excedidas': row.get('Horas excedidas en el plazo de reparación de acuerdo a bases (Hrs)', '')
        }
        tabla_resumen.append(resumen_info)

    reportes_detallados = []
    for _, row in df_reporte_detalle.iterrows():
        reporte = {
            'nro_ticket': row.get('nro_incidencia', ''),
            'fecha_gene': row.get('fecha_generacion', ''),
            'tipo_servicio': row.get('tipo_servicio', ''),
            'cid': row.get('cid', ''),
            'tipo_caso': row.get('tipo_caso', ''),
            #'descripcion_problema': row.get('it_descripcion_problema', ''),
            'determinacion_causa': row.get('it_determinacion_de_la_causa', ''),
            'medidas_tomadas': row.get('it_medidas_tomadas', ''),
            'conclusiones': row.get('it_conclusiones', '')
        }
        reportes_detallados.append(reporte)

    doc = DocxTemplate(word_template_path)
    context  = {

        'fecha_inicio': "01/01/2025",
        'fecha_fin': "31/01/2025",

        'registros_resumen' : tabla_resumen,

        'reportes': reportes_detallados

        }

    doc.render(context)
    output_file = os.path.join(word_output_path, f"reporte_mensual_completo.docx")
    doc.save(output_file)
    print(f"Reporte mensual generado exitosamente: {output_file}")

if __name__ == "__main__":

    #base_path = "C:/Users/katana/Desktop/proyectos/bots_rpa"
    base_path = "C:/Users/mozac/Documents/pruebas/Fast-API-Bots-RPA-python-"

    excel_path = os.path.join(base_path, "media", "pronatel", "data", "DataTicketsPronatel.xlsx")
    word_template_path = os.path.join(base_path, "media", "pronatel", "plantillas", "plantilla_word.docx")
    word_output_path = os.path.join(base_path, "media", "pronatel", "reportes/")

    fill_word_template(excel_path, word_template_path, word_output_path)


