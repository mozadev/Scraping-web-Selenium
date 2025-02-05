import datetime
import os
import pandas as pd
from docxtpl import DocxTemplate

def limpiar_texto(texto):
    """Limpia caracteres especiales como _x000D_, saltos de línea y espacios innecesarios"""
    if isinstance(texto, str):  # Verificar si es string
        return texto.replace("\r", "").replace("\n", "").replace("_x000D_", "").strip()
    return texto

def fill_word_template(excel_path, word_template_path, word_output_path):
  
    """
    Llena un plantilla de word con datos extraidos de un archivo excel, utilizando jinja2.

    Args:

    excel_path (str) : ruta del archivo de Excel
    word_template_path (str) : Ruta de la plantilla de word
    word_outh_path (str) : Ruta donde se guardara el documento word final.

    """
    df_reporte_detalle = pd.read_excel(excel_path, sheet_name="reporte_detalle", dtype=str, engine="openpyxl")

    df_reporte_detalle.columns = df_reporte_detalle.columns.str.strip()

    df_reporte_detalle = df_reporte_detalle.map(limpiar_texto)

    reportes_detallados = []
    for _, row in df_reporte_detalle.iterrows():
        reporte = {
            'nro_ticket': row.get('nro_incidencia', 'no data'),
            'fecha_gene': row.get('fecha_generacion', 'no data'),
            'tipo_servicio': row.get('tipo_servicio', 'no data'),
            'cid': row.get('cid', 'no data'),
            'tipo_caso': row.get('tipo_caso', ''),
            'averia': row.get('tipificacion_problema','no data'),
            'determinacion_causa': row.get('it_determinacion_de_la_causa', 'no data'),
            'medidas_tomadas': row.get('it_medidas_tomadas', 'no data'),
            'recomendaciones': row.get('it_conclusiones', 'no data')
        
        }
        reportes_detallados.append(reporte)

    doc = DocxTemplate(word_template_path)
    context  = {

        'fecha_inicio': "01/01/2025",
        'fecha_fin': "31/01/2025",
        'reportes': reportes_detallados
        }
    
    doc.render(context)
    output_file = os.path.join(word_output_path, f"reporte_mensual_completo.docx")
    doc.save(output_file)
    print(f"Reporte mensual generado exitosamente: {output_file}")
    

if __name__ == "__main__":

    base_path = "C:/Users/katana/Desktop/proyectos/bots_rpa"
    #base_path = "C:/Users/mozac/Documents/pruebas/Fast-API-Bots-RPA-python-"
    excel_path = os.path.join(base_path, "media", "pronatel", "data", "DataTicketsPronatel.xlsx")
    word_template_path = os.path.join(base_path, "media", "pronatel", "plantillas", "plantilla_word5.docx")
    word_output_path = os.path.join(base_path, "media", "pronatel", "reportes/")

    fill_word_template(excel_path, word_template_path, word_output_path)


