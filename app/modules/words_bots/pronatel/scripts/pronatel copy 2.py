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
    df_hoja5 = pd.read_excel(excel_path, sheet_name="Hoja5", dtype=str, engine="openpyxl")
    df_hoja5.columns = df_hoja5.columns.str.strip()
    # print(df_hoja4.columns.tolist())
    # env = Environment(loader=FileSystemLoader(os.path.dirname(word_template_path)))
    # template = env.get_template(os.path.basename(word_template_path))

    reportes = []
    for _, row in df_hoja5.iterrows():
    
        reporte = {
            'determinacion_causa': row.get('it_determinacion_de_la_causa', 'no data'),
            'medidas_tomadas': row.get('it_medidas_tomadas'),
            'nro_ticket': row.get('nro_incidencia'),
            'cid': row.get('cid'),
            # 'fecha_inicio': row['fecha_inicio'].strftime('%d/%m%Y'),
            # 'fecha_fin': row['fecha_fin'].strftime('%d/%m%Y')
        }
        reportes.append(reporte)

    doc = DocxTemplate(word_template_path)
    context  = {
            'reportes': reportes

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


