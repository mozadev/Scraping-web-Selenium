from docxtpl import DocxTemplate
from datetime import datetime
import os

base_path = "C:/Users/katana/Desktop/proyectos/bots_rpa"
word_template_path = os.path.join(base_path, "media", "pronatel", "plantillas", "template.docx")
word_output_path = os.path.join(base_path, "media", "pronatel", "reportes/")

doc = DocxTemplate(word_template_path)

my_name ="jose torrez"
my_phone = "(123) 456-156"
my_email = "jose@gmail.com"
my_address = "123 Main Street, NY"
today_date= datetime.today().strftime("%d %b, %Y")

context = {
     'my_name': my_name,
     'my_phone': my_phone,
     'my_email': my_email,
     'my_address': my_address,
     'today_date': today_date
     }
doc.render(context)
doc.save(f"{word_output_path}generate_doc.docx")
