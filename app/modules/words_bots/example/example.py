from docxtpl import DocxTemplate
from datetime import datetime
import os
import pandas as pd

base_path = "C:/Users/katana/Desktop/proyectos/bots_rpa"

word_template_path = os.path.join(base_path, "media", "words_docxtpl", "plantillas", "template-manager-info.docx")
excel_path = os.path.join(base_path, "media", "words_docxtpl", "data", "sp_fake_data.csv")
word_output_path = os.path.join(base_path, "media", "words_docxtpl", "reportes/")

doc = DocxTemplate(word_template_path)

my_name ="jose torrez"
my_phone = "(123) 456-156"
my_email = "jose@gmail.com"
my_address = "123 Main Street, NY"
today_date= datetime.today().strftime("%d %b, %Y")

my_context = {
     'my_name': my_name,
     'my_phone': my_phone,
     'my_email': my_email,
     'my_address': my_address,
     'today_date': today_date
     }

df = pd.read_csv(excel_path)

for index, row in df.iterrows():
   
#    print(f"index: {index}")
#    print(f"row type:{type(row)}")
#    print(f"Row index (columns name): {row.index}")
#    print(f"Row values: {row.values}")
#    print("---")


   context = {
        'hiring_manager_name':row['name'],
        'address':row['address'],
        'phone_number':row['phone_number'],
        'email':row['email'],
        'job_position':row['job'],
        'company_name':row['company'],
   }
   context.update(my_context)
   doc.render(context)
   doc.save(f"{word_output_path}generate_doc_{index}.docx")
