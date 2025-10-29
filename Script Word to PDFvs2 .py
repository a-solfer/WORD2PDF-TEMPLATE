
# THIS VERSION OF WORD2PDF VS 2 Includes multiple validation checks such as
# - Validating if the necessary folders exist (creating them if not)
# - Validating if the template and excel files exist
# - Validating if the necessary columns exist in the excel file
# Also includes better debugging messages to track the progress of the script


# #para la encriptacion de doc PDF
# !pip install PyPDF2
# from PyPDF2 import PdfReader, PdfWriter

#PARTE 1

import os
import pandas as pd
from docxtpl import DocxTemplate
import re
#from win32com import client  DELTE THIS LINE
import win32com.client
from PyPDF2 import PdfReader, PdfWriter

#we TEST win32com has been succesfully installed
#print("win32com.client imported succesfully")

# Function to clean filename (remove invalid characters)
def clean_filename(name):
    return re.sub(r'[\\/*?:"<>|]', "_", name)

#MAIN Directories and FILES
base_dir = r"C:\Users\Owner\Documents\personal\professional\code\Python\Word to PDF"
excel_path = os.path.join(base_dir, "Excel_Document.xlsx")
template_path = os.path.join(base_dir, "word_document.docx")

# FOLDER STRUCTURE (WHERE THE DOCUMENTS BE SAVED)
# Directories for saving files
#general directory
output_dir = r'C:\Users\Owner\Documents\personal\professional\code\Python\Word to PDF\automatically saved docs folder'
#Directory for Word docs
word_dir = os.path.join(output_dir, 'save docx here folder')
#Directory for not password protected pdf's
pdf_dir = os.path.join(output_dir, 'save pdf here folder')
#Directory for encripted pdf's
encripted_dir = os.path.join(output_dir, 'save pdf encrypted here folder')

#VALIDATE IT EXISTS
# 📁 Make sure the PDF and PROTECTED directories exist (create them if necessary)
for folder in [output_dir, word_dir, pdf_dir, encripted_dir]:
    os.makedirs(folder, exist_ok=True) #* we simplify the cheking of the necessary folders

# ✅ Validar existencia de plantilla
if not os.path.isfile(template_path):
    raise FileNotFoundError(f"❌ Plantilla Word no encontrada en: {template_path}")

# 📊 Cargar Excel
if not os.path.isfile(excel_path):
    raise FileNotFoundError(f"❌ Archivo Excel no encontrado en: {excel_path}")

data_frame = pd.read_excel(excel_path)


# ✅ Validar columnas necesarias
required_columns = ['DOC_NAME', 'PASSWORD', 'RECIPIENTE', 'ADDRESS', 'PHONE', 'CREDIT_CARD_BALANCE', 'CREDIT_CARD_TYPE', 'BUSINESS']
missing = [col for col in required_columns if col not in data_frame.columns]
if missing:
    raise ValueError(f"❌ Faltan columnas requeridas en el Excel: {missing}")


# Initialize Word application
word_app = win32com.client.Dispatch("Word.Application")
word_app.Visible = False  # Run Word in the background


# Iterate over the DataFrame rows
for r_index, row in data_frame.iterrows():
    print(f"Processing row {r_index}")  # Debugging to check iteration
    
    raw_name = row['DOC_NAME'] #'DOC_NAME' is the column that will determine the name of the word template
    password = str(row['PASSWORD'])  # The password will take the value under the PASSWORD column
    #timestamp = datetime.now().strftime("%Y%m%d_%H%M")
    #document_name = clean_filename(f"{raw_name}_{timestamp}") #I dont want the timestamp in the name of the document
    document_name = clean_filename(raw_name)
    
    #Directories to save file by type
    docx_path = os.path.join(word_dir,f"{document_name}.docx")
    #template_path.save(docx_path) #TODO maybe delete
    pdf_path = os.path.join(pdf_dir,f"{document_name}.pdf")
    encrypted_pdf_path = os.path.join(encripted_dir,f"{document_name}_encrypted.pdf")




    # 🛑 Si ya existe el PDF, saltar todo
    if os.path.isfile(pdf_path):
        print(f"⏭️ PDF ya existe: {pdf_path}. Saltando fila {r_index}.")
        continue

    # 🛑 Si ya existe el Word pero no el PDF → convertir a PDF
    if os.path.isfile(docx_path) and not os.path.isfile(pdf_path):
        print(f"📄 Word ya existe. Convirtiendo a PDF: {docx_path}")
        try:
            doc = word_app.Documents.Open(docx_path)
            doc.SaveAs(pdf_path, FileFormat=17)
            doc.Close()
            print(f"✅ PDF generado desde Word existente: {pdf_path}")
        except Exception as e:
            print(f"❌ Error al convertir Word existente a PDF: {e}")
        continue

    # 🛑 Si ya existe el Word y el PDF → saltar
    if os.path.isfile(docx_path):
        print(f"⏭️ Word ya existe: {docx_path}. Saltando fila {index}.")
        continue


    # 📝Convert the DataFrame to a list of dictionaries (for the template rendering)
        #context = data_frame.to_dict(orient='records') #!old code
    
    context = {
        "DOC_NAME": row['DOC_NAME'],
        "PASSWORD": password,
        "RECIPIENTE": row['RECIPIENTE'],
        "ADDRESS": row['ADDRESS'],
        "PHONE": row['PHONE'],
        "CREDIT_CARD_BALANCE": row['CREDIT_CARD_BALANCE'],
        "CREDIT_CARD_TYPE": row['CREDIT_CARD_TYPE'],
        "BUSINESS": row['BUSINESS']
    }

    try:
        template = DocxTemplate(template_path)
        template.render(context)
        template.save(docx_path)
        print(f"✅ Documento Word creado: {docx_path}")
    except Exception as e:
        print(f"❌ Error al generar Word para fila {index}: {e}")
        continue

        # 📄 Convertir a PDF
    try:
        doc = word_app.Documents.Open(docx_path)
        doc.SaveAs(pdf_path, FileFormat=17)
        doc.Close()
        print(f"✅ PDF generado: {pdf_path}")
    except Exception as e:
        print(f"❌ Error al convertir a PDF: {e}")
        continue

    # 🔐 Encriptar PDF #todo it does not encript if pdf is available
    try:
        with open(pdf_path, "rb") as f:
            reader = PdfReader(f)
            writer = PdfWriter()
            for page in reader.pages:
                writer.add_page(page)
            writer.encrypt(password)
            with open(encrypted_pdf_path, "wb") as ef:
                writer.write(ef)
        print(f"🔐 PDF encriptado guardado: {encrypted_pdf_path}")
    except Exception as e:
        print(f"❌ Error al encriptar PDF: {e}")

# 🧹 Cerrar Word
try:
    word_app.Quit()
except Exception as e:
    print(f"⚠️ Error al cerrar Word: {e}")

print("\n✅ Conversion and encryption completed for all documents.")



