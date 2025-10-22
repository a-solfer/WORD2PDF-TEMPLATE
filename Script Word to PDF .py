#!/usr/bin/env python
# coding: utf-8

# #para la encriptacion de doc PDF
# !pip install PyPDF2
# from PyPDF2 import PdfReader, PdfWriter
# import os

# # we retrieve the file name of the pdf
# pdf_list = [file for file in os.listdir(r'C:\Users\Owner\Documents\personal\professional\JOBS\CREDICORP\CONTRATOS XLSX TO PDF\contratos\Trial Docs\Save Doc Here\PDF') if '.pdf' in file]
# print(pdf_list)
# 
# file_name = pdf_list
# print(file_name)
# 
# # we want the password to 
# 
# #now we encryopt the PDF
# 
# password= input('Set your password for the PDF\n')
# 

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

# Initialize Word application
word_app = win32com.client.Dispatch("Word.Application")

# Load the EXCEL data into a DataFrame
#REEMPLAZAR EL FILE PATH POR EL FILE PATH DEL EXCEL A USAR
data_frame = pd.read_excel(r'C:\Users\Owner\Documents\personal\professional\code\Python\Word to PDF\Excel_Document.xlsx')


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

# Make sure the PDF and PROTECTED directories exist (create them if necessary)
os.makedirs(pdf_dir, exist_ok=True)
os.makedirs(encripted_dir, exist_ok=True)



# Iterate over the DataFrame rows
for r_index, row in data_frame.iterrows():
    print(f"Processing row {r_index}")  # Debugging to check iteration
    
    document_name = row['DOC_NAME'] #'DOC_NAME' is the column that will determine the name of the word template
    password = str(row['PASSWORD'])  # The password will take the value under the PASSWORD column
    
    print(f"A new document has been created: {document_name}")  # Debugging to check cta inversion
    print(f"Password for {document_name}: {password}")  # Debugging password
    



#PARTE 2
    
    # Clean client name to ensure it's a valid filename
        #when using clean_filename make sure to convert the value to text
    document_name = clean_filename(str(document_name))

    # Load the template document
    #REEMPLAZAR SEGUN UBICACION DE PLANTILLA DEL TEMPLATE 
    template = DocxTemplate(r"C:\Users\Owner\Documents\personal\professional\code\Python\Word to PDF\word_document.docx")

    # Convert the DataFrame to a list of dictionaries (for the template rendering)
    context = data_frame.to_dict(orient='records')

    # Render the template with the current context (data for the current row)
    template.render(context[r_index])  # Use r_index to get the corresponding dictionary

    # Save the rendered document as .docx
    docx_path = os.path.join(word_dir, document_name + ".docx")
    template.save(docx_path)
    print(f"Document saved as .docx for {document_name}")

    # Convert the .docx file to .pdf using Word
    pdf_path = os.path.join(pdf_dir, document_name + ".pdf")
    print(f"Saving {document_name}.pdf...")

    doc = word_app.Documents.Open(docx_path)
    print('Exporting to PDF...')
    
    # Save as PDF
    doc.SaveAs(pdf_path, FileFormat=17)  # FileFormat=17 corresponds to PDF in Word
    doc.Close()  # Close the document
    
#PARTE3

    # Encrypt the PDF file using the password from the NUM_DOC column
    if password:
        print(f"Encrypting {document_name}.pdf with password...")
        
        try:
            # Open the generated PDF file
            with open(pdf_path, "rb") as pdf_file:
                pdf_reader = PdfReader(pdf_file)
                pdf_writer = PdfWriter()

                # Add all pages to the writer
                for page_num in range(len(pdf_reader.pages)):
                    pdf_writer.add_page(pdf_reader.pages[page_num])

                # Encrypt the PDF with the password
                pdf_writer.encrypt(password)

                # Save the encrypted PDF to the PROTECTED folder
                encrypted_pdf_path = os.path.join(encripted_dir, document_name + "_encrypted.pdf")
                with open(encrypted_pdf_path, "wb") as encrypted_pdf:
                    pdf_writer.write(encrypted_pdf)

                print(f"Encrypted PDF saved as: {encrypted_pdf_path}")

            # Optionally, remove the unencrypted PDF (if you want to)
            #os.remove(pdf_path)  # Uncomment if you want to delete the unencrypted PDF

        except Exception as e:
            print(f"Error encrypting {document_name}: {e}")
    else:
        print(f"No password found for {document_name}. PDF not encrypted.")

# Close the Word application (once after all rows are processed)
word_app.Quit()

print("Conversion and encryption completed for all documents.")



