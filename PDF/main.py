import os
import PyPDF2
import re
import pandas as pd
from pandas import ExcelWriter
import openpyxl
from openpyxl import load_workbook

def extraccion_info(pdf_file_path):
    #abrir el PDF file
    with open(pdf_file_path,'rb') as file:

        pdf_render = PyPDF2.PdfReader(file)

        text =''

        for page_num in range(len(pdf_render.pages)):
            page = pdf_render.pages[page_num]
            text += page.extract_text()

        #Extresiones Regulares
        numero_f = r"Nro:\s*(\d+-\d+)"
        items_e= r"Calibración \s*(.*)"
        Importe_f = r"Total\s*(.*)"
        Proveedor_f = r"FACTURA\s*(.*)"
        
        # extrae info con la extresion regular
        numero_f_match = re.search(numero_f, text)
        items_match = re.search(items_e,text)
        import_match = re.search(Importe_f, text)
        proveedor_match = re.search(Proveedor_f, text)
        #guarda
        numero = numero_f_match.group(1) if numero_f_match else None
        item= items_match.group() if items_match else None
        Importe= import_match.group(1) if import_match else None
        Proveedor= proveedor_match.group(1) if proveedor_match else None

        

        

        return numero, item, Importe, Proveedor
    
def get_files_in_folder(folder_path):
    files= []
    for root, dirs, filenames in os.walk(folder_path):
        for filename in filenames:
            files.append(os.path.join(root,filename))
    return files

if __name__ == "__main__":
    folder_path = "Facturas"
    files= get_files_in_folder(folder_path)

    for file in files:

        print("File: ", file)
        numero, item, Importe, Proveedor = extraccion_info(file)

        print(f"numero: {numero}")
        print(f"item: {item}")
        print(f"total: {Importe}")
        print(f"proveedor: {Proveedor}")

        data={
            "Numero":[numero],
            "items":[item],
            "Importe":[Importe],
            "Proveedor":[Proveedor]
        }
        archivo_excel = 'Libro2.xlsx'
        df_existente = pd.read_excel(archivo_excel)
        nuevo_df = pd.DataFrame(data, index=[0])
        df_actualizado = pd.concat([df_existente, nuevo_df], ignore_index=True)
        df_actualizado.to_excel(archivo_excel, index=False)

        #mover a la carpeta de vista
        realizado_file = "Faturas_vista"
        os.makedirs(realizado_file, exist_ok=True)
        new_files_path= os.path.join(realizado_file, os.path.basename(file))
        os.rename(file,new_files_path)