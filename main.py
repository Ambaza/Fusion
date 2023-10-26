import tkinter as tk
from tkinter import filedialog
import os
import pandas as pd
import io
import xlsxwriter
from reportlab.lib.pagesizes import A4
from reportlab.lib.units import mm
from reportlab.pdfgen import canvas
from PyPDF2 import PdfFileMerger, PdfFileReader


# Création de la fenêtre principale
root = tk.Tk()
root.title("Application de fusion de fichiers Excel")

# Fonction pour importer des fichiers Excel
def import_excel():
    files = filedialog.askopenfilenames(title="Sélectionner des fichiers Excel", filetypes=[("Fichiers Excel", "*.xlsx;*.xls")])
    for file in files:
        print(file)

# Fonction pour fusionner les fichiers Excel
def fusion_excel():
    # Input validation
    files = filedialog.askopenfilenames(title="Sélectionner des fichiers Excel", filetypes=[("Fichiers Excel", "*.xlsx;*.xls")])
    if not files:
        print("Aucun fichier sélectionné.")
        return

    workbook = xlsxwriter.Workbook("fusion.xlsx", {'nan_inf_to_errors': True})
    for i, file in enumerate(files):
        try:
            sheets = pd.read_excel(file, sheet_name=None)
        except Exception as e:
            print(f"Erreur lors de la lecture du fichier {file}: {str(e)}")
            continue

        # Create a new worksheet for each sheet in each file
        for sheet_name, sheet_df in sheets.items():
            sheet_name = f"{os.path.basename(file)} - {sheet_name}"
            worksheet = workbook.add_worksheet(sheet_name)

            # Write the headers with the name of the original file
            header_format = workbook.add_format({"bold": True})
            worksheet.write(0, 0, os.path.basename(file), header_format)
            worksheet.write_row(1, 0, sheet_df.columns)

            # Write the data
            for row_num, row_data in sheet_df.iterrows():
                worksheet.write_row(row_num + 2, 0, row_data)

        # Progress reporting
        print(f"{i+1}/{len(files)} fichiers lus.")

    workbook.close()

    # Create a PDF from the Excel file
    pdf_bytes = io.BytesIO()
    workbook = xlsxwriter.Workbook(pdf_bytes, {"in_memory": True})
    workbook.filename = "fusion.xlsx"
    for sheet_name in workbook.sheetnames:
        worksheet = workbook.get_worksheet_by_name(sheet_name)

        # Set the page header to the file name
        file_name = sheet_name.split(" - ")[0]
        worksheet.set_header(f"&L{file_name}&R&P / &N")

    workbook.close()

    # Save the PDF
    pdf_bytes.seek(0)
    with open("fusion.pdf", "wb") as pdf_file:
        canvas_obj = canvas.Canvas(pdf_file, pagesize=A4)
        pdf_merger = PdfFileMerger()

        for page_num, page in enumerate(PdfFileReader(pdf_bytes).pages, start=1):
            page.mergePage(canvas_obj.getPage(page_num))
            pdf_merger.addPage(page)

        canvas_obj.save()
        pdf_merger.write(pdf_file)

    print("Fusion terminée.")


# Fonction pour télécharger le fichier fusionné
def download_fusion():
    root.filename = filedialog.asksaveasfilename(defaultextension=".xlsx", filetypes=[("Fichiers Excel", "*.xlsx")])
    if root.filename:
        with open("fusion.xlsx", "rb") as file:
            data = file.read()
        with open(root.filename, "wb") as file:
            file.write(data)

# Création des boutons
button_import = tk.Button(root, text="Importer des fichiers Excel", command=import_excel)
button_import.pack(pady=10)

button_fusion = tk.Button(root, text="Fusionner les fichiers Excel", command=fusion_excel)
button_fusion.pack(pady=10)

button_download = tk.Button(root, text="Télécharger le fichier fusionné", command=download_fusion)
button_download.pack(pady=10)

# Lancement de la boucle principale de l'interface graphique
root.mainloop()
