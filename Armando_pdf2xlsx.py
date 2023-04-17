# pip install openpyxl
# pip install pdfquery

from pdfquery import PDFQuery
import os
import glob
import openpyxl

# Inserisco tutti i pdf, presenti nella cartella dove risiede questo script, in un array
path = os.getcwd()
pdf_files = glob.glob(os.path.join(path, "*.pdf"))

# Scrivo la prima riga del foglio excel
header = ("System", "Drawing number", "Revisione")
wb = openpyxl.Workbook()
ws = wb.active
ws.append(header)

# Scorro singolarmente e leggo i pdf
for f in pdf_files:
    pdf = PDFQuery(f)
    pdf.load()

    # Lettura valori nelle rispettive celle del pdf
    system = pdf.pq(
        'LTTextLineHorizontal:in_bbox("1957.9, 82.742, 1973.786, 97.029")').text()
    drawingNumber = pdf.pq(
        'LTTextLineHorizontal:in_bbox("2050.592, 82.742, 2227.597, 97.029")').text()
    rev = pdf.pq(
        'LTTextLineHorizontal:in_bbox("2305.145, 82.742, 2321.031, 97.029")').text()

    # Divisione per colonne dei valori letti
    data = ((int(system), drawingNumber, int(rev)))
    ws.append(data)

# Savlataggio file
wb.save('output.xlsx')
