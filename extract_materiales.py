import pdfplumber
from openpyxl import Workbook

pdf_path = "GM_2019_V1_1-303.pdf"
start_page = 8
output_file = "materiales.xlsx"

wb = Workbook()
ws = wb.active
ws.title = "Materiales"
ws.append(["Nombre", "Formato", "Rendimiento", "Precio", "Observaciones"])

with pdfplumber.open(pdf_path) as pdf:
    for i in range(start_page - 1, len(pdf.pages)):
        page = pdf.pages[i]
        text = page.extract_text()
        if text:
            for line in text.split("\n"):
                parts = line.split()
                if len(parts) >= 4:
                    nombre = parts[0]
                    formato = parts[1]
                    rendimiento = parts[2]
                    precio = parts[-1]
                    ws.append([nombre, formato, rendimiento, precio, ""])

wb.save(output_file)
print(f"Archivo Excel generado: {output_file}")
