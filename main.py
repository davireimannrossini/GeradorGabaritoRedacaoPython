import pandas as pd
from PyPDF2 import PdfReader, PdfWriter
from reportlab.lib.pagesizes import letter
from reportlab.pdfgen import canvas
import io
import os
from datetime import datetime

def preencher_campos_pdf(modelo_pdf, dados, novo_pdf):
    reader = PdfReader(modelo_pdf)
    writer = PdfWriter()

    for page_num in range(len(reader.pages)):
        page = reader.pages[page_num]

        packet = io.BytesIO()
        pdf_canvas = canvas.Canvas(packet, pagesize=letter)

        # Preencher os campos com os valores dos dados
        pdf_canvas.setFont("Helvetica", 11)

        # LINHA 1
        pdf_canvas.drawString(108, 733, str(dados["Nome"]))
        pdf_canvas.drawString(392, 733, str(dados["RA"]))
        data_formatada = dados["Data"].strftime("%d/%m/%Y")
        pdf_canvas.drawString(488, 733, str(data_formatada))

        # LINHA 2
        pdf_canvas.drawString(108, 717, str(dados["Turma"]))
        pdf_canvas.drawString(506, 717, str(dados["Unidade"]))

        # LINHA 3
        pdf_canvas.drawString(108, 703, str(dados["Tema"]))
        pdf_canvas.drawString(492, 703, str(dados["Ciclo"]))





        pdf_canvas.save()

        packet.seek(0)
        new_pdf = PdfReader(packet)

        new_page = new_pdf.pages[0]
        page.merge_page(new_page)

        writer.add_page(page)

        gabaritos_dir = "gabaritos"
        os.makedirs(gabaritos_dir, exist_ok=True)
        novo_pdf_path = os.path.join(gabaritos_dir, novo_pdf)
        
    with open(novo_pdf_path, "wb") as output_file:
        writer.write(output_file)

# Ler o arquivo Excel
dados_excel = pd.read_excel('arquivo.xlsx')

# Iterar sobre as linhas do arquivo Excel e gerar um PDF para cada linha
for _, linha in dados_excel.iterrows():
    dados = {
        "Nome": linha["Nome"],
        "RA": linha["RA"],
        "Turma": linha["Turma"],
        "Unidade": linha["Unidade"],
        "Tema": linha["Tema"],
        "Ciclo": linha["Ciclo"],
        "Data": linha["Data"]
    }

    # Gerar o nome do arquivo PDF
    nome_arquivo_pdf = f"{dados['Ciclo']} - {dados['Nome']} - {dados['RA']} - {dados['Turma']}.pdf"

    # Gerar o novo arquivo PDF preenchendo os campos
    preencher_campos_pdf("modelo.pdf", dados, nome_arquivo_pdf)
