import tkinter as tk
from tkinter import filedialog
from reportlab.pdfgen import canvas
from reportlab.lib.pagesizes import letter
from indicadores import _indicadores
from tkinter import messagebox

lista_pesos = []
def getPesos(pesos):
    global lista_pesos
    lista_pesos = pesos

def criarPDF():
    logo_path = "../src/logo_tcm.jpg"
    graph_path1 = "src/report1.png"
    graph_path2 = "src/report2.png"

    page_width, page_height = letter

    root = tk.Toplevel()
    root.withdraw()
    fileName = filedialog.asksaveasfilename(
        defaultextension=".pdf",
        filetypes=[("PDF files", "*.pdf")],
        title="Salvar Relatório"
    )

    if not fileName:
        print("Salvamento cancelado pelo usuário.")
        return

    # Criar o objeto PDF
    pdf = canvas.Canvas(fileName, pagesize=letter)

    # Configurar título do documento
    pdf.setTitle("Relatório TCM-BA")

    # Página 1: Cabeçalho e Introdução
    try:
        # Inserir logotipo centralizado
        logo_width, logo_height = 120, 120
        pdf.drawImage(
            logo_path,
            (page_width - logo_width) / 2,  # Centraliza horizontalmente
            page_height - logo_height - 50,  # Posição vertical
            width=logo_width,
            height=logo_height
        )
    except Exception as e:
        print(f"Erro ao adicionar logotipo: {e}")

    # Título e descrição
    pdf.setFont("Helvetica-Bold", 20)
    pdf.drawCentredString(page_width / 2, page_height - 200, "Tribunal de Contas dos Municípios do Estado da Bahia")
    pdf.setFont("Helvetica", 14)
    pdf.drawCentredString(page_width / 2, page_height - 230, "Relatório de Indicadores")

    pdf.setFont("Helvetica", 12)
    description_lines = [
        "Este relatório apresenta os indicadores analisados afim de análise.",
        "Os indicadores são baseados em dados das matrizes construidas pelo",
        "Núcleo de Inovação e Sistemas (NICE)."
    ]
    y_position = page_height - 280
    for line in description_lines:
        pdf.drawCentredString(page_width / 2, y_position, line)
        y_position -= 15

    pdf.setFont("Helvetica-Oblique", 10)
    pdf.drawRightString(page_width - 50, 30, "Página 1 de 4")
    pdf.showPage()

    # Página 2: Lista de Indicadores
    pdf.setFont("Helvetica", 12)
    pdf.drawCentredString(page_width / 2, page_height - 100, "Indicadores Utilizados:")
    try:
        description_lines2 = _indicadores.getIndicadoresAtivos()
        y_position = page_height - 125
        for line in description_lines2:
            pdf.drawCentredString(page_width / 2, y_position, f"- {line}")
            y_position -= 15
    except Exception as e:
        print(f"Erro ao listar indicadores: {e}")

    # lista de pesos
    pdf.setFont("Helvetica", 12)
    pdf.drawCentredString(page_width / 2, page_height - 400, "Distribuição de Pesos Utilizada:")
    tipos = ['Risco','Materialidade','Relevância','Oportunidade']
    try:
        description_lines3 = lista_pesos
        y_position = page_height - 425

        for line, tipos in zip(description_lines3,tipos):
            pdf.drawCentredString(page_width / 2, y_position, f"- {tipos}: {line}%")
            y_position -= 15
    except Exception as e:
        print(f"Erro ao listar pesos: {e}")

    pdf.setFont("Helvetica-Oblique", 10)
    pdf.drawRightString(page_width - 50, 30, "Página 2 de 4")
    pdf.showPage()

    # Página 3: Gráfico 1
    try:
        graph_width, graph_height = 588, 460
        pdf.setFont("Helvetica-Bold", 14)
        pdf.drawCentredString(page_width / 2, page_height - 200, "Ranking Geral")
        pdf.drawImage(
            graph_path1,
            (page_width - graph_width) / 2,
            (page_height - graph_height) / 2 - 50,
            width=graph_width,
            height=graph_height
        )
    except Exception as e:
        print(f"Erro ao adicionar gráfico 1: {e}")

    pdf.setFont("Helvetica-Oblique", 10)
    pdf.drawRightString(page_width - 50, 30, "Página 3 de 4")
    pdf.showPage()

    # Página 4: Gráfico 2
    try:
        pdf.setFont("Helvetica-Bold", 14)
        pdf.drawCentredString(page_width / 2, page_height - 200, "Ranking Específico")
        pdf.drawImage(
            graph_path2,
            (page_width - graph_width) / 2,
            (page_height - graph_height) / 2 - 50,
            width=graph_width,
            height=graph_height
        )
    except Exception as e:
        print(f"Erro ao adicionar gráfico 2: {e}")

    pdf.setFont("Helvetica-Oblique", 10)
    pdf.drawRightString(page_width - 50, 30, "Página 4 de 4")
    pdf.showPage()

    # Salvar o PDF
    try:
        pdf.save()
        messagebox.showinfo('Sucesso',f'Relatório criado com sucesso em: \n{fileName}')
        print(f"PDF criado com sucesso: {fileName}")
    except Exception as e:
        messagebox.showerror('Erro',f'Erro ao criar relatório. {e}')

