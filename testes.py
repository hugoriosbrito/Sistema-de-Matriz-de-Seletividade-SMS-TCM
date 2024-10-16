import folium.features
import openpyxl as xl
import customtkinter as ctk
from tkinter import messagebox
import pandas as pd
import matplotlib.pyplot as plt
from matplotlib.backends.backend_tkagg import FigureCanvasTkAgg
import os
import win32com.client
import folium
import webview

# Função para carregar e processar os dados do Excel
def carregar_dados():
    global dfPlot, irce_por_mun_dce1, irce_por_mun_dce2

    # Carregar dados do Excel
    df = pd.read_excel("dados\\Matriz Modelo - VERSÃO SISTEMA.xlsx", sheet_name='MATRIZ CONTRATOS')

    # Filtrando os dados relevantes
    dfIDs = df.iloc[6:, 0]
    dfMunicipio = df.iloc[6:, 1]
    dfNota = df.iloc[6:, 34]
    dfIRCE = df.iloc[6:, 2]
    dfDCE = df.iloc[6:, 3]

    # Criando um novo DataFrame para facilitar o uso
    novo_df = {
        'id': dfIDs.values,
        'municipio': dfMunicipio.values,
        'irce': dfIRCE.values,
        'dce': dfDCE.values,
        'nota': dfNota.values
    }
    dfPlot = pd.DataFrame(novo_df)
    dfPlot = dfPlot.sort_values(by='nota', ascending=False)

    # Filtrando por DCE
    dce_1, dce_2 = '1ª DCE', '2ª DCE'

    # IRCEs associadas à 1ª e 2ª DCE
    df_1_DCE = dfPlot[dfPlot['dce'] == dce_1]
    df_2_DCE = dfPlot[dfPlot['dce'] == dce_2]

    # Listas de IRCEs únicas para cada DCE
    irce_list_dce1 = df_1_DCE['irce'].drop_duplicates().tolist()
    irce_list_dce2 = df_2_DCE['irce'].drop_duplicates().tolist()

    # Criando dicionários com IRCEs como chaves e municípios como valores
    irce_por_mun_dce1 = {irce: df_1_DCE[df_1_DCE['irce'].str.strip() == irce.strip()]['municipio'].tolist() for irce in irce_list_dce1}
    irce_por_mun_dce2 = {irce: df_2_DCE[df_2_DCE['irce'].str.strip() == irce.strip()]['municipio'].tolist() for irce in irce_list_dce2}

# Função para plotar o gráfico baseado na filtragem de DCE e IRCE
def plotar_ranking_filtrado():
    selected_dce = dce_var.get()
    selected_irce = irce_var.get()

    if selected_dce == '1ª DCE':
        df_filtrado = dfPlot[(dfPlot['dce'] == '1ª DCE') & (dfPlot['irce'] == selected_irce)]
    elif selected_dce == '2ª DCE':
        df_filtrado = dfPlot[(dfPlot['dce'] == '2ª DCE') & (dfPlot['irce'] == selected_irce)]
    else:
        messagebox.showerror("Erro", "Selecione uma DCE e uma IRCE válidas")
        return

    if df_filtrado.empty:
        messagebox.showinfo("Informação", "Nenhum município encontrado para a IRCE selecionada.")
        return

    df_filtrado = df_filtrado.sort_values(by='nota', ascending=False)

    fig = plt.figure(figsize=(10, 10))

    # Plotagem do ranking
    plt.barh(df_filtrado['municipio'], df_filtrado['nota'], color='orange', height=0.5)
    plt.gca().invert_yaxis()
    plt.xlabel('Nota', fontsize=12, color='white')
    plt.ylabel('Município', fontsize=12, color='white')
    plt.title(f'Municípios da {selected_irce} ({selected_dce})', fontsize=15, color='white')
    plt.gca().set_facecolor("#3C91E6")  # Fundo do gráfico
    fig.patch.set_facecolor("#3C91E6")  # Fundo da figura
    plt.gca().tick_params(axis='y', colors='white')
    plt.gca().tick_params(axis='x', colors='white')
    plt.tight_layout()

    # Limpando o frame antes de desenhar o novo gráfico
    for widget in frame_ranking_geral.winfo_children():
        widget.destroy()

    # Adicionando o gráfico à interface
    canvas = FigureCanvasTkAgg(fig, master=frame_ranking_geral)
    canvas.draw()
    canvas.get_tk_widget().pack(side="top", fill='both', expand=True)

# Função para atualizar as IRCEs com base na DCE selecionada
def atualizar_irces():
    selected_dce = dce_var.get()

    if selected_dce == '1ª DCE':
        irce_menu.configure(values=list(irce_por_mun_dce1.keys()))
    elif selected_dce == '2ª DCE':
        irce_menu.configure(values=list(irce_por_mun_dce2.keys()))
    irce_menu.set('')  # Limpar a seleção da IRCE

# Interface gráfica
window = ctk.CTk()
window.geometry("800x600")

# Frame principal
frame = ctk.CTkFrame(master=window, border_width=0, corner_radius=20, bg_color="#2F83D7", fg_color="#2F83D7", height=250)
frame.pack(fill="both", expand=True)

# Frame para o gráfico
frame_ranking_geral = ctk.CTkFrame(master=window)
frame_ranking_geral.pack(side='left', fill='both', expand=True)

# Variável para armazenar a DCE selecionada
dce_var = ctk.StringVar()

# Botão de lista para selecionar 1ª ou 2ª DCE
dce_menu = ctk.CTkOptionMenu(master=frame, variable=dce_var, values=['1ª DCE', '2ª DCE'], command=lambda _: atualizar_irces())
dce_menu.pack(pady=10)

# Botão de lista para mostrar IRCEs de acordo com a DCE
irce_var = ctk.StringVar()
irce_menu = ctk.CTkOptionMenu(master=frame, variable=irce_var, values=[])
irce_menu.pack(pady=10)

# Botão para plotar o gráfico baseado na filtragem
plotar_button = ctk.CTkButton(master=frame, text="Plotar Ranking", command=plotar_ranking_filtrado)
plotar_button.pack(pady=10)

# Carregar os dados do Excel
carregar_dados()

# Iniciar a interface
window.mainloop()
