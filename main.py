import folium.features
import openpyxl as xl
import customtkinter as ctk
from tkinter import messagebox
import pandas as pd
import matplotlib.pyplot as plt
from matplotlib.ticker import MaxNLocator
from matplotlib import rcParams
import numpy as np
from matplotlib.backends.backend_tkagg import FigureCanvasTkAgg
import os, sys, win32com.client
import folium
import webview
import asyncio
from indicadores.indicadores_block import indicadores_block
from pdf_report.report_pdf import criarPDF, getPesos

icon ="src\\icon.ico"
map_file = "dados\\mapa_cloropleto_bahia.html"
file = "dados\\Matriz Modelo- VERSÃO SISTEMA - ATUAL.xlsx"
wb = xl.load_workbook(file)
sheet = wb['SÍNTESE']

window = ctk.CTk()
fonte_geral_texto = ctk.CTkFont(family='Arial', size=15, weight='bold')


def resource_path(relative_path):
    """ Retorna o caminho absoluto do recurso, mesmo quando empacotado como .exe """
    try:
        # PyInstaller cria uma pasta temporária e armazena o caminho nela
        base_path = sys._MEIPASS
    except Exception:
        base_path = os.path.abspath(".")
    return os.path.join(base_path, relative_path)


#destroi a instancia criada pelo openpyxl
def on_closing():
    try:
        wb.close()
    finally:
        window.destroy()

class MainWindow:
    def window_config(window):
        window.geometry("800x600")
        ctk.set_appearance_mode("light")
        ctk.set_default_color_theme("dark-blue")
        window.configure(fg_color="#3C91E6")
        window.title("Sistema de Matriz de Seletividade")
        window.iconbitmap(icon)


fonte_titulo = ctk.CTkFont(family='Arial', size=40, weight='bold')
titulo = ctk.CTkLabel(window, text="Sistema de Matriz de Seletividade", font=fonte_titulo, anchor="center", corner_radius=20, text_color="white")
titulo.pack(pady=20, padx=20, anchor="center")  # Pequeno espaçamento nas laterais do título

frame_botoes = ctk.CTkFrame(master=window, border_width=0, corner_radius=20, bg_color="#2F83D7", fg_color="transparent", height=50)
frame_botoes.pack(fill='both', padx=20)

frame_dist_peso = ctk.CTkFrame(master=window, border_width=0, corner_radius=20, bg_color="#2F83D7", fg_color="#2F83D7", height=20)
frame_dist_peso.pack(fill = 'both', padx = 20)

frame = ctk.CTkScrollableFrame(
    master=window,
    border_width=0,
    corner_radius=20,
    bg_color="#2F83D7",
    fg_color="#2F83D7",
    height=250,
    scrollbar_button_color="#D9D9D9",
    scrollbar_button_hover_color="#FFFFFF",
)
frame.pack(fill='both', padx=20,pady=10,expand=1)

frame.update_idletasks()
frame.update()

frame.grid_columnconfigure(0,weight=1)
frame.grid_columnconfigure(1,weight=1)
frame.grid_columnconfigure(2,weight=1)
frame.grid_columnconfigure(3,weight=1)

#Alterações na distribuição de peso

distribuicao_fonte = ctk.CTkFont(family='Arial', size=15, weight='bold')
distribuicao_titulo = ctk.CTkLabel(master=frame_dist_peso, text= "Distribuição de peso \n por tipo ", font=distribuicao_fonte, text_color='white', corner_radius=20, anchor="center")
distribuicao_titulo.grid(padx=(20,20),pady=10, row = 0, column=0)
valores = list(map(str,list(range(0,105,5))))

tipos_fonte = ctk.CTkFont(family='Arial', size= 15, weight= "bold")
# todas as vezes que um valor é escolhido em um  combo, ele é adicionado à uma lista e somado para verificação, mensagem de verificação
# apenas para quando a lista tiver 4 elementos
# se a soma for igual a 100, a mensagem é de sucesso, e é salvo, se não, a mensagem é de erro apenas, sem salvamento

def validar_distribuicao():
    list_pesos = salvar_indicadores(value=0)
    if sum(list_pesos) == 100:return True

def salvar_indicadores(value):
      tipo_risco = tipo_risco_var.get()
      tipo_relevancia = tipo_relevancia_var.get()
      tipo_materialidade = tipo_materialidade_var.get()
      tipo_oportunidade = tipo_oportunidade_var.get()

      indicadores_pesos = [tipo_risco, tipo_relevancia, tipo_materialidade, tipo_oportunidade]

      distribuicao_aviso_frame = ctk.CTkFrame(frame_dist_peso, corner_radius=10)
      distribuicao_aviso_fonte = ctk.CTkFont(family='Arial', size=12, weight='bold')
      distribuicao_aviso_titulo = ctk.CTkLabel(master=distribuicao_aviso_frame, text= "", font=distribuicao_aviso_fonte, text_color='white', corner_radius=0, anchor="center")
      distribuicao_aviso_titulo.grid(padx=(20,20),pady=10, row = 0, column=0)
      distribuicao_aviso_frame.grid(padx=(10,10),pady=10, row = 1, column=0)

      if sum(indicadores_pesos) != 100:
        distribuicao_aviso_titulo.configure(text_color = "#5C492C", text = "Soma de porcentagens\n diferente de 100%!")
        distribuicao_aviso_frame.configure(fg_color = "#FDC57F")
      else:
        distribuicao_aviso_titulo.configure(text_color = "#1B5E2E", text = "Soma de porcentagens\n igual a 100%")
        distribuicao_aviso_frame.configure(fg_color = "#3CE270")

      sheet['F4'] = tipo_risco/100
      sheet['F5'] = tipo_materialidade/100
      sheet['F6'] = tipo_relevancia/100
      sheet['F7'] = tipo_oportunidade/100

      getPesos(indicadores_pesos)
      return indicadores_pesos


tipo_risco_titulo = ctk.CTkLabel(master=frame_dist_peso, text= "Risco (%)", text_color='white', corner_radius=10, font= tipos_fonte)
tipo_risco_var = ctk.IntVar(value=0)
tipo_risco_box = ctk.CTkOptionMenu(master= frame_dist_peso, values=valores, width=100,height=20, fg_color="white", text_color="black", command=salvar_indicadores, variable=tipo_risco_var)
tipo_risco_titulo.grid(padx=10,pady=5, row= 0, column=1)
tipo_risco_box.grid(padx=10,pady=5, row=1, column=1)

tipo_relevancia_titulo = ctk.CTkLabel(master=frame_dist_peso, text= "Relevância (%)", text_color='white', corner_radius=10, font= tipos_fonte)
tipo_relevancia_var = ctk.IntVar(value=0)
tipo_relevancia_box = ctk.CTkOptionMenu(master= frame_dist_peso, values=valores, width=100,height=20, fg_color="white", text_color="black", command=salvar_indicadores, variable=tipo_relevancia_var)
tipo_relevancia_titulo.grid(padx=10,pady=5, row= 0, column=2)
tipo_relevancia_box.grid(padx=10,pady=5, row=1, column=2)

tipo_materialidade_titulo = ctk.CTkLabel(master=frame_dist_peso, text= "Materialidade (%)", text_color='white', corner_radius=10, font= tipos_fonte)
tipo_materialidade_var = ctk.IntVar(value=0)
tipo_materialidade_box = ctk.CTkOptionMenu(master= frame_dist_peso, values=valores, width=100,height=20, fg_color="white", text_color="black", command=salvar_indicadores,variable=tipo_materialidade_var)
tipo_materialidade_titulo.grid(padx=10,pady=5, row= 0, column=3)
tipo_materialidade_box.grid(padx=10,pady=5, row=1, column=3)

tipo_oportunidade_titulo = ctk.CTkLabel(master=frame_dist_peso, text= "Oportunidade (%)", text_color='white', corner_radius=10, font= tipos_fonte)
tipo_oportunidade_var = ctk.IntVar(value=0)
tipo_oportunidade_box = ctk.CTkOptionMenu(master= frame_dist_peso, values=valores, width=100,height=20, fg_color="white", text_color="black", command=salvar_indicadores,variable=tipo_oportunidade_var)
tipo_oportunidade_titulo.grid(padx=10,pady=5, row= 0, column=4)
tipo_oportunidade_box.grid(padx=10,pady=5, row=1, column=4)

frame_dist_peso.columnconfigure(1,weight=1)
frame_dist_peso.columnconfigure(2,weight=1)
frame_dist_peso.columnconfigure(3,weight=1)
frame_dist_peso.columnconfigure(4,weight=1)

# -------------------------------------------- BLOCOS DE INDICADORES --------------------------------------------------------------
def bloco_indicadores():
    indicadores_block(frame=frame,sheet=sheet)

#----------------------------------------------------------------------------------------------------------------------------------

xlapp = None

def get_excel_app():
    global xlapp
    if not xlapp:
        xlapp = win32com.client.DispatchEx("Excel.Application")
    return xlapp

def refresh_file(file):
    xlapp = get_excel_app()
    path = os.path.abspath(file)
    wb = xlapp.Workbooks.Open(path)

    try:
        wb.RefreshAll()
        xlapp.CalculateUntilAsyncQueriesDone()
        wb.Save()
    finally:
        wb.Close()
        xlapp.Quit()

def hide_all():
    frame.pack_forget()
    frame_dist_peso.pack_forget()
    frame_ranking_geral.pack_forget()
def show_all():
    frame_dist_peso.pack(fill='both', padx=20)
    frame.pack(fill='both', padx=20, pady=10, expand=1)

# Frames principais
frame_ranking_geral = ctk.CTkFrame(master=window,fg_color='#3C91E6')
frame_plotagem_ranking_geral = ctk.CTkFrame(master=frame_ranking_geral, fg_color='#3C91E6')

frame_ranking_filtrado = ctk.CTkFrame(master=frame_ranking_geral, fg_color="#3C91E6", width=1000)
frame_ranking_filtrado.grid(padx=10, pady=5,sticky="nsew",column=1)

frame_plotagem_ranking_filtrado = ctk.CTkFrame(master=frame_ranking_filtrado, fg_color='#3C91E6',width=1000)
frame_plotagem_ranking_filtrado.grid(padx=10, pady=5, row=0, column=0)

frame_botoes_ranking_filtrado = ctk.CTkFrame(master=frame_ranking_filtrado, fg_color='#3C91E6',width=1000)
frame_botoes_ranking_filtrado.grid(padx=10, row=1, column=0, sticky='w')


def plotar_ranking_geral(dfPlot):
    global canvas, frame_ranking_geral
    corGeral = "white"
    dfTop50 = dfPlot.head(50)

    frame_ranking_geral.pack(expand=True, fill='both')
    frame_plotagem_ranking_geral.grid(padx=5, row=0, column=0)

    # Configurações gerais do estilo
    rcParams.update({
        "axes.edgecolor": f"{corGeral}",
        "axes.grid": False,
        "axes.titleweight": "bold",
        "font.family": "Arial",
        "font.size": 10,
    })

    fig, ax = plt.subplots(figsize=(10, 8))

    # Gradiente de vermelho (#D03645) para amarelo (#d0cd36)
    red = np.array([208, 54, 69]) / 255
    green = np.array([208, 205, 54]) / 255

    gradients = [red + (green - red) * (i / len(dfTop50)) for i in range(len(dfTop50))]

    # Aplicação do gradiente como cores das barras
    bars = ax.barh(
        dfTop50['municipio'],
        dfTop50['nota'],
        color=gradients,
        height=0.6,  # Ajustando a altura das barras
        linewidth=0.5
    )

    ax.invert_yaxis()  # Inverter a ordem dos municípios
    ax.xaxis.set_major_locator(MaxNLocator(integer=True))

    # Ajustando a largura das margens para os nomes
    ax.set_xlim([dfTop50['nota'].min() - 1, dfTop50['nota'].max() + 1])
    ax.set_xlabel("Nota", fontsize=12, color=f"{corGeral}", labelpad=15)
    ax.set_ylabel("Município", fontsize=12, color=f"{corGeral}", labelpad=15)
    ax.set_title("Top 50 Municípios por Nota", fontsize=16, color=f"{corGeral}", pad=20)
    print(dfTop50)

    # Fundo
    ax.set_facecolor("#3C91E6")
    fig.patch.set_facecolor("#3C91E6")

    # Cor e tamanho dos ticks
    ax.tick_params(axis='y', colors=f"{corGeral}", which="major", labelsize=8.5)
    ax.tick_params(axis='x', colors=f"{corGeral}", which="major", labelsize=8.5)

    # Adiciona valores ao final das barras
    for bar in bars:
        ax.text(
            bar.get_width() + 0.3,  # Ajustando posição X para os valores
            bar.get_y() + bar.get_height() / 2,  # Posição central no eixo Y
            f"{bar.get_width():.1f}",  # Valor formatado
            va="center", ha="left", fontsize=9, color=corGeral
        )

    # Rotacionando os nomes dos municípios para melhorar a legibilidade
    plt.yticks(rotation=0)

    plt.tight_layout()  # Ajustando o layout para evitar corte

    # Limpando o frame antes de desenhar o novo gráfico
    for widget in frame_plotagem_ranking_geral.winfo_children():
        widget.destroy()

    # Adiciona o gráfico ao frame
    canvas = FigureCanvasTkAgg(fig, master=frame_plotagem_ranking_geral)
    canvas.draw()
    canvas.get_tk_widget().grid(padx=10, sticky='nsew', pady=10, column=0)

    try:
        plt.savefig("src/report1.png", bbox_inches='tight', dpi=400)
        print("ranking geral salvo (report1.png)")
    except:
        print("Erro em salvar ranking geral (report1.png)")


def plotar_ranking_filtrado():
    selected_dce = dce_var.get()
    selected_irce = irce_var.get()

    if selected_dce == '1ª DCE':
        df_filtrado = dfPlot[(dfPlot['dce'] == '1ª DCE') & (dfPlot['irce'] == selected_irce)]
        print(df_filtrado)
    elif selected_dce == '2ª DCE':
        df_filtrado = dfPlot[(dfPlot['dce'] == '2ª DCE') & (dfPlot['irce'] == selected_irce)]
        print(df_filtrado)
    else:
        messagebox.showerror("Erro", "Selecione uma DCE e uma IRCE válidas")
        return

    if df_filtrado.empty:
        messagebox.showerror("Erro", "Nenhum município encontrado")
        return

    df_filtrado = df_filtrado.sort_values(by='nota', ascending=False)

    fig = plt.figure(figsize=(5.5, 5.5))

    # Plotagem do ranking
    plt.barh(df_filtrado['municipio'], df_filtrado['nota'], color='orange', height=0.5)
    plt.gca().invert_yaxis()
    plt.xlabel('Nota', fontsize=12, color='white')
    plt.ylabel('Município', fontsize=12, color='white')
    plt.title(f'{selected_irce} ({selected_dce})', fontsize=12, color='white')
    plt.gca().set_facecolor("#3C91E6")  # Fundo do gráfico
    fig.patch.set_facecolor("#3C91E6")  # Fundo da figura
    plt.gca().tick_params(axis='y', colors='white')
    plt.gca().tick_params(axis='x', colors='white')
    plt.xlim([df_filtrado['nota'].min() - 1, df_filtrado['nota'].max() + 1])

    plt.tight_layout()

    # Limpando o frame antes de desenhar o novo gráfico
    for widget in frame_plotagem_ranking_filtrado.winfo_children():
        widget.destroy()

    # Adicionando o gráfico à interface
    canvas = FigureCanvasTkAgg(fig, master=frame_plotagem_ranking_filtrado)
    canvas.draw()
    canvas.get_tk_widget().grid(padx=10, pady=5, column=0, row=0)
    try:
        plt.savefig("src/report2.png", bbox_inches='tight', dpi=400)
        print("ranking filtrado salvo (report2.png)")
    except:
        print("Erro em salvar ranking filtrado (report2.png)")


dce_var = ctk.StringVar()
irce_var = ctk.StringVar()

dce_label = ctk.CTkLabel(frame_botoes_ranking_filtrado, text="DCE", text_color="white")
dce_menu = ctk.CTkOptionMenu(frame_botoes_ranking_filtrado, variable=dce_var, values=["1ª DCE", "2ª DCE"], command=lambda _: atualizar_irces())

irce_label = ctk.CTkLabel(frame_botoes_ranking_filtrado, text="IRCE", text_color="white")
irce_menu = ctk.CTkOptionMenu(frame_botoes_ranking_filtrado, variable=irce_var, values=[])

plotar_button = ctk.CTkButton(frame_botoes_ranking_filtrado, text="Plotar Ranking", command=plotar_ranking_filtrado)
report_download_button = ctk.CTkButton(frame_botoes_ranking_filtrado, text="Download Relatório", command=criarPDF)

def show_filter():
    dce_label.grid(padx=10, pady=5, row=0, column=0, sticky='w')
    dce_menu.grid(padx=10, pady=5, row=0, column=1, sticky='w')

    irce_label.grid(padx=10, pady=5, row=1, column=0, sticky='w')
    irce_menu.grid(padx=10, pady=5, row=1, column=1, sticky='w')

    plotar_button.grid(padx=20, pady=5, row=2, column=0, sticky='w')
    report_download_button.grid(padx=20, pady=5, row=2, column=1, sticky='w')
def hide_filter():
  dce_menu.grid_forget()
  irce_menu.grid_forget()
  plotar_button.grid_forget()
  report_download_button.grid_forget()

def mapa_cloropletico_bahia():
  global map_file
  # URL do GeoJSON
  geojson_url = 'https://raw.githubusercontent.com/tbrugz/geodata-br/refs/heads/master/geojson/geojs-29-mun.json'

  # Criar o mapa
  mapa_mun_bahia = folium.Map(location=[-12.9704, -38.5124], zoom_start=6, tiles='cartodbpositron')

  print(dfPlot)
  # Criar o choropleth
  folium.Choropleth(
      geo_data=geojson_url,
      data=dfPlot,
      columns=['id', 'nota'],
      key_on='feature.properties.id',
      fill_color='YlOrRd',
      fill_opacity=0.9,
      line_opacity=0.5,
      legend_name="Notas"
  ).add_to(mapa_mun_bahia)

  estilo = lambda x: {"fillColor": "white",
                   "color": "black",
                   "fillOpacity": 0.001,
                   "weight": 0.001}

  estilo_destaque = lambda x: {"fillColor": "darkblue",
                              "color": "black",
                              "fillOpacity": 0.5,
                              "weight": 1}

  highlight = folium.features.GeoJson(data = geojson_url,
                                    style_function = estilo,
                                    highlight_function = estilo_destaque,
                                    name = "Destaque")

  #Adicionando caixa de texto
  folium.features.GeoJsonTooltip(fields = ["name"],
                                aliases = ["municipio"],
                                labels = False,
                                style = ("background-color: white; color: black; font-family: arial; font-size: 16px; padding: 10px;")).add_to(highlight)

  #Adicionando o destaque ao mapa
  mapa_mun_bahia.add_child(highlight)

  mapa_mun_bahia.save(map_file)

def show_mapa_cloropletico():
    map_file = resource_path('dados\\mapa_cloropleto_bahia.html')
    map_url = 'file://' + os.path.abspath(str(map_file))

    if os.path.exists(map_file):
      webview.create_window('Mapa Cloropleto - Municípios da Bahia', str(map_file))
      webview.start()
    else:
      print(f"Erro: Arquivo {map_file} não encontrado!")
      messagebox.showerror("Erro: Arquivo Inexistente", "Erro!.\n Clique em Salvar para visualizar o mapa")

def dashboard():
    try:
       wb.close()
    finally:
      global dfPlot, irce_por_mun_dce1, irce_por_mun_dce2
      df = pd.read_excel(file, sheet_name='MATRIZ CONTRATOS')

      dfIDs = df.iloc[6:, 0]
      dfMunicipio = df.iloc[6:, 1]
      dfNota = df.iloc[6:, 84]
      dfIRCE = df.iloc[6:, 2]
      dfDCE = df.iloc[6:, 3]

      novo_df = {
        'id': dfIDs.values,
        'municipio': dfMunicipio.values,
        'irce': dfIRCE.values,
        'dce': dfDCE.values,
        'nota': dfNota.values
    }
      print(novo_df)
      dfPlot = pd.DataFrame(novo_df)
      dfPlot = dfPlot.sort_values(by='nota', ascending=False)
      print(dfPlot.head(50))
      dfPlot['irce'] = list(map(lambda x: " ".join(x.split()), list(dfPlot['irce'])))
      #print(f"IRCEs:\n {dfPlot['irce']}")
      print(f'id:{dfPlot["id"].head(50)},\nmunicipio:{dfPlot["municipio"].head(50)},\nnota:{dfPlot["nota"].head(50)}')
      dce_1, dce_2 = '1ª DCE', '2ª DCE'

      # IRCEs associadas à 1ª e 2ª DCE
      df_1_DCE = dfPlot[dfPlot['dce'] == dce_1]
      df_2_DCE = dfPlot[dfPlot['dce'] == dce_2]
      #print(df_1_DCE)
      #print(df_2_DCE)

      # Listas de IRCEs únicas para cada DCE
      irce_list_dce1 = df_1_DCE['irce'].drop_duplicates().tolist()
      irce_list_dce1_2 = list(map(lambda x: " ".join(x.split()), irce_list_dce1))
      irce_list_dce2 = df_2_DCE['irce'].drop_duplicates().tolist()
      irce_list_dce2_2 = list(map(lambda x: " ".join(x.split()), irce_list_dce2))

      #print(irce_list_dce1_2)
      #print(irce_list_dce2_2)

      # Criando dicionários com IRCEs como chaves e municípios como valores
      irce_por_mun_dce1 = {irce: df_1_DCE[df_1_DCE['irce'].str.strip() == irce.strip()]['municipio'].tolist() for irce in irce_list_dce1_2}
      irce_por_mun_dce2 = {irce: df_2_DCE[df_2_DCE['irce'].str.strip() == irce.strip()]['municipio'].tolist() for irce in irce_list_dce2_2}

      #print(irce_por_mun_dce1)
      #print(irce_por_mun_dce2)
      print(dfPlot)
      plotar_ranking_geral(dfPlot)

def atualizar_irces():
    selected_dce = dce_var.get()

    if selected_dce == '1ª DCE':
        mun_list_dce1 = list(irce_por_mun_dce1.keys())
        #mun_list_dce1_2 = list(map(lambda x: x.upper(), mun_list_dce1))
        irce_menu.configure(values=mun_list_dce1)
    elif selected_dce == '2ª DCE':
        mun_list_dce2 = list(irce_por_mun_dce2.keys())
        #mun_list_dce2_2 = list(map(lambda x: x.upper(), mun_list_dce2))
        irce_menu.configure(values=mun_list_dce2)
        irce_menu.set('')

    #-----------------------------------------------------------------------------------------------------------------------------------------
async def show_loading_text():
    for i in range(7):
        window.title(f"Sistema de Matriz de Seletividade (Carregando.)")
        await asyncio.sleep(0.3)
        window.title(f"Sistema de Matriz de Seletividade (Carregando..)")
        await asyncio.sleep(0.3)
        window.title(f"Sistema de Matriz de Seletividade (Carregando...)")
        await asyncio.sleep(0.3)


def save_file_and_refresh(file_modified):
    wb.save(file_modified)
    refresh_file(file_modified)

def hide_loading_text():
   window.title('Sistema de Matriz de Seletividade')

async def main_save_task():
    global loading_task
    try:
        loading_task = asyncio.create_task(show_loading_text())
        await asyncio.to_thread(save_file_and_refresh, file)
        messagebox.showinfo("Sucesso", "Alterações salvas com sucesso!", icon='info')
    except Exception as e:
        messagebox.showerror("Erro", f"Falha ao salvar: {e}")
    finally:
        loading_task.cancel()
        hide_loading_text()



fonte_botao=ctk.CTkFont("Arial",size=15,weight='bold')

def botao_salvar_config(frame_botoes):
    botao_salvar = ctk.CTkButton(frame_botoes, text="Salvar", command=botao_salvar_event, font = fonte_botao, fg_color="#2F83D7")
    botao_salvar.grid(pady=(10, 10), padx=20, sticky="w",row=10, column=0)

def botao_salvar_event():
    print("Botão salvar clicado")

    if validar_distribuicao():
        asyncio.run(main_save_task())
    else:
      messagebox.showerror("Erro", "Houve um erro ao salvar as alterações!\nVerifique se a soma de porcentagens é igual a 100%.", icon='error')

def botao_visualizar_dashboard_config(frame_botoes):
    botao_visualizar = ctk.CTkButton(frame_botoes, text="Ranking", command=botao_visualizar_dashboard_event, font=fonte_botao, fg_color="#2F83D7")
    botao_visualizar.grid(pady=(10, 10), padx=20, sticky="w",row=10,column=1)

def botao_visualizar_dashboard_event():
  if validar_distribuicao():
      asyncio.run(show_loading_text())
      hide_all()
      dashboard()
      show_filter()
      hide_loading_text()
  else:
      messagebox.showerror("Erro", "Salve as alterações para visualizar o Ranking")

def botao_visualizar_mapa_config(frame_botoes):
    botao_visualizar = ctk.CTkButton(frame_botoes, text="Mapa", command=botao_visualizar_mapa_event, font=fonte_botao, fg_color="#2F83D7")
    botao_visualizar.grid(pady=(10, 10), padx=20, sticky="w",row=10,column=2)

def botao_visualizar_mapa_event():
    if validar_distribuicao():
      mapa_cloropletico_bahia()
      show_mapa_cloropletico()
    else:
       messagebox.showerror("Erro", "Salve o Arquivo para visualizar o Mapa")

def botao_voltar_config(frame_botoes):
    botao_voltar = ctk.CTkButton(frame_botoes, text="Voltar", command=botao_voltar_event, font = fonte_botao, fg_color="#2F83D7")
    botao_voltar.grid(pady=(10, 10), padx=20, sticky="w",row=10,column=3)

def botao_voltar_event():
    show_all()
    hide_filter()

    canvas.get_tk_widget().destroy()
    frame_ranking_geral.pack_forget()

bloco_indicadores()
botao_salvar_config(frame_botoes)
botao_visualizar_dashboard_config(frame_botoes)
botao_voltar_config(frame_botoes)
botao_visualizar_mapa_config(frame_botoes)
salvar_indicadores(value=0)

def on_motion(event):
    window.update_idletasks()

window.bind("<B1-Motion>", on_motion)
window.wm_attributes("-transparentcolor", "")

MainWindow.window_config(window)
window.protocol("WM_DELETE_WINDOW", on_closing)
window.mainloop()