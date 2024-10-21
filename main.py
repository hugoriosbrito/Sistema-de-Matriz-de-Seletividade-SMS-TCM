import folium.features
import openpyxl as xl
import customtkinter as ctk
from tkinter import messagebox
import pandas as pd
import matplotlib.pyplot as plt
from matplotlib.backends.backend_tkagg import FigureCanvasTkAgg
import os, sys, win32com.client
import folium
import webview

icon ="_internal\src\icon.ico"
map_file = "_internal\dados\mapa_cloropleto_bahia.html"
file = "_internal\dados\Matriz modelo - VERSÃO SISTEMA.xlsx"
wb = xl.load_workbook(file)
sheet = wb['SÍNTESE']

window = ctk.CTk()
fonte_geral_texto = ctk.CTkFont(family='Arial', size=15, weight='bold')

# Função para encontrar o caminho do arquivo
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
        ctk.set_appearance_mode("light")  # Modo de aparência escuro
        ctk.set_default_color_theme("dark-blue")
        window.configure(fg_color="#3C91E6")
        window.title("Sistema de Gerenciamento de Indicadores")
        window.iconbitmap(icon)

fonte_titulo = ctk.CTkFont(family='Arial', size=40, weight='bold')
titulo = ctk.CTkLabel(window, text="Sistema de Gerenciamento de Indicadores", font=fonte_titulo, anchor="center", corner_radius=20, text_color="white")
titulo.pack(pady=20, padx=20, anchor="center")  # Pequeno espaçamento nas laterais do título

frame_botoes = ctk.CTkFrame(master=window, border_width=0, corner_radius=20, bg_color="#2F83D7", fg_color="transparent", height=50)
frame_botoes.pack(fill='both', padx=20)

frame_dist_peso = ctk.CTkFrame(master=window, border_width=0, corner_radius=20, bg_color="#2F83D7", fg_color="#2F83D7", height=20)
frame_dist_peso.pack(fill = 'both', padx = 20)

frame = ctk.CTkScrollableFrame(master=window, border_width=0, corner_radius=20, bg_color="#2F83D7", fg_color="#2F83D7", height=250, scrollbar_button_color="white")
frame.pack(fill='both', padx=20,pady=10,expand=1)

frame.grid_columnconfigure(0,weight=1)
frame.grid_columnconfigure(1,weight=1)
frame.grid_columnconfigure(2,weight=1)
frame.grid_columnconfigure(3,weight=1)

#Alterações na distribuição de peso

distribuicao_fonte = ctk.CTkFont(family='Arial', size=15, weight='bold')
distribuicao_titulo = ctk.CTkLabel(master=frame_dist_peso, text= "Distribuição de peso \n por tipo ", font=distribuicao_fonte, text_color='white', corner_radius=20, anchor="center")
distribuicao_titulo.grid(padx=(20,20),pady=10, row = 0, column=0)
valores = ["5", "10", "15", "20", "25", "30", "35", "40", "45", "50", "55", "60", "65", "70", "75", "80", "85", "90", "95", "100"]


tipos_fonte = ctk.CTkFont(family='Arial', size= 15, weight= "bold")
# todas as vezes que um valor é escolhido em um  combo, ele é adicionado à uma lista e somado para verificação, mensagem de verificação
# apenas para quando a lista tiver 4 elementos
# se a soma for igual a 100, a mensagem é de sucesso, e é salvo, se não, a mensagem é de erro apenas, sem salvamento

def validar_distribuicao():
  list_pesos = salvar_indicadores(value=0)
  if sum(list_pesos) == 100: 
    return True
  else:
    return False   

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

    fonte_colunas = ctk.CTkFont(family='Arial', size=15, weight='bold',)
    
    coluna_risco = ctk.CTkLabel(master=frame, text= "RISCO", font=fonte_colunas, text_color='white', corner_radius=20)
    coluna_risco.grid(padx=10,pady=5,row=0,column=0)

    coluna_relevancia = ctk.CTkLabel(master=frame, text= "RELEVÂNCIA", font=fonte_colunas, text_color='white', corner_radius=20)
    coluna_relevancia.grid(padx=10,pady=5,row=0,column=1)

    coluna_materialidade = ctk.CTkLabel(master=frame, text= "MATERIALIDADE", font=fonte_colunas, text_color='white', corner_radius=20)
    coluna_materialidade.grid(padx=10,pady=5,row=0,column=2)

    coluna_oportunidade = ctk.CTkLabel(master=frame, text= "OPORTUNIDADE", font=fonte_colunas, text_color='white', corner_radius=20)
    coluna_oportunidade.grid(padx=0,pady=5,row=0,column=3)


#TIPOS DE RISCO PRIMEIRA COLUNA

    # HISTÓRICO PARECER PRÉVIO (ÚLTIMOS 3 ANOS)
    indicador1_frame = ctk.CTkFrame(frame, fg_color="#FAFFFD", corner_radius=10,width=200,height=100)
    indicador1_frame.grid(padx=20, pady=10, sticky="nsew",column=0)

    # Botão Switch event
    def botao_switch1_event():
        print("Switch 1 toggled, current value:", switch1_var.get())
        switch1.configure(text=f"Habilitar? ({switch1_var.get()})")

        # Habilitar/Desabilitar indicador
        if switch1_var.get() == "SIM":
            sheet['F11'] = 'SIM'
        else:
            sheet['F11'] = 'NÃO'

    # Título indicador
    indicador1_title = ctk.CTkLabel(indicador1_frame, text="HISTÓRICO PARECER PRÉVIO (ÚLTIMOS 3 ANOS)", text_color="black", corner_radius=20, anchor="w")
    indicador1_title.grid(pady=5, sticky="w")  

    # Switch referente ao indicador
    switch1_var = ctk.StringVar(value="NÃO")
    switch1 = ctk.CTkSwitch(indicador1_frame, text=f"Habilitar? ({switch1_var.get()})", command=botao_switch1_event, variable=switch1_var, onvalue="SIM", offvalue="NÃO", text_color="black")
    botao_switch1_event()
    switch1.grid(padx=10, pady=5, sticky="w") 


    # QTDE DE DÉBITO/MULTAS
    indicador2_frame = ctk.CTkFrame(frame, fg_color="#FAFFFD", corner_radius=10)
    indicador2_frame.grid(padx=20, pady=10, sticky="nsew",column=0)

    def botao_switch2_event():
        print("Switch 2 toggled, current value:", switch2_var.get())
        switch2.configure(text=f"Habilitar? ({switch2_var.get()})")

        # Habilitar/Desabilitar indicador
        if switch2_var.get() == "SIM":
            sheet['F13'] = 'SIM'
        else:
            sheet["F13"] = 'NÃO'

    indicador2_title = ctk.CTkLabel(indicador2_frame, text="QTDE DE DÉBITO/MULTAS", text_color="black", corner_radius=20, anchor="w")
    indicador2_title.grid(pady=5, sticky="w")

    switch2_var = ctk.StringVar(value="NÃO")
    switch2 = ctk.CTkSwitch(indicador2_frame, text=f"Habilitar? ({switch2_var.get()})", command=botao_switch2_event, variable=switch2_var, onvalue="SIM", offvalue="NÃO", text_color="black")
    botao_switch2_event()
    switch2.grid(padx=10, pady=5, sticky="w") 

    # ÍNDICE DE TRANSPARÊNCIA PÚBLICA
    indicador3_frame = ctk.CTkFrame(frame, fg_color="#FAFFFD", corner_radius=10)
    indicador3_frame.grid(padx=20, pady=10, sticky="nsew",column=0)

    def botao_switch3_event():
        print("Switch 3 toggled, current value:", switch3_var.get())
        switch3.configure(text=f"Habilitar? ({switch3_var.get()})")

        if switch3_var.get() == "SIM":
          sheet['F15'] = 'SIM'
        else:
          sheet['F15'] = 'NÃO'

    indicador3_title = ctk.CTkLabel(indicador3_frame, text="ÍNDICE DE TRANSPARÊNCIA PÚBLICA", text_color="black", corner_radius=20, anchor="w")
    indicador3_title.grid(pady=5, sticky="w")

    switch3_var = ctk.StringVar(value="NÃO")
    switch3 = ctk.CTkSwitch(indicador3_frame, text=f"Habilitar? ({switch3_var.get()})", command=botao_switch3_event, variable=switch3_var, onvalue="SIM", offvalue="NÃO", text_color="black")
    botao_switch3_event()
    switch3.grid(padx=10, pady=5, sticky="w") 

    #PERFIL DE CONTRATAÇÃO DO ENTE
    
    indicador4_frame = ctk.CTkFrame(frame, fg_color="#FAFFFD", corner_radius=10)
    indicador4_frame.grid(padx=20, pady=10, sticky="nsew",column=0)

    def botao_switch4_event():
        print("Switch 4 toggled, current value:", switch4_var.get())
        switch4.configure(text=f"Habilitar? ({switch4_var.get()})")

        if switch4_var.get() == "SIM":
          sheet['F16'] = 'SIM'
        else:
          sheet['F16'] = 'NÃO'

    indicador4_title = ctk.CTkLabel(indicador4_frame, text="PERFIL DE CONTRATAÇÃO DO ENTE", text_color="black", corner_radius=20, anchor="w")
    indicador4_title.grid(pady=5, sticky="w")

    switch4_var = ctk.StringVar(value="NÃO")
    switch4 = ctk.CTkSwitch(indicador4_frame, text=f"Habilitar? ({switch4_var.get()})", command=botao_switch4_event, variable=switch4_var, onvalue="SIM", offvalue="NÃO", text_color="black")
    botao_switch4_event()
    switch4.grid(padx=10, pady=5, sticky="w") 

    # QTDE DE DENÚNCIAS E REPRESENTAÇÕES  (ÚLTIMOS 5 ANOS)
    
    indicador5_frame = ctk.CTkFrame(frame, fg_color="#FAFFFD", corner_radius=10)
    indicador5_frame.grid(padx=20, pady=10, sticky="nsew",column=0)

    def botao_switch5_event():
        print("Switch 5 toggled, current value:", switch5_var.get())
        switch5.configure(text=f"Habilitar? ({switch5_var.get()})")

        if switch5_var.get() == "SIM":
          sheet['F21'] = 'SIM'
        else:
          sheet['F21'] = 'NÃO'

    indicador5_title = ctk.CTkLabel(indicador5_frame, text="QTDE DE DENÚNCIAS E REPRESENTAÇÕES  (ÚLTIMOS 5 ANOS)", text_color="black", corner_radius=20, anchor="w")
    indicador5_title.grid(pady=5, sticky="w")

    switch5_var = ctk.StringVar(value="NÃO")
    switch5 = ctk.CTkSwitch(indicador5_frame, text=f"Habilitar? ({switch5_var.get()})", command=botao_switch5_event, variable=switch5_var, onvalue="SIM", offvalue="NÃO", text_color="black")
    botao_switch5_event()
    switch5.grid(padx=10, pady=5, sticky="w") 

    # QTDE DE TOC  (ÚLTIMOS 5 ANOS)
    
    indicador6_frame = ctk.CTkFrame(frame, fg_color="#FAFFFD", corner_radius=10)
    indicador6_frame.grid(padx=20, pady=10, sticky="nsew",column=0)

    def botao_switch6_event():
        print("Switch 6 toggled, current value:", switch6_var.get())
        switch6.configure(text=f"Habilitar? ({switch6_var.get()})")

        if switch6_var.get() == "SIM":
          sheet['F22'] = 'SIM'
        else:
          sheet['F22'] = 'NÃO'

    indicador6_title = ctk.CTkLabel(indicador6_frame, text="QTDE DE TOC  (ÚLTIMOS 5 ANOS)", text_color="black", corner_radius=20, anchor="w")
    indicador6_title.grid(pady=5, sticky="w")

    switch6_var = ctk.StringVar(value="NÃO")
    switch6 = ctk.CTkSwitch(indicador6_frame, text=f"Habilitar? ({switch6_var.get()})", command=botao_switch6_event, variable=switch6_var, onvalue="SIM", offvalue="NÃO", text_color="black")
    botao_switch6_event()
    switch6.grid(padx=10, pady=5, sticky="w") 

    # QTDE DE TCE  (ÚLTIMOS 5 ANOS)

    indicador7_frame = ctk.CTkFrame(frame, fg_color="#FAFFFD", corner_radius=10)
    indicador7_frame.grid(padx=20, pady=10, sticky="nsew",column=0)

    def botao_switch7_event():
        print("Switch 7 toggled, current value:", switch7_var.get())
        switch7.configure(text=f"Habilitar? ({switch7_var.get()})")

        if switch7_var.get() == "SIM":
          sheet['F23'] = 'SIM'
        else:
          sheet['F23'] = 'NÃO'

    indicador7_title = ctk.CTkLabel(indicador7_frame, text="QTDE DE TCE  (ÚLTIMOS 5 ANOS)", text_color="black", corner_radius=20, anchor="w")
    indicador7_title.grid(pady=5, sticky="w")

    switch7_var = ctk.StringVar(value="NÃO")
    switch7 = ctk.CTkSwitch(indicador7_frame, text=f"Habilitar? ({switch7_var.get()})", command=botao_switch7_event, variable=switch7_var, onvalue="SIM", offvalue="NÃO", text_color="black")
    botao_switch7_event()
    switch7.grid(padx=10, pady=5, sticky="w") 


    # QTDE DE AUDITORIAS  (ÚLTIMOS 5 ANOS)

    indicador8_frame = ctk.CTkFrame(frame, fg_color="#FAFFFD", corner_radius=10)
    indicador8_frame.grid(padx=20, pady=10, sticky="nsew",column=0)

    def botao_switch8_event():
        print("Switch 8 toggled, current value:", switch8_var.get())
        switch8.configure(text=f"Habilitar? ({switch8_var.get()})")

        if switch8_var.get() == "SIM":
          sheet['F24'] = 'SIM'
        else:
          sheet['F24'] = 'NÃO'

    indicador8_title = ctk.CTkLabel(indicador8_frame, text="QTDE DE AUDITORIAS  (ÚLTIMOS 5 ANOS)", text_color="black", corner_radius=20, anchor="w")
    indicador8_title.grid(pady=5, sticky="w")

    switch8_var = ctk.StringVar(value="NÃO")
    switch8 = ctk.CTkSwitch(indicador8_frame, text=f"Habilitar? ({switch8_var.get()})", command=botao_switch8_event, variable=switch8_var, onvalue="SIM", offvalue="NÃO", text_color="black")
    botao_switch8_event()
    switch8.grid(padx=10, pady=5, sticky="w") 

    # QTDE DE  MEDIDAS CAUTELARES  (ÚLTIMOS 5 ANOS)

    indicador9_frame = ctk.CTkFrame(frame, fg_color="#FAFFFD", corner_radius=10)
    indicador9_frame.grid(padx=20, pady=10, sticky="nsew",column=0)

    def botao_switch9_event():
        print("Switch 9 toggled, current value:", switch9_var.get())
        switch9.configure(text=f"Habilitar? ({switch9_var.get()})")

        if switch9_var.get() == "SIM":
          sheet['F25'] = 'SIM'
        else:
          sheet['F25'] = 'NÃO'

    indicador9_title = ctk.CTkLabel(indicador9_frame, text="QTDE DE  MEDIDAS CAUTELARES  (ÚLTIMOS 5 ANOS)", text_color="black", corner_radius=20, anchor="w")
    indicador9_title.grid(pady=5, sticky="w")

    switch9_var = ctk.StringVar(value="NÃO")
    switch9 = ctk.CTkSwitch(indicador9_frame, text=f"Habilitar? ({switch9_var.get()})", command=botao_switch9_event, variable=switch9_var, onvalue="SIM", offvalue="NÃO", text_color="black")
    botao_switch9_event()
    switch9.grid(padx=10, pady=5, sticky="w") 


#TIPOS DE RELEVÂNCIA SEGUNDA COLUNA

    # POPULAÇÃO MUNICÍPIO

    indicador10_frame = ctk.CTkFrame(frame, fg_color="#FAFFFD", corner_radius=10)
    indicador10_frame.grid(padx=20, pady=10, sticky="nsew", row = 1, column=1)

    def botao_switch10_event():
        print("Switch 10 toggled, current value:", switch10_var.get())
        switch10.configure(text=f"Habilitar? ({switch10_var.get()})")

        if switch10_var.get() == "SIM":
          sheet['F17'] = 'SIM'
        else:
          sheet['F17'] = 'NÃO'

    indicador10_title = ctk.CTkLabel(indicador10_frame, text="POPULAÇÃO MUNICÍPIO", text_color="black", corner_radius=20, anchor="w")
    indicador10_title.grid(pady=5, sticky="w")

    switch10_var = ctk.StringVar(value="NÃO")
    switch10 = ctk.CTkSwitch(indicador10_frame, text=f"Habilitar? ({switch10_var.get()})", command=botao_switch10_event, variable=switch10_var, onvalue="SIM", offvalue="NÃO", text_color="black")
    botao_switch10_event()
    switch10.grid(padx=10, pady=5, sticky="w") 

    # IDH

    indicador11_frame = ctk.CTkFrame(frame, fg_color="#FAFFFD", corner_radius=10)
    indicador11_frame.grid(padx=20, pady=10, sticky="nsew", row = 2, column=1)

    def botao_switch11_event():
        print("Switch 11 toggled, current value:", switch11_var.get())
        switch11.configure(text=f"Habilitar? ({switch11_var.get()})")

        if switch11_var.get() == "SIM":
          sheet['F18'] = 'SIM'
        else:
          sheet['F18'] = 'NÃO'

    indicador11_title = ctk.CTkLabel(indicador11_frame, text="IDH", text_color="black", corner_radius=20, anchor="w")
    indicador11_title.grid(pady=5, sticky="w")

    switch11_var = ctk.StringVar(value="NÃO")
    switch11 = ctk.CTkSwitch(indicador11_frame, text=f"Habilitar? ({switch11_var.get()})", command=botao_switch11_event, variable=switch11_var, onvalue="SIM", offvalue="NÃO", text_color="black")
    botao_switch11_event()
    switch11.grid(padx=10, pady=5, sticky="w")

    # IEGM

    indicador12_frame = ctk.CTkFrame(frame, fg_color="#FAFFFD", corner_radius=10)
    indicador12_frame.grid(padx=20, pady=10, sticky="nsew", row = 3, column=1)

    def botao_switch12_event():
        print("Switch 12 toggled, current value:", switch12_var.get())
        switch12.configure(text=f"Habilitar? ({switch12_var.get()})")

        if switch12_var.get() == "SIM":
          sheet['F19'] = 'SIM'
        else:
          sheet['F19'] = 'NÃO'

    indicador12_title = ctk.CTkLabel(indicador12_frame, text="IEGM", text_color="black", corner_radius=20, anchor="w")
    indicador12_title.grid(pady=5, sticky="w")

    switch12_var = ctk.StringVar(value="NÃO")
    switch12 = ctk.CTkSwitch(indicador12_frame, text=f"Habilitar? ({switch12_var.get()})", command=botao_switch12_event, variable=switch12_var, onvalue="SIM", offvalue="NÃO", text_color="black")
    botao_switch12_event()
    switch12.grid(padx=10, pady=5, sticky="w")

    # IDTRU-DL
 
    indicador13_frame = ctk.CTkFrame(frame, fg_color="#FAFFFD", corner_radius=10)
    indicador13_frame.grid(padx=20, pady=10, sticky="nsew", row = 4, column=1)

    def botao_switch13_event():
        print("Switch 13 toggled, current value:", switch13_var.get())
        switch13.configure(text=f"Habilitar? ({switch13_var.get()})")

        if switch13_var.get() == "SIM":
          sheet['F20'] = 'SIM'
        else:
          sheet['F20'] = 'NÃO'

    indicador13_title = ctk.CTkLabel(indicador13_frame, text="IDTRU-DL", text_color="black", corner_radius=20, anchor="w")
    indicador13_title.grid(pady=5, sticky="w")

    switch13_var = ctk.StringVar(value="NÃO")
    switch13 = ctk.CTkSwitch(indicador13_frame, text=f"Habilitar? ({switch13_var.get()})", command=botao_switch13_event, variable=switch13_var, onvalue="SIM", offvalue="NÃO", text_color="black")
    botao_switch13_event()
    switch13.grid(padx=10, pady=5, sticky="w")

#TIPOS DE MATERIALIDADE TERCEIRA COLUNA

    #VALOR DE DÉBITO E MULTAS

    indicador14_frame = ctk.CTkFrame(frame, fg_color="#FAFFFD", corner_radius=10)
    indicador14_frame.grid(padx=20, pady=10, sticky="nsew", row = 1, column=2)

    def botao_switch14_event():
        print("Switch 14 toggled, current value:", switch14_var.get())
        switch14.configure(text=f"Habilitar? ({switch14_var.get()})")

        if switch14_var.get() == "SIM":
          sheet['F14'] = 'SIM'
        else:
          sheet['F14'] = 'NÃO'

    indicador14_title = ctk.CTkLabel(indicador14_frame, text="VALOR DE DÉBITO E MULTAS", text_color="black", corner_radius=20, anchor="w")
    indicador14_title.grid(pady=5, sticky="w")

    switch14_var = ctk.StringVar(value="NÃO")
    switch14 = ctk.CTkSwitch(indicador14_frame, text=f"Habilitar? ({switch14_var.get()})", command=botao_switch14_event, variable=switch14_var, onvalue="SIM", offvalue="NÃO", text_color="black")
    botao_switch14_event()
    switch14.grid(padx=10, pady=5, sticky="w")

#TIPOS DE OPORTUNIDADE QUARTA COLUNA

    #DATA ÚLTIMA AUDITORIA (3DCE)
    
    indicador15_frame = ctk.CTkFrame(frame, fg_color="#FAFFFD", corner_radius=10)
    indicador15_frame.grid(padx=20, pady=10, sticky="nsew", row = 1, column=3)

    def botao_switch15_event():
        print("Switch 15 toggled, current value:", switch15_var.get())
        switch15.configure(text=f"Habilitar? ({switch15_var.get()})")

        if switch15_var.get() == "SIM":
          sheet['F12'] = 'SIM'
        else:
          sheet['F12'] = 'NÃO'

    indicador15_title = ctk.CTkLabel(indicador15_frame, text="DATA ÚLTIMA AUDITORIA (3DCE)", text_color="black", corner_radius=20, anchor="w")
    indicador15_title.grid(pady=5, sticky="w")

    switch15_var = ctk.StringVar(value="NÃO")
    switch15 = ctk.CTkSwitch(indicador15_frame, text=f"Habilitar? ({switch15_var.get()})", command=botao_switch15_event, variable=switch15_var, onvalue="SIM", offvalue="NÃO", text_color="black")
    botao_switch15_event()
    switch15.grid(padx=10, pady=5, sticky="w")

# -------------------------------------------- FIM DO BLOCO DE INDICADORES ---------------------------------------------------------------

#-----------------------------------------------------------------------------------------------------------------------------------------

def refresh_file(file):
    xlapp = win32com.client.DispatchEx("Excel.Application")
    path = os.path.abspath(file)
    wb =  xlapp.Workbooks.Open(path)
    wb.RefreshAll()
    xlapp.CalculateUntilAsyncQueriesDone()
    wb.Save()
    xlapp.Quit()

def hide_all():
    frame.pack_forget()
    frame_dist_peso.pack_forget()
    frame_ranking_geral.pack_forget()  
def show_all():
    frame_dist_peso.pack(fill='both', padx=20)
    frame.pack(fill='both', padx=20, pady=10, expand=1)

frame_ranking_geral = ctk.CTkFrame(master=window,fg_color='#3C91E6')
frame_plotagem_ranking_geral = ctk.CTkFrame(master=frame_ranking_geral, fg_color='#3C91E6')

filter_button_frame = ctk.CTkFrame(master=frame_ranking_geral, fg_color="#3C91E6")
filter_button_frame.grid(padx=10, pady=10,sticky="nsew",column=1)

frame_plotagem_ranking_filtrado = ctk.CTkFrame(master=filter_button_frame, fg_color='#3C91E6')
frame_plotagem_ranking_filtrado.grid(padx=10, column=1, row=3)


def plotar_ranking_geral(dfPlot):
    global canvas, frame_ranking_geral
    dfTop50 = dfPlot.head(50)
    
    # Garantindo que o ranking geral será mostrado ao plotar
    frame_ranking_geral.pack(expand=True,fill='both')
    frame_plotagem_ranking_geral.grid(padx=10,row=0,column=0)

    fig = plt.figure(figsize=(8, 7.5))

    # Plotagem do ranking
    plt.barh(dfTop50['municipio'], dfTop50['nota'], color='#D03645',height=0.5)
    plt.gca().invert_yaxis()

    plt.xlabel('Nota', fontsize=12, color='white')
    plt.ylabel('Município', fontsize=12, color='white')
    plt.title('Top 50 Municípios por Nota', fontsize=14, color='white')

    plt.gca().set_facecolor("#3C91E6")
    fig.patch.set_facecolor("#3C91E6")

    plt.gca().tick_params(axis='y', colors='white')
    plt.gca().tick_params(axis='x', colors='white')

    plt.tight_layout()
    #plt.autoscale(enable=True, axis='both')

    
    # Limpando o frame antes de desenhar o novo gráfico
    for widget in frame_plotagem_ranking_geral.winfo_children():
        widget.destroy()


    canvas = FigureCanvasTkAgg(fig, master=frame_plotagem_ranking_geral)
    canvas.draw()
    canvas.get_tk_widget().grid(padx=20,sticky='nsew', pady=10,column=0)

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
        messagebox.showerror("Erro", "Nenhum município encontrado")
        return

    df_filtrado = df_filtrado.sort_values(by='nota', ascending=False)

    fig = plt.figure(figsize=(5.5, 5.5))

    # Plotagem do ranking
    plt.barh(df_filtrado['municipio'], df_filtrado['nota'], color='orange', height=0.5)
    plt.gca().invert_yaxis()
    plt.xlabel('Nota', fontsize=12, color='white')
    plt.ylabel('Município', fontsize=12, color='white')
    plt.title(f'      Municípios da IRCE {selected_irce} ({selected_dce})', fontsize=15, color='white')
    plt.gca().set_facecolor("#3C91E6")  # Fundo do gráfico
    fig.patch.set_facecolor("#3C91E6")  # Fundo da figura
    plt.gca().tick_params(axis='y', colors='white')
    plt.gca().tick_params(axis='x', colors='white')
    plt.tight_layout()

    # Limpando o frame antes de desenhar o novo gráfico
    for widget in frame_plotagem_ranking_filtrado.winfo_children():
        widget.destroy()

    # Adicionando o gráfico à interface
    canvas = FigureCanvasTkAgg(fig, master=frame_plotagem_ranking_filtrado)
    canvas.draw()
    canvas.get_tk_widget().grid(padx=10, pady=10,column=3,row=4)



# Variável para armazenar a DCE selecionada
dce_var = ctk.StringVar()
#Botão de lista para selecionar 1ª ou 2ª DCE
dce_text= ctk.CTkLabel(master=filter_button_frame,text='DCE', font=fonte_geral_texto,text_color='white')
dce_menu = ctk.CTkOptionMenu(master=filter_button_frame, variable=dce_var, values=['1ª DCE', '2ª DCE'], command=lambda _: atualizar_irces())
# Botão de lista para mostrar IRCEs de acordo com a DCE
irce_var = ctk.StringVar()
irce_text= ctk.CTkLabel(master=filter_button_frame,text='IRCE', font=fonte_geral_texto, text_color='white')
irce_menu = ctk.CTkOptionMenu(master=filter_button_frame, variable=irce_var, values=[])
# Botão para plotar o gráfico baseado na filtragem
plotar_button = ctk.CTkButton(master=filter_button_frame, text="Plotar Ranking", command=plotar_ranking_filtrado)

def show_filter():
  dce_text.grid(padx=10,column=0,row=0)
  dce_menu.grid(padx=10,pady=5, column=1,row=0,sticky='e')
  irce_text.grid(padx=10,column=0,row=1)
  irce_menu.grid(padx=10,pady=5, column=1,row=1,sticky='e')
  plotar_button.grid(padx=10,pady=5, column=1,row=2,sticky='e')
def hide_filter():
  dce_menu.grid_forget()
  irce_menu.grid_forget()
  plotar_button.grid_forget()

def mapa_cloropletico_bahia():
  global map_file
  # URL do GeoJSON
  geojson_url = 'https://raw.githubusercontent.com/tbrugz/geodata-br/refs/heads/master/geojson/geojs-29-mun.json'

  # Criar o mapa
  mapa_mun_bahia = folium.Map(location=[-12.9704, -38.5124], zoom_start=6, tiles='cartodbpositron')

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
    map_file = resource_path('dados\mapa_cloropleto_bahia.html')
    map_url = 'file://' + os.path.abspath(map_file)

    if os.path.exists(map_file):
      webview.create_window('Mapa Cloropleto - Municípios da Bahia', map_file)
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
      dfNota = df.iloc[6:, 34]
      dfIRCE = df.iloc[6:, 2]
      dfDCE = df.iloc[6:, 3]

      novo_df = {
        'id': dfIDs.values,
        'municipio': dfMunicipio.values,
        'irce': dfIRCE.values,
        'dce': dfDCE.values,
        'nota': dfNota.values
    }

      dfPlot = pd.DataFrame(novo_df)
      dfPlot = dfPlot.sort_values(by='nota', ascending=False)
      #print(f'id:{novo_df["id"]},\nmunicipio:{novo_df["municipio"]},\nnota:{novo_df["nota"]}')
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

      plotar_ranking_geral(dfPlot)

def atualizar_irces():
    selected_dce = dce_var.get()

    if selected_dce == '1ª DCE':
        irce_menu.configure(values=list(irce_por_mun_dce1.keys()))
    elif selected_dce == '2ª DCE':
        irce_menu.configure(values=list(irce_por_mun_dce2.keys()))
        irce_menu.set('                                        ')  # Limpar a seleção da IRCE

#-----------------------------------------------------------------------------------------------------------------------------------------

def show_loading_text():
   window.title('Sistema de Gerenciamento de Indicadores (Carregando...)')
   
def hide_loading_text():
   window.title('Sistema de Gerenciamento de Indicadores')

fonte_botao=ctk.CTkFont("Arial",size=15,weight='bold')

class Botao:
    def botao_salvar_config(frame_botoes):
        botao_salvar = ctk.CTkButton(frame_botoes, text="Salvar", command=Botao.botao_salvar_event, font = fonte_botao, fg_color="#2F83D7")
        botao_salvar.grid(pady=(10, 10), padx=20, sticky="w",row=10, column=0)  

    def botao_salvar_event():
        print("Botão salvar clicado")
        if validar_distribuicao():
          try:
            show_loading_text()
          finally:
            wb.save(file)
            refresh_file(file)
            hide_loading_text()
            messagebox.showinfo("Sucesso", "Alterações salvas com sucesso!", icon='info')
        else:
          messagebox.showerror("Erro", "Houve um erro ao salvar as alterações!\nVerifique se a soma de porcentagens é igual a 100%.", icon='error')

    def botao_visualizar_dashboard_config(frame_botoes):
        botao_visualizar = ctk.CTkButton(frame_botoes, text="Ranking", command=Botao.botao_visualizar_dashboard_event, font=fonte_botao, fg_color="#2F83D7")
        botao_visualizar.grid(pady=(10, 10), padx=20, sticky="w",row=10,column=1)  

    def botao_visualizar_dashboard_event():
      if validar_distribuicao():
          show_loading_text()
          hide_all()
          dashboard()
          show_filter()
          hide_loading_text() 
      else:
          messagebox.showerror("Erro", "Salve as alterações para visualizar o Ranking")
                                
    def botao_visualizar_mapa_config(frame_botoes):
        botao_visualizar = ctk.CTkButton(frame_botoes, text="Mapa", command=Botao.botao_visualizar_mapa_event, font=fonte_botao, fg_color="#2F83D7")
        botao_visualizar.grid(pady=(10, 10), padx=20, sticky="w",row=10,column=2)  

    def botao_visualizar_mapa_event():
        if validar_distribuicao():
          mapa_cloropletico_bahia()
          show_mapa_cloropletico()
        else:
           messagebox.showerror("Erro", "Salve o Arquivo para visualizar o Mapa")

    def botao_voltar_config(frame_botoes):
        botao_voltar = ctk.CTkButton(frame_botoes, text="Voltar", command=Botao.botao_voltar_event, font = fonte_botao, fg_color="#2F83D7")
        botao_voltar.grid(pady=(10, 10), padx=20, sticky="w",row=10,column=3) 

    def botao_voltar_event():
      show_all()
      hide_filter()
      
      canvas.get_tk_widget().destroy()
      frame_ranking_geral.pack_forget()  

      

def main(): 
    bloco_indicadores()
    Botao.botao_salvar_config(frame_botoes)
    Botao.botao_visualizar_dashboard_config(frame_botoes)
    Botao.botao_voltar_config(frame_botoes)
    Botao.botao_visualizar_mapa_config(frame_botoes)

    MainWindow.window_config(window)



main()
window.mainloop()
window.protocol("WM_DELETE_WINDOW", on_closing())
try:
  exit()
finally:
  if os.path.exists(map_file):
      os.remove(map_file)