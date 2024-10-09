import openpyxl as xl
import customtkinter as ctk
from tkinter import messagebox
import tkinter as tk
from tkinter import ttk
import os
import time

#corrigir falhas em relação a migração de arquivo base
#adicionar os outros indicadores com respectivas funcionalidades
#adicionar distribuição de peso por tipo
#adicionar validações de distribuição
#adicionar restante das porcentagens


file = "dados\\Matriz modelo - VERSÃO SISTEMA.xlsx"
wb = xl.load_workbook(file, )
sheet = wb['SÍNTESE']


class xlsx:
    def xlsx_state(self):
        diretorio_atual = os.path.dirname(os.path.abspath(__file__))
        nome_arquivo = file
        pasta_arquivo = os.path.join(diretorio_atual, "dados")
        os.makedirs(pasta_arquivo, exist_ok=True)
        path = os.path.join(pasta_arquivo, nome_arquivo)
        exists = os.path.exists(path)

        if exists:
            return path

#Classe para obter o caminho do arquivo .xlsx
class XlsxPath(xlsx):
    def get_path(self):
        return xlsx.xlsx_state(self)

#destroi a instancia criada pelo openpyxl
def on_closing():
    try:
        wb.close()
    finally:
        window.destroy()

window = ctk.CTk()

class MainWindow:
    def window_config(window):
        window.geometry("1080x720")
        ctk.set_appearance_mode("dark")  # Modo de aparência escuro
        ctk.set_default_color_theme("dark-blue")
        window.configure(fg_color="#3C91E6")

def titulo():
    fonte_titulo = ctk.CTkFont(family='Arial', size=50, weight='bold')
    titulo = ctk.CTkLabel(window, text="Indicadores", font=fonte_titulo, anchor="center", corner_radius=20, text_color="white")
    titulo.pack(pady=20, padx=20, anchor="center")  # Pequeno espaçamento nas laterais do título

titulo()

frame_dist_peso = ctk.CTkFrame(master=window, border_width=0, corner_radius=20, bg_color="#2F83D7", fg_color="#2F83D7", height=50)
frame_dist_peso.pack(fill = 'both', padx = 20)

frame = ctk.CTkScrollableFrame(master=window, border_width=0, corner_radius=20, bg_color="#2F83D7", fg_color="#2F83D7", height=600, scrollbar_button_color="white")
frame.pack(fill='both', padx=20, pady=10)

frame.grid_columnconfigure(0,weight=1)
frame.grid_columnconfigure(1,weight=1)
frame.grid_columnconfigure(2,weight=1)
frame.grid_columnconfigure(3,weight=1)

#Alterações na distribuição de peso

def alterar_distribuicao():
    pass
def distribuicao():
    distribuicao_fonte = ctk.CTkFont(family='Arial', size=15, weight='bold')
    distribuicao_titulo = ctk.CTkLabel(master=frame_dist_peso, text= "Distribuição de peso por tipo: ", font=distribuicao_fonte, text_color='white', corner_radius=20)
    distribuicao_titulo.grid(padx=10,pady=10, row = 0, column=0)

    tipo_risco_titulo = ctk.CTkLabel(master=frame_dist_peso, text= "Risco", text_color='white', corner_radius=20)
    tipo_risco_box = ttk.Combobox(master= frame_dist_peso, values=["Definir (%)", "5","10","15","20","25","30","35","40"])
    tipo_risco_titulo.grid(padx=5,pady=5, row= 0, column=1)
    tipo_risco_box.grid(padx=10,pady=10, row=1, column=1)

distribuicao()

# -------------------------------------------- BLOCOS DE INDICADORES --------------------------------------------------------------

def bloco_indicadores():

    fonte_colunas = ctk.CTkFont(family='Arial', size=15, weight='bold')
    
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

class Botao:
    def botao_salvar_config(frame):
        botao_salvar = ctk.CTkButton(frame, text="Salvar", command=Botao.botao_salvar_event)
        botao_salvar.grid(pady=(20, 10), padx=20, sticky="w",row=10, column=0)  

    def botao_salvar_event():
            print("Botão salvar clicado")
            wb.save(file)
            messagebox.showinfo("Sucesso", "Alterações salvas com sucesso!", icon='info')

            

    def botao_visualizar_dashboard_config(frame):
        botao_visualizar = ctk.CTkButton(frame, text="Visualizar Dashboard", command=Botao.botao_visualizar_dashboard_event)
        botao_visualizar.grid(pady=(20, 10), padx=20, sticky="w",row=10,column=1)  

    def botao_visualizar_dashboard_event():
        print("Visualizar Dashboard clicado")



def main():
    caminho = xlsx.xlsx_state(self=xlsx)
    print(caminho)

    # Adicionando o bloco de indicadores
    bloco_indicadores()

    # Adicionando os botões
    Botao.botao_salvar_config(frame)
    Botao.botao_visualizar_dashboard_config(frame)

    MainWindow.window_config(window)

if __name__ == "__main__":
    main()
    window.mainloop()

