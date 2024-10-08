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
        window.geometry("800x600")
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

frame = ctk.CTkScrollableFrame(master=window, border_width=0, corner_radius=20, bg_color="#2F83D7", fg_color="#2F83D7", height=600)
frame.pack(fill='both', padx=20, pady=10)

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
    coluna_risco.grid(padx=20,pady=5,row=0,column=0)

    coluna_relevancia = ctk.CTkLabel(master=frame, text= "RELEVÂNCIA", font=fonte_colunas, text_color='white', corner_radius=20)
    coluna_relevancia.grid(padx=20,pady=5,row=0,column=1)

    coluna_materialidade = ctk.CTkLabel(master=frame, text= "MATERIALIDADE", font=fonte_colunas, text_color='white', corner_radius=20)
    coluna_materialidade.grid(padx=20,pady=5,row=0,column=2)

    coluna_oportunidade = ctk.CTkLabel(master=frame, text= "OPORTUNIDADE", font=fonte_colunas, text_color='white', corner_radius=20)
    coluna_oportunidade.grid(padx=20,pady=5,row=0,column=3)


#TIPOS DE RISCO PRIMEIRA COLUNA

    # HISTÓRICO PARECER PRÉVIO (ÚLTIMOS 3 ANOS)
    indicador1_frame = ctk.CTkFrame(frame, fg_color="#FAFFFD", corner_radius=10)
    indicador1_frame.grid(padx=20, pady=10, sticky="nsew")

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
    indicador1_title.grid(pady=5, sticky="w")  # Alinhado à esquerda

    # Switch referente ao indicador
    switch1_var = ctk.StringVar(value="NÃO")
    switch1 = ctk.CTkSwitch(indicador1_frame, text=f"Habilitar? ({switch1_var.get()})", command=botao_switch1_event, variable=switch1_var, onvalue="SIM", offvalue="NÃO", text_color="black")
    botao_switch1_event()
    switch1.grid(padx=10, pady=5, sticky="w")  # Adicionando padx


    # DATA ÚLTIMA AUDITORIA (3DCE)
    indicador2_frame = ctk.CTkFrame(frame, fg_color="#FAFFFD", corner_radius=10)
    indicador2_frame.grid(padx=20, pady=10, sticky="nsew")

    def botao_switch2_event():
        print("Switch 2 toggled, current value:", switch2_var.get())
        switch2.configure(text=f"Habilitar? ({switch2_var.get()})")

        # Habilitar/Desabilitar indicador
        if switch2_var.get() == "SIM":
            sheet['F12'] = 'SIM'
        else:
            sheet["F12"] = 'NÃO'

    indicador2_title = ctk.CTkLabel(indicador2_frame, text="DATA ÚLTIMA AUDITORIA (3DCE)", text_color="black", corner_radius=20, anchor="w")
    indicador2_title.grid(pady=5, sticky="w")

    switch2_var = ctk.StringVar(value="NÃO")
    switch2 = ctk.CTkSwitch(indicador2_frame, text=f"Habilitar? ({switch2_var.get()})", command=botao_switch2_event, variable=switch2_var, onvalue="SIM", offvalue="NÃO", text_color="black")
    botao_switch2_event()
    switch2.grid(padx=10, pady=5, sticky="w") 

    # POSIÇÃO - QTDE DE DÉBITO/MULTAS
    indicador3_frame = ctk.CTkFrame(frame, fg_color="#FAFFFD", corner_radius=10)
    indicador3_frame.grid(padx=20, pady=10, sticky="nsew")

    def botao_switch3_event():
        print("Switch 3 toggled, current value:", switch3_var.get())
        switch3.configure(text=f"Habilitar? ({switch3_var.get()})")

        if switch3_var.get() == "SIM":
          sheet['F13'] = 'SIM'
        else:
          sheet['F13'] = 'NÃO'

    indicador3_title = ctk.CTkLabel(indicador3_frame, text="POSIÇÃO - QTDE DE DÉBITO/MULTAS", text_color="black", corner_radius=20, anchor="w")
    indicador3_title.grid(pady=5, sticky="w")

    switch3_var = ctk.StringVar(value="NÃO")
    switch3 = ctk.CTkSwitch(indicador3_frame, text=f"Habilitar? ({switch3_var.get()})", command=botao_switch3_event, variable=switch3_var, onvalue="SIM", offvalue="NÃO", text_color="black")
    botao_switch3_event()
    switch3.grid(padx=10, pady=5, sticky="w") 


#TIPOS DE RELEVÂNCIA SEGUNDA COLUNA

    # POSIÇÃO - POPULAÇÃO MUNICÍPIO
    indicador4_frame = ctk.CTkFrame(frame, fg_color="#FAFFFD", corner_radius=10)
    indicador4_frame.grid(padx=20, pady=10, sticky="nsew",column=1, row= 1)

    def botao_switch4_event():
        print("Switch 4 toggled, current value:", switch3_var.get())
        switch4.configure(text=f"Habilitar? ({switch3_var.get()})")

        if switch4_var.get() == "SIM":
          sheet['F17'] = 'SIM'
        else:
          sheet['F17'] = 'NÃO'

    indicador4_title = ctk.CTkLabel(indicador4_frame, text="POSIÇÃO - POPULAÇÃO MUNICÍPIO", text_color="black", corner_radius=20, anchor="w")
    indicador4_title.grid(padx=10,pady=5, sticky="w")

    switch4_var = ctk.StringVar(value="NÃO")
    switch4 = ctk.CTkSwitch(indicador4_frame, text=f"Habilitar? ({switch4_var.get()})", command=botao_switch4_event, variable=switch4_var, onvalue="SIM", offvalue="NÃO", text_color="black")
    botao_switch4_event()
    switch4.grid(padx=10, pady=5, sticky="w") 


# -------------------------------------------- FIM DO BLOCO DE INDICADORES ---------------------------------------------------------------

#-----------------------------------------------------------------------------------------------------------------------------------------

class Botao:
    def botao_salvar_config(frame):
        botao_salvar = ctk.CTkButton(frame, text="Salvar", command=Botao.botao_salvar_event)
        botao_salvar.grid(pady=(20, 10), padx=20, sticky="w",row=5, column=0)  # Adicionei row e column

    def botao_salvar_event():
            print("Botão salvar clicado")
            wb.save(file)
            messagebox.showinfo("Sucesso", "Alterações salvas com sucesso!", icon='info')

            

    def botao_visualizar_dashboard_config(frame):
        botao_visualizar = ctk.CTkButton(frame, text="Visualizar Dashboard", command=Botao.botao_visualizar_dashboard_event)
        botao_visualizar.grid(pady=(20, 10), padx=20, sticky="w",row=5,column=1)  # Adicionei row e column

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

