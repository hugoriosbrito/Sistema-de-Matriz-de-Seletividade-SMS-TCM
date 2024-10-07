import openpyxl
import xlwings as xw
import customtkinter as ctk
from tkinter import messagebox
import os

file = "dados\\Matriz modelo - REV3 - Otimizado.xlsx"
app = xw.App(visible=False)
wb = app.books.open(file, read_only=False)
sheet = wb.sheets['SÍNTESE MUN.']
wb.autosave = True
wb.autosave_interval = 1


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

def textos():
    fonte_titulo = ctk.CTkFont(family='Arial', size=50, weight='bold')
    titulo = ctk.CTkLabel(window, text="Indicadores", font=fonte_titulo, anchor="center", corner_radius=20, text_color="white")
    titulo.pack(pady=20, padx=20, anchor="center")  # Pequeno espaçamento nas laterais do título

textos()

frame = ctk.CTkScrollableFrame(master=window, border_width=0, corner_radius=2, bg_color="#2F83D7", fg_color="#2F83D7", height=600)
frame.pack(fill='both', padx=20, pady=20)

# -------------------------------------------- BLOCOS DE INDICADORES --------------------------------------------------------------

def bloco_indicadores():
    tipos = [
    'RISCO',
    'MATERIALIDADE',
    'RELEVÂNCIA',
    'OPORTUNIDADE']

    # HISTÓRICO PARECER PRÉVIO (ÚLTIMOS 3 ANOS)
    indicador1_frame = ctk.CTkFrame(frame, fg_color="#FAFFFD", corner_radius=10)
    indicador1_frame.grid(padx=20, pady=10, sticky="nsew")

    # Variáveis globais para os widgets
    global indicador1_text, respostaIndicador1
    indicador1_text = ctk.CTkLabel(indicador1_frame, text=f"Quanto maior HISTÓRICO PARECER PRÉVIO (ÚLTIMOS 3 ANOS) maior {tipos[0]} ?", text_color="black", corner_radius=20, anchor="w")
    respostaIndicador1 = ctk.CTkComboBox(indicador1_frame, values=["Sim", "Não"], command=lambda resp: print(f"Valor selecionado no ComboBox (Histórico): {resp}"))

    # Inicializa a visibilidade dos widgets como "não visíveis"
    indicador1_text.grid_forget()
    respostaIndicador1.grid_forget()

    # Botão Switch event
    def botao_switch1_event():
        print("Switch 1 toggled, current value:", switch1_var.get())
        switch1.configure(text=f"Habilitar? ({switch1_var.get()})")

        if  switch1_var.get() == 'SIM':
            sheet.range('G11').value = 'SIM'
        else:
            sheet.range('G11').value = 'NÃO'

        # Altera a visibilidade dos widgets
        if switch1_var.get() == "SIM":
            indicador1_text.grid(pady=5, sticky="w")
            respostaIndicador1.grid(padx=10, pady=5, sticky="w")

            if respostaIndicador1.get() == "Sim":
                sheet.range('J11').value = 'CRESCENTE'
            else:
                sheet.range('J11').value = 'DECRESCENTE'

        else:
            indicador1_text.grid_forget()
            respostaIndicador1.grid_forget()
            
    # Título indicador
    indicador1_title = ctk.CTkLabel(indicador1_frame, text="HISTÓRICO PARECER PRÉVIO (ÚLTIMOS 3 ANOS)", text_color="black", corner_radius=20, anchor="w")
    indicador1_title.grid(pady=5, sticky="w")  # Alinhado à esquerda

    # Switch referente ao indicador
    switch1_var = ctk.StringVar(value="NÃO")
    switch1 = ctk.CTkSwitch(indicador1_frame, text=f"Habilitar? ({switch1_var.get()})", command=botao_switch1_event, variable=switch1_var, onvalue="SIM", offvalue="NÃO", text_color="black")
    switch1.grid(padx=10, pady=5, sticky="w")  # Adicionando padx


    # DATA ÚLTIMA AUDITORIA (3DCE)
    indicador2_frame = ctk.CTkFrame(frame, fg_color="#FAFFFD", corner_radius=10)
    indicador2_frame.grid(padx=20, pady=10, sticky="nsew")

    # Variáveis globais para os widgets do segundo indicador
    global indicador2_text, respostaIndicador2
    indicador2_text = ctk.CTkLabel(indicador2_frame, text="Quanto maior DATA ÚLTIMA AUDITORIA (3DCE) maior (tipo)?", text_color="black", corner_radius=20, anchor="w")
    respostaIndicador2 = ctk.CTkComboBox(indicador2_frame, values=["Sim", "Não"], command=lambda resp: print(f"Valor selecionado no ComboBox (Auditoria): {resp}"))

    # Inicializa a visibilidade dos widgets como "não visíveis"
    indicador2_text.grid_forget()
    respostaIndicador2.grid_forget()

    def botao_switch2_event():
        print("Switch 2 toggled, current value:", switch2_var.get())
        switch2.configure(text=f"Habilitar? ({switch2_var.get()})")

        # Altera a visibilidade dos widgets
        if switch2_var.get() == "SIM":
            indicador2_text.grid(pady=5, sticky="w")
            respostaIndicador2.grid(padx=10, pady=5, sticky="w")
        else:
            indicador2_text.grid_forget()
            respostaIndicador2.grid_forget()

    indicador2_title = ctk.CTkLabel(indicador2_frame, text="DATA ÚLTIMA AUDITORIA (3DCE)", text_color="black", corner_radius=20, anchor="w")
    indicador2_title.grid(pady=5, sticky="w")

    switch2_var = ctk.StringVar(value="NÃO")
    switch2 = ctk.CTkSwitch(indicador2_frame, text=f"Habilitar? ({switch2_var.get()})", command=botao_switch2_event, variable=switch2_var, onvalue="SIM", offvalue="NÃO", text_color="black")
    switch2.grid(padx=10, pady=5, sticky="w")  # Adicionando padx

    # POSIÇÃO - QTDE DE DÉBITO/MULTAS
    indicador3_frame = ctk.CTkFrame(frame, fg_color="#FAFFFD", corner_radius=10)
    indicador3_frame.grid(padx=20, pady=10, sticky="nsew")

    # Variáveis globais para os widgets do terceiro indicador
    global indicador3_text, respostaIndicador3
    indicador3_text = ctk.CTkLabel(indicador3_frame, text="Quanto maior POSIÇÃO - QTDE DE DÉBITO/MULTAS maior (tipo)?", text_color="black", corner_radius=20, anchor="w")
    respostaIndicador3 = ctk.CTkComboBox(indicador3_frame, values=[ "Sim", "Não"], command=lambda resp: print(f"Valor selecionado no ComboBox (Débito/Multas): {resp}"))

    # Inicializa a visibilidade dos widgets como "não visíveis"
    indicador3_text.grid_forget()
    respostaIndicador3.grid_forget()

    def botao_switch3_event():
        print("Switch 3 toggled, current value:", switch3_var.get())
        switch3.configure(text=f"Habilitar? ({switch3_var.get()})")

        if switch3_var.get() == "SIM":
            indicador3_text.grid(pady=5, sticky="w")
            respostaIndicador3.grid(padx=10, pady=5, sticky="w")
        else:
            indicador3_text.grid_forget()
            respostaIndicador3.grid_forget()

    indicador3_title = ctk.CTkLabel(indicador3_frame, text="POSIÇÃO - QTDE DE DÉBITO/MULTAS", text_color="black", corner_radius=20, anchor="w")
    indicador3_title.grid(pady=5, sticky="w")

    switch3_var = ctk.StringVar(value="NÃO")
    switch3 = ctk.CTkSwitch(indicador3_frame, text=f"Habilitar? ({switch3_var.get()})", command=botao_switch3_event, variable=switch3_var, onvalue="SIM", offvalue="NÃO", text_color="black")
    switch3.grid(padx=10, pady=5, sticky="w")  # Adicionando padx


# -------------------------------------------- FIM DO BLOCO DE INDICADORES ---------------------------------------------------------------

#lista para verificações
#respostasComboBox = [respostaIndicador1,respostaIndicador2,respostaIndicador3]

#-----------------------------------------------------------------------------------------------------------------------------------------

class Botao:
    def botao_salvar_config(frame):
        botao_salvar = ctk.CTkButton(frame, text="Salvar", command=Botao.botao_salvar_event)
        botao_salvar.grid(pady=(20, 10), padx=20, sticky="w",row=5, column=0)  # Adicionei row e column

    def botao_salvar_event():
            print("Botão salvar clicado")
            wb.save(file)
            messagebox.showinfo("Sucesso", "Alterações salvas com sucesso!", icon='info')
            wb.close()

    def botao_visualizar_dashboard_config(frame):
        botao_visualizar = ctk.CTkButton(frame, text="Visualizar Dashboard", command=Botao.botao_visualizar_dashboard_event)
        botao_visualizar.grid(pady=(20, 10), padx=20, sticky="w",row=5,column=1)  # Adicionei row e column

    def botao_visualizar_dashboard_event():
        print("Visualizar Dashboard clicado")



def main():
    caminho = xlsx.xlsx_state(self=xlsx)
    print(caminho)

    #inicialização padrão das células
    sheet['G11:G30'].value = 'NÃO'

    # Adicionando o bloco de indicadores
    bloco_indicadores()

    # Adicionando os botões
    Botao.botao_salvar_config(frame)
    Botao.botao_visualizar_dashboard_config(frame)

    MainWindow.window_config(window)

if __name__ == "__main__":
    main()
    window.mainloop()

