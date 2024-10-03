import openpyxl
import customtkinter as ctk

# wb = openpyxl.load_workbook("pythonGUI_TCM\Matriz modelo - REV3.xlsx")
# sheet = wb['INDI GERAIS - MUNICÍPIO']

window = ctk.CTk()

class MainWindow:
    @staticmethod
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
    # HISTÓRICO PARECER PRÉVIO (ÚLTIMOS 3 ANOS)
    indicador1_frame = ctk.CTkFrame(frame, fg_color="#FAFFFD", corner_radius=10)
    indicador1_frame.grid(padx=20, pady=10, sticky="nsew")

    # Botão Switch event
    def botao_switch1_event():
        print("Switch 1 toggled, current value:", switch1_var.get())
        switch1.configure(text=f"Habilitar? ({switch1_var.get()})")
        if switch1_var.get() == "SIM":
                # Texto indicador
            indicador1_text = ctk.CTkLabel(indicador1_frame, text="Quanto maior HISTÓRICO PARECER PRÉVIO (ÚLTIMOS 3 ANOS) maior (tipo)?", text_color="black", corner_radius=20, anchor="w")
            indicador1_text.grid(pady=5, sticky="w")

            # ComboBox resposta
            def respostaIndicador1Value(resp):
                print(f"Valor selecionado no ComboBox (Histórico): {resp}")

            respostaIndicador1 = ctk.CTkComboBox(indicador1_frame, values=["...", "Sim", "Não"], command=respostaIndicador1Value)
            respostaIndicador1.grid(padx=10, pady=5, sticky="w")  # Adicionando padx


    # Título indicador
    indicador1_title = ctk.CTkLabel(indicador1_frame, text="HISTÓRICO PARECER PRÉVIO (ÚLTIMOS 3 ANOS)", text_color="black", corner_radius=20, anchor="w")
    indicador1_title.grid(pady=5, sticky="w")  # Alinhado à esquerda

    # Switch referente ao indicador
    switch1_var = ctk.StringVar(value="SIM")
    switch1 = ctk.CTkSwitch(indicador1_frame, text=f"Habilitar? ({switch1_var.get()})", command=botao_switch1_event, variable=switch1_var, onvalue="SIM", offvalue="NÃO", text_color="black")
    switch1.grid(padx=10, pady=5, sticky="w")  # Adicionando padx


    # DATA ÚLTIMA AUDITORIA (3DCE)
    indicador2_frame = ctk.CTkFrame(frame, fg_color="#FAFFFD", corner_radius=10)
    indicador2_frame.grid(padx=20, pady=10, sticky="nsew")

    def botao_switch2_event():
        print("Switch 2 toggled, current value:", switch2_var.get())
        switch2.configure(text=f"Habilitar? ({switch2_var.get()})")

    indicador2_title = ctk.CTkLabel(indicador2_frame, text="DATA ÚLTIMA AUDITORIA (3DCE)", text_color="black", corner_radius=20, anchor="w")
    indicador2_title.grid(pady=5, sticky="w")

    switch2_var = ctk.StringVar(value="SIM")
    switch2 = ctk.CTkSwitch(indicador2_frame, text=f"Habilitar? ({switch2_var.get()})", command=botao_switch2_event, variable=switch2_var, onvalue="SIM", offvalue="NÃO", text_color="black")
    switch2.grid(padx=10, pady=5, sticky="w")  # Adicionando padx

    indicador2_text = ctk.CTkLabel(indicador2_frame, text="Quanto maior DATA ÚLTIMA AUDITORIA (3DCE) maior (tipo)?", text_color="black", corner_radius=20, anchor="w")
    indicador2_text.grid(pady=5, sticky="w")

    def respostaIndicador2Value(resp):
        print(f"Valor selecionado no ComboBox (Auditoria): {resp}")

    respostaIndicador2 = ctk.CTkComboBox(indicador2_frame, values=["...", "Sim", "Não"], command=respostaIndicador2Value)
    respostaIndicador2.grid(padx=10, pady=5, sticky="w")  # Adicionando padx


    # POSIÇÃO - QTDE DE DÉBITO/MULTAS
    indicador3_frame = ctk.CTkFrame(frame, fg_color="#FAFFFD", corner_radius=10)
    indicador3_frame.grid(padx=20, pady=10, sticky="nsew")

    def botao_switch3_event():
        print("Switch 3 toggled, current value:", switch3_var.get())
        switch3.configure(text=f"Habilitar? ({switch3_var.get()})")

    indicador3_title = ctk.CTkLabel(indicador3_frame, text="POSIÇÃO - QTDE DE DÉBITO/MULTAS", text_color="black", corner_radius=20, anchor="w")
    indicador3_title.grid(pady=5, sticky="w")

    switch3_var = ctk.StringVar(value="SIM")
    switch3 = ctk.CTkSwitch(indicador3_frame, text=f"Habilitar? ({switch3_var.get()})", command=botao_switch3_event, variable=switch3_var, onvalue="SIM", offvalue="NÃO", text_color="black")
    switch3.grid(padx=10, pady=5, sticky="w")  # Adicionando padx

    indicador3_text = ctk.CTkLabel(indicador3_frame, text="Quanto maior POSIÇÃO - QTDE DE DÉBITO/MULTAS maior (tipo)?", text_color="black", corner_radius=20, anchor="w")
    indicador3_text.grid(pady=5, sticky="w")

    def respostaIndicador3Value(resp):
        print(f"Valor selecionado no ComboBox (Débito/Multas): {resp}")

    respostaIndicador3 = ctk.CTkComboBox(indicador3_frame, values=["...", "Sim", "Não"], command=respostaIndicador3Value)
    respostaIndicador3.grid(padx=10, pady=5, sticky="w")  # Adicionando padx


# -------------------------------------------- FIM DO BLOCO DE INDICADORES ---------------------------------------------------------------


class Botao:

    @staticmethod
    def botao_salvar_config():
        botao_salvar = ctk.CTkButton(frame, text="Salvar", command=Botao.botao_salvar_event)
        botao_salvar.grid(pady=(20, 10), padx=20, sticky="w")

    @staticmethod
    def botao_salvar_event():
        print("Botão salvar clicado")

    @staticmethod
    def botao_visualizar_dashboard_config():
        botao_visualizar = ctk.CTkButton(frame, text="Visualizar Dashboard", command=Botao.botao_visualizar_dashboard_event)
        botao_visualizar.grid(pady=10, padx=20, sticky="w")

    @staticmethod
    def botao_visualizar_dashboard_event():
        print("Botão visualizar dashboard clicado")


def widgets_init():
    bloco_indicadores()
    Botao.botao_salvar_config()
    Botao.botao_visualizar_dashboard_config()


def main():
    MainWindow.window_config(window)
    widgets_init()
    window.mainloop()


if __name__ == "__main__":
    main()
