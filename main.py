import openpyxl
import customtkinter as ctk

#wb = openpyxl.load_workbook("pythonGUI_TCM\Matriz modelo - REV3.xlsx")
#sheet = wb['INDI GERAIS - MUNICÍPIO']


window = ctk.CTk()

class mainWindow:
    def windowConfig(window):
        window.geometry("800x600")
        window._set_appearance_mode("dark")#Modo escuro da interface
        ctk.set_default_color_theme("dark-blue")
        window.grid_columnconfigure(0, weight=0) #Centralizar

class botao:
    def widgetsConfig():
        #botao para salvar arquivo com alterações
        botaoSalvar = ctk.CTkButton(window, text="Salvar", command=botao.botaoSalvar)
        botaoSalvar.grid(row=0, column=0, padx=20,pady=20)

        #botao Switch on/off referente ao indicador HISTÓRICO PARECER PRÉVIO (ÚLTIMOS 3 ANOS)
        switch1_var = ctk.StringVar(value="on")
        switch1 = ctk.CTkSwitch(window, text="Habilitar?", command=botao.switch1_event,
                                        variable=switch1_var, onvalue="on", offvalue="off")
        switch1.grid(row=1,column=0,padx = 20, pady = 20)

    def botaoSalvarAction():
        print("botao salvar clicado")

    def switch1_event():
      
      print("switch toggled, current value:", switch1_var.get())

def main():
    mainWindow.windowConfig(window)
    botao.widgetsConfig()
    window.mainloop()


main()