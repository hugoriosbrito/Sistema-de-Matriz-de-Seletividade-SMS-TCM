import customtkinter as ctk
from idlelib.tooltip import Hovertip


class Indicador:
    FG_COLOR = "#FAFFFD"
    CORNER_RADIUS = 10
    row_counter = {0: 0, 1: 0, 2: 0, 3: 0}

    def __init__(self, tipo, nome, celula_xlsx,descricao, sheet):
        self.tipo = tipo
        self.nome = nome
        self.celula_xlsx = celula_xlsx
        self.descricao = descricao
        self.sheet = sheet
        self.switch_var = ctk.StringVar(value="NÃO")

    def criarFrame(self, master):
        column = self._determine_column()
        row = self.row_counter[column]

        indicador_frame = ctk.CTkFrame(
            master=master, fg_color=self.FG_COLOR, corner_radius=self.CORNER_RADIUS
        )

        # Adiciona o frame à grid na coluna e linha corretas
        indicador_frame.grid(row=row, column=column, padx=20, pady=10, sticky="nsew")

        self._create_title(indicador_frame)
        self._create_description(indicador_frame)
        self._create_switch(indicador_frame)

        # Atualiza o contador de linha para a coluna
        self.row_counter[column] += 1

    def _determine_column(self):
        return {
            'risco': 0,
            'relevancia': 1,
            'materialidade': 2,
            'oportunidade': 3
        }.get(self.tipo)

    def _create_description(self, parent):
        tooltipWidget = ctk.CTkLabel(parent, text='[?]', text_color="#bababa", anchor="w")
        tooltipWidget.grid(pady=5,padx=2, row=0, sticky='e')
        tooltip = Hovertip(tooltipWidget, text=self.descricao, background='white', hover_delay=0,foreground='black')

    def _create_title(self, parent):
        title_label = ctk.CTkLabel(
            parent, text=self.nome, text_color="black", corner_radius=20, anchor="w"
        )
        title_label.grid(pady=5, sticky="w")

    def _create_switch(self, parent):
        self.switch = ctk.CTkSwitch(
            parent, text=f"Habilitar? ({self.switch_var.get()})",
            command=self.botao_switch_event, variable=self.switch_var,
            onvalue="SIM", offvalue="NÃO", text_color="black"
        )
        self.switch.grid(padx=10, pady=5, sticky="w")

    def botao_switch_event(self):
        print(f"Switch {self.nome} toggled, current value:{self.switch_var.get()}")
        self.switch.configure(text=f"Habilitar? ({self.switch_var.get()})")
        self.sheet[self.celula_xlsx] = 'SIM' if self.switch_var.get() == "SIM" else 'NÃO'
