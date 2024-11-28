import asyncio
from tkinter import messagebox
import customtkinter as ctk

# Simulação das funções de dashboard (substitua pelas reais)
def dashboard():
    print("Carregando dashboard...")
    time.sleep(3)  # Simula o tempo de carregamento

def show_filter():
    print("Exibindo filtros...")
    time.sleep(1)  # Simula o tempo de exibição

# Função para carregar dashboard em uma thread separada
def load_dashboard():
    dashboard()
    show_filter()

# Animação de carregamento
async def show_loading_text():
    for _ in range(10):  # Limita o tempo da animação
        window.title("Sistema de Matriz de Seletividade (Carregando.)")
        await asyncio.sleep(0.6)
        window.title("Sistema de Matriz de Seletividade (Carregando..)")
        await asyncio.sleep(0.6)
        window.title("Sistema de Matriz de Seletividade (Carregando...)")
        await asyncio.sleep(0.6)

def hide_loading_text():
    window.title('Sistema de Matriz de Seletividade')

# Função principal para carregar o dashboard com animação
async def loading_text_animation():
    try:
        loading_task = asyncio.create_task(show_loading_text())
        await asyncio.to_thread(load_dashboard)
    finally:
        loading_task.cancel()  # Cancela a animação quando o dashboard for carregado
        hide_loading_text()
        print("Dashboard carregado com sucesso!")

# Validação de distribuição (substitua pela lógica real)
def validar_distribuicao():
    return True  # Simulação: sempre válido

# Botão para visualizar o dashboard
fonte_botao = ctk.CTkFont("Arial", size=15, weight='bold')

def botao_visualizar_dashboard_config(frame_botoes):
    botao_visualizar = ctk.CTkButton(
        frame_botoes,
        text="Ranking",
        command=botao_visualizar_dashboard_event,
        font=fonte_botao,
        fg_color="#2F83D7"
    )
    botao_visualizar.grid(pady=(10, 10), padx=20, sticky="w", row=10, column=1)

def botao_visualizar_dashboard_event():
    if validar_distribuicao():
        # Agendar a animação e carregamento no loop Tkinter
        asyncio.create_task(loading_text_animation())
    else:
        messagebox.showerror("Erro", "Salve as alterações para visualizar o Ranking")

# Exemplo de janela Tkinter (adapte ao seu contexto)
window = ctk.CTk()  # Inicializa janela
window.geometry("800x600")
frame_botoes = ctk.CTkFrame(window)
frame_botoes.pack(pady=20)
botao_visualizar_dashboard_config(frame_botoes)
window.mainloop()
