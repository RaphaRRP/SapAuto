import tkinter as tk
from tkinter import messagebox
import main

def botao_1():
    botao_1.pack_forget()  # Esconde o botão 1
    main.main()
    messagebox.showwarning("Aviso", "assim que a transação z22m051 carregar, aperte o botão continuar")
    botao_2.pack(pady=20)  # Mostra o botão 2

def botao_2():
    main.submain()
    janela.destroy()

# Criar a janela principal
janela = tk.Tk()
janela.title("Troca de Botões")
janela.geometry("300x200")

# Criar os botões
botao_1 = tk.Button(janela, text="Iniciar", command=botao_1)
botao_2 = tk.Button(janela, text="Continuar", command=botao_2)


# Mostrar inicialmente o Botão 1
botao_1.pack(pady=20)

janela.mainloop()
