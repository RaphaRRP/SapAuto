import tkinter as tk
from tkinter import messagebox
import tabela, janelasSap, mb25, transacoes
import time
import win32gui

processo_iniciado = False

def botao_1():

    if janelasSap.janela_aberta_com_excel():
        messagebox.showwarning("Aviso", "Por favor, feche suas tabelas do excel abertas para continuar o processo")
    else:
        janelasSap.janelasSap()
        time.sleep(1)
        mb25.rodar_mb25(usuario.get())
        time.sleep(1)
        transacoes.rodar_transacoes_sap(usuario.get())
        messagebox.showinfo("Sucesso", "Primeira parte do processo concluida, continue quando carregar a transação 'z22m051'")
        global processo_iniciado 
        processo_iniciado = True

def botao_2():
    if processo_iniciado:
        transacoes.exportar_051(usuario.get())
        time.sleep(1)
        tabela.alterar_tabelas()
        messagebox.showinfo("Sucesso", "Prpcesso finalizado!")

    else:
        messagebox.showwarning("Aviso", "Comece o processo antes de continuar")

janela = tk.Tk()
janela.title("Troca de Botões")
janela.geometry("300x300")
janela.geometry("+760+390")


botao_1 = tk.Button(janela, text="Começar processo", command=botao_1)
botao_2 = tk.Button(janela, text="Continuar processo", command=botao_2)




usuario = tk.StringVar(value="Isa")

# Criando os botões de opção (Radiobuttons)
radio1 = tk.Radiobutton(janela, text="Isa", variable=usuario, value="Isa")
radio2 = tk.Radiobutton(janela, text="Rapha", variable=usuario, value="Rapha")

radio1.pack(anchor=tk.W)
radio2.pack(anchor=tk.W)

botao_1.pack(pady=20)
botao_2.pack(pady=15)



janela.mainloop()

