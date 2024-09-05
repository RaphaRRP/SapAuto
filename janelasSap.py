import win32com.client, win32gui, win32con
import subprocess
import time
import pyautogui
import exportar


def mensagem_de_erro(funcao, erro):
    print(f"Função: {funcao}. Erro: {erro}")


def maximizar_janela(window_title):
    try:
        hwnd = win32gui.FindWindow(None, window_title)
        win32gui.ShowWindow(hwnd, win32con.SW_MAXIMIZE)
        win32gui.SetForegroundWindow(hwnd)
        print("Janela maximizada com sucesso.")
    except Exception as erro:
        mensagem_de_erro("maximizar_janela", erro)


def esperar_janela_abrir(window_title):
    try:
        hwnd = win32gui.FindWindow(None, window_title)
        while not hwnd:
            hwnd = win32gui.FindWindow(None, window_title)
            time.sleep(0.5)
            print(f"esperando {window_title}")
        print(f"{window_title} abriu")

    except Exception as erro:
        mensagem_de_erro("esperar_janela_abrir", erro)

# Função para esperar uma janela abrir, que pode não abrir, se acontecer, tenta abrir denovo
def esperar_janela_abrir_sap(window_title):
    contador = 0
    try:
        hwnd = win32gui.FindWindow(None, window_title)
        while not hwnd:
            hwnd = win32gui.FindWindow(None, window_title)
            time.sleep(0.5)
            print(f"esperando {window_title}")
            contador += 1
            if contador >= 20:
                print("Janela 4 não abriu de primeira")
                return True

        print(f"{window_title} abriu")
        return False

    except Exception as erro:
        mensagem_de_erro("esperar_janela_abrir", erro)


def janela_esta_aberta(window_title):
    #se fechado retorna 0, se aberto retorna um numero =! 0(hwnd)
    try:
        hwnd = win32gui.FindWindow(None, window_title)
        return hwnd
    
    except Exception as erro:
        mensagem_de_erro("janela_esta_aberta", erro)


def abrir_sap():
    try:
        if janela_esta_aberta("SAP Logon 800") == 0:
            path = r"C:\Program Files (x86)\SAP\FrontEnd\SAPGUI\saplogon.exe"
            subprocess.Popen(path)
            esperar_janela_abrir("SAP Logon 800")
        
    except Exception as erro:
        mensagem_de_erro("abrir_sap", erro)


def abrir_ps0():
    try:
        exportar.fechar_janelas("PS0")

        SapGuiAuto = win32com.client.GetObject("SAPGUI")
        application = SapGuiAuto.GetScriptingEngine
        connection = application.OpenConnection("PS0: Produktion Brazil 2BIP", True)
            

        esperar_janela_abrir("PS0(1)/011 SAP Easy Access")
        maximizar_janela("PS0(1)/011 SAP Easy Access")
    except Exception as erro:
        mensagem_de_erro("abrir_ps0", erro)


def abrir_todas_janelas():
    time.sleep(1)
    pyautogui.click(x=755, y=75)
    time.sleep(0.3)
    pyautogui.click(x=755, y=75)
    time.sleep(0.3)
    pyautogui.click(x=755, y=75)

    if esperar_janela_abrir_sap("PS0(4)/011 SAP Easy Access"):
        pyautogui.click(x=755, y=75)
    time.sleep(3)


    #=====Main======#
def janelasSap():
    abrir_sap()
    abrir_ps0()
    abrir_todas_janelas()
    