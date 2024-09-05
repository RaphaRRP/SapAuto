import pyautogui
import leitorDeTela
from datetime import datetime
import win32gui, win32con
import time 

pyautogui.PAUSE = 0.5


def fechar_janelas(janela):

    def enum_windows_callback(hwnd, window_list):
        if win32gui.IsWindowVisible(hwnd):
            janela_titulo = win32gui.GetWindowText(hwnd)
            if janela in janela_titulo:
                window_list.append(hwnd)

    windows_to_close = []
    win32gui.EnumWindows(enum_windows_callback, windows_to_close)
    
    for hwnd in windows_to_close:
        win32gui.PostMessage(hwnd, win32con.WM_CLOSE, 0, 0)
    
    print(F"Todas as janelas {janela} foram fechadas.")


def exportar_plainha(transacao):
    data_hoje = datetime.now().strftime("%d-%m")
    caminho = f"C:/Users/prr8ca/Desktop/Projetos/SapAuto/tabelas/{transacao}"
    arquivo = f"{transacao} - {data_hoje}.xlsx"

    if transacao == "mb25" or transacao == "z22k032":
        cordenada_clique_1 = 77, 27
        cordenada_clique_2 = 178, 98
        cordenada_clique_3 = 554, 136

    elif transacao == "me3m":
        cordenada_clique_1 = 78, 32
        cordenada_clique_2 = 209, 134
        cordenada_clique_3 = 615, 167

    elif transacao == "z22m051":
        cordenada_clique_1 = 77, 32
        cordenada_clique_2 = 237, 165
        cordenada_clique_3 = 726, 202

    time.sleep(5)
    pyautogui.click(cordenada_clique_1)
    time.sleep(0.2)
    pyautogui.click(cordenada_clique_2)
    time.sleep(0.2)
    pyautogui.click(cordenada_clique_3)
    
    #  Isso é temporário
    #pyautogui.hotkey('ctrl', 'alt', 'f7', interval=0.1)

    leitorDeTela.esperar_tela(["Salvar como"])

    pyautogui.click(x=240, y=80)
    pyautogui.write(caminho)
    pyautogui.press("enter")

    pyautogui.click(x=1780, y=840)
    pyautogui.write(arquivo)
    pyautogui.press("enter")

    leitorDeTela.esperar_tela(["Salvo", "Exibir"])
    fechar_janelas("Excel")
