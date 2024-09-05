import win32gui, win32con
import time
import pyautogui
import tabela
import leitorDeTela, exportar, janelasSap, transacoes


pyautogui.PAUSE = 0.5


def rodar_mb25():
    hwnd = win32gui.FindWindow(None, "PS0(4)/011 SAP Easy Access")
    time.sleep(0.5)
    win32gui.ShowWindow(hwnd, win32con.SW_MAXIMIZE)
    win32gui.SetForegroundWindow(hwnd)
    time.sleep(0.5)
    pyautogui.click(x=170, y=75)
    pyautogui.write("mb25")
    pyautogui.press("enter")
    time.sleep(0.5)
    janelasSap.esperar_janela_abrir("PS0(4)/011 Lista de reservas administração de estoques")
    transacoes.puxar_variante("mb25")
    leitorDeTela.esperar_tela(["Denom."])
    exportar.exportar_plainha("mb25")
    time.sleep(1)
    tabela.alterar_tabela_mb25()
