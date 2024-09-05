import win32gui, win32con
import time
import pyautogui
import leitorDeTela, exportar

pyautogui.PAUSE = 0.5

def puxar_variante(transacao):
    

    time.sleep(0.5)
    pyautogui.click(73, 181)
    pyautogui.click(673, 654)
    pyautogui.hotkey('ctrl', 'a')
    pyautogui.write("mis7ca")
    pyautogui.press("enter")
    if transacao == "mb25":
        pyautogui.click(x=632, y=267)
        pyautogui.hotkey('ctrl', 'a')
        pyautogui.press('backspace')
        pyautogui.click(40, 180)
    
    time.sleep(0.5)


def rodar_transacoes_sap():
    
    time.sleep(2)
    transacoes = ["z22m051", "z22k032", "me3m"]
    for transacao in transacoes:

#==========================Procurar Transação================================#

        if transacao == "z22m051":
            hwnd = win32gui.FindWindow(None, "PS0(3)/011 SAP Easy Access")
            

        elif transacao == "z22k032":
            hwnd = win32gui.FindWindow(None, "PS0(2)/011 SAP Easy Access")
            
    
        elif transacao == "me3m":
            hwnd = win32gui.FindWindow(None, "PS0(1)/011 SAP Easy Access")
            

        time.sleep(0.5)
        win32gui.ShowWindow(hwnd, win32con.SW_MAXIMIZE)
        win32gui.SetForegroundWindow(hwnd)
        time.sleep(0.5)
        pyautogui.click(x=170, y=75)
        pyautogui.write(transacao)
        pyautogui.press("enter")
        time.sleep(0.5)

#==========================Puxar Variantes================================#

        if transacao == "z22m051":
            puxar_variante("z22m051")

        elif transacao == "z22k032":
            puxar_variante("z22k032")
            pyautogui.click(x=1050, y=370)
            pyautogui.hotkey('ctrl', 'a')
            pyautogui.press('backspace')
            pyautogui.click(x=717, y=782)
            time.sleep(0.5)

        elif transacao == "me3m":
            time.sleep(1)
            pyautogui.click(x=536, y=371)
            pyautogui.hotkey('ctrl', 'a')
            pyautogui.write("ALLES_ALV")


#==========================Editar Materiais================================#

        if transacao == "z22m051":
            cordenada_clique_material = 1063, 268
            
        elif transacao == "z22k032":
            cordenada_clique_material = 1049, 439
            
        elif transacao == "me3m":
            cordenada_clique_material = 1062, 231

        pyautogui.click(cordenada_clique_material)
        time.sleep(0.5)
        pyautogui.click(x=888, y=782)
        time.sleep(0.5)
        pyautogui.click(x=1186, y=778)
        time.sleep(2)
        pyautogui.click(x=718, y=781)
        time.sleep(0.5)
        pyautogui.click(x=40, y=182)
        time.sleep(0.5)

#==========================exportar planilhas================================#

    time.sleep(2)
    transacoes = ["me3m", "z22k032"]
    for transacao in transacoes:
    
        if transacao == "z22k032":
            hwnd = win32gui.FindWindow(None, "PS0(2)/011 Estoque: Valorização de Material por Depósito (IM)")
            print("hwnd 2: ", hwnd)
            palavras_chave = ["Denom."]

        elif transacao == "me3m":
            hwnd = win32gui.FindWindow(None, "PS0(1)/011 Documentos de compras para material")
            print("hwnd 1: ", hwnd)
            palavras_chave = ["Nome do fornecedor"]

        time.sleep(0.5)
        win32gui.ShowWindow(hwnd, win32con.SW_MAXIMIZE)
        print("hwnd: ", hwnd)
        win32gui.SetForegroundWindow(hwnd)
        time.sleep(0.5)
        leitorDeTela.esperar_tela(palavras_chave)
        exportar.exportar_plainha(transacao)



def exportar_051():
    hwnd = win32gui.FindWindow(None, "PS0(3)/011 Lista Programação de Compras e Recebimentos Efetuados")
    print("hwnd 3: ", hwnd)
    palavras_chave = ["NBM"]
    time.sleep(0.5)
    win32gui.ShowWindow(hwnd, win32con.SW_MAXIMIZE)
    print("hwnd: ", hwnd)
    win32gui.SetForegroundWindow(hwnd)
    time.sleep(0.5)
    leitorDeTela.esperar_tela(palavras_chave)
    exportar.exportar_plainha("z22m051")

