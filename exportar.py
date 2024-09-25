import pyautogui
import leitorDeTela
from datetime import datetime
import win32gui, win32con
import time 
import janelasSap

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


def exportar_plainha(transacao, usuario):
    data_hoje = datetime.now().strftime("%d-%m")
    if usuario == "Rapha":
        caminho = f"C:/Users/prr8ca/Desktop/Projetos/SapAuto/tabelas/{transacao}"
    elif usuario == "Isa":
        caminho = f"C:/Users/mis7ca/Desktop/Projetos/SapAuto/tabelas/{transacao}"
        
    arquivo = f"{transacao} - {data_hoje}.xlsx"


    if transacao == "mb25":
        cordenada_clique_1 = 77, 27
        cordenada_clique_2 = 178, 98
        cordenada_clique_3 = 554, 136

        if usuario == "Rapha":
            cordenada_clique_4 = 621, 744
            cordenada_clique_5 = 759, 616
            cordenada_clique_6 = 894, 848
        
        elif usuario == "Isa":
            cordenada_clique_4 = 617, 685
            cordenada_clique_5 = 759, 684
            cordenada_clique_6 = 893, 847


    elif transacao == "me3m":
        cordenada_clique_1 = 78, 32
        cordenada_clique_2 = 209, 134
        cordenada_clique_3 = 615, 167

        if usuario == "Rapha":
            cordenada_clique_4 = 620, 743
            cordenada_clique_5 = 759, 616
            cordenada_clique_6 = 893, 851
                    
        elif usuario == "Isa":
            cordenada_clique_4 = 622, 683
            cordenada_clique_5 = 759, 684
            cordenada_clique_6 = 893, 847


    elif transacao == "z22k032":
        cordenada_clique_1 = 77, 27
        cordenada_clique_2 = 178, 98
        cordenada_clique_3 = 554, 136

        if usuario == "Rapha":
            cordenada_clique_4 = 617, 769
            cordenada_clique_5 = 757, 618
            cordenada_clique_6 = 896, 849
                    
        elif usuario == "Isa":
            cordenada_clique_4 = 621, 832
            # cordenada_clique_5 = 759, 682
            # cordenada_clique_6 = 894, 848


    elif transacao == "z22m051":
        cordenada_clique_1 = 77, 32
        cordenada_clique_2 = 237, 165
        cordenada_clique_3 = 726, 202

        if usuario == "Rapha":
            cordenada_clique_4 = 620, 768
            cordenada_clique_5 = 758, 618
            cordenada_clique_6 = 894, 850
                    
        elif usuario == "Isa":
            cordenada_clique_4 = 609, 682
            # cordenada_clique_5 = 
            # cordenada_clique_6 = 

    time.sleep(5)
    pyautogui.click(cordenada_clique_1)
    time.sleep(0.2)
    pyautogui.click(cordenada_clique_2)
    time.sleep(0.2)
    pyautogui.click(cordenada_clique_3)
    

    leitorDeTela.esperar_tela(["Salvar como"])

    if usuario == "Rapha":
        pyautogui.click(x=239, y=77) #meu
    elif usuario == "Isa":
        pyautogui.click(x=180, y=67) #isa
    pyautogui.write(caminho)
    pyautogui.press("enter")

    if usuario == "Rapha":
        pyautogui.click(x=1606, y=840) #meu
    elif usuario == "Isa":
        pyautogui.click(x=1716, y=920) #isa 
    pyautogui.write(arquivo)
    pyautogui.press("enter")

    while janelasSap.janela_aberta_com_excel() != True:
        print("espeando excel abrir")
        time.sleep(1)
    
    time.sleep(3)
    fechar_janelas("Excel")



    # elif usuario == "Isa":
    #     if transacao == "mb25" or transacao == "me3m":
    #         pyautogui.click(cordenada_clique_5)
    #         time.sleep(2)
    #         pyautogui.click(cordenada_clique_6)

    #     elif (transacao == "z22k032" or transacao == "z22m051"):
    #         leitorDeTela.esperar_tela(["Salvo", "Exibir"])
    #         fechar_janelas("Excel")
