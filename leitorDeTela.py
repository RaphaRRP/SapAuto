from PIL import Image
import pytesseract
import pyautogui
import time

# Caminho para o executável do Tesseract
pytesseract.pytesseract.tesseract_cmd = r'S:/PM/ter/tef/tef3/Inter_Setor/TEF3 Controles/Bloco K Planejamento/Automações/Tesseract-OCR/tesseract.exe'

def verificar_palavras_na_tela(palavras):
    # Captura uma captura de tela
    screenshot = pyautogui.screenshot()
    # Extrai texto da captura de tela usando o Tesseract
    texto_extraido = pytesseract.image_to_string(screenshot, lang="por")
    
    # Verifica se alguma palavra da lista está no texto extraído
    for palavra in palavras:
        if palavra.lower() in texto_extraido.lower():
            return True, palavra  # Retorna True e a palavra encontrada
    
    return False, None  # Retorna False se nenhuma palavra for encontrada

def esperar_tela(palavras):
    while True:
        encontrado, palavra = verificar_palavras_na_tela(palavras)
        if encontrado:
            print(f"A palavra '{palavra}' foi encontrada na tela!")
            break  
        else:
            print(f"Nenhuma palavra da lista foi encontrada. Verificando novamente em 1 segundo...")
        
        # Aguarda 1 segundos antes de verificar novamente
        time.sleep(1)



