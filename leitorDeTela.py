from PIL import Image
import pytesseract
import pyautogui
import time

# Caminho para o executável do Tesseract
pytesseract.pytesseract.tesseract_cmd = r'C:/Users/prr8ca/AppData/Local/Programs/Tesseract-OCR/tesseract.exe'

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
            print(f"Nenhuma palavra da lista foi encontrada. Verificando novamente em 2 segundos...")
        
        # Aguarda 2 segundos antes de verificar novamente
        time.sleep(2)

