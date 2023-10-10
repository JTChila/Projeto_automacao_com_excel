import openpyxl
import pyautogui
import time
from PIL import ImageGrab
import sys

# Substitua 'caminho_completo_para_o_arquivo' pelo caminho completo do seu arquivo Excel
caminho_arquivo = 'automacao.xlsx'

workbook = openpyxl.load_workbook(caminho_arquivo)
pagina_teste = workbook['teste']

# Inicialize uma variável para contar as iterações
contador = 0

# Especifica as coordenadas exatas (x e y) do pixel que você deseja verificar
posicao_x = 740
posicao_y = 300

# Especifica a cor que você deseja verificar (no formato RGB)
cor_desejada = (33,150,243)  

for linha in pagina_teste.iter_rows(min_row=1):
    pyautogui.click(713,144, duration=0.3)
    pyautogui.write(str(linha[0].value))
    pyautogui.click(737,192, duration=0.3)
    pyautogui.write(str(linha[1].value))
    pyautogui.click(984,313, duration=0.3)

    while True:
        # Captura uma captura de tela da área que contém o pixel desejado
        screenshot = ImageGrab.grab(bbox=(posicao_x, posicao_y, posicao_x + 1, posicao_y + 1))

        # Obtém a cor do pixel da captura de tela
        cor_pixel = screenshot.getpixel((0, 0))

        if cor_pixel == cor_desejada:
            break  # Se a cor do pixel for igual à cor desejada, saia do loop

        print("A cor do pixel nao corresponde a cor desejada. Aguardando 2 segundos...")
        time.sleep(2)  # Aguarde 2 segundos e depois tente novamente

    # Incrementa o contador de iterações
    contador += 1

# Exibe o número de iterações quando a automação terminar
print(f"O loop rodou {contador} vezes.")
