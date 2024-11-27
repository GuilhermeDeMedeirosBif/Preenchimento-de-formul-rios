import pyautogui as pygui
from time import sleep
import openpyxl
import webbrowser

webbrowser.open('https://forms.gle/ED57ScaAvGWwhoGv9')
workbook = openpyxl.load_workbook('Planilha_pessoas.xlsm')
participantes_planilha = workbook['Planilha1']
pygui.click(1825, 35, duration=2)
sleep(2)

for linha in participantes_planilha.iter_rows(min_row=2):
    nome = linha[0].value
    cpf = linha[1].value
    ingresso = linha[2].value
    n_ingresso = linha[3].value
    presente = linha[4].value
    pix = linha[5].value

    nome_completo = pygui.click(600, 500, duration=2)
    preencher_nome = pygui.write(f'{nome}')
    sleep(2)
    cpf_f = pygui.click(600, 725, duration=2)
    preencher_cpf = pygui.write(f'{cpf}')
    print(cpf)
    sleep(2)
    pygui.scroll(-625)
    sleep(2)

    if ingresso == 'Vip':
        pygui.click(593, 218,duration=2)
    elif ingresso == 'Normal':
        pygui.click(593, 267,duration=2)
    elif ingresso == 'Meia-Entrada':
        pygui.click(593, 317,duration=2)
    
    n_ingresso_f = pygui.click(600, 500, duration=2)
    preencher_ingresso = pygui.write(f'{n_ingresso}')
    
    if presente == 'sim':
        pygui.click(593, 680,duration=2)
    else:
        pygui.click(593, 732)
        pygui.click(600, 915, duration=2)
        pygui.write(f'{pix}')
    
    pygui.scroll(-500)
    pygui.click(604, 788, duration=2)
    sleep(5)
    pygui.scroll(1000)
    pygui.hotkey('ctrl', 'w')
    sleep(2)
    webbrowser.open('https://forms.gle/ED57ScaAvGWwhoGv9')
    sleep(2)
