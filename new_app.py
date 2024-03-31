"""
Demanada: criar uma automatização p/ enviar mensagens aos contatos de uma lista, informando a programação semanal das atividades.
"""
import openpyxl
import pyautogui
import webbrowser
import pandas as pd

from time import sleep
from selenium import webdriver
from urllib.parse import quote #→ configura a mensagem para não ocorrer o formato correto de link

#acessando o web.whatsapp

#driver = webdriver.Chrome()
#driver = webdriver.Firefox()

#driver.get('https://web.whatsapp.com/')
#sleep(20)
#pyautogui.hotkey('ctrl','w')
#sleep(10)

#print(pyautogui.position())

#acessando o web.whatsapp
#webbrowser.open('https://web.whatsapp.com/')
#sleep(20)
#pyautogui.hotkey('ctrl','w')
#sleep(2)


#Inserir cada célula de cada linha em um campo do sistema
#lendo o arquivo
workbook = openpyxl.load_workbook('contatos.xlsx')
#indicando a aba da planilha
contatos = workbook['sexta']
#contatos = workbook['domingo']



#l localizando a posi;'ao do mouse
#print(pyautogui.position())
#pyautogui.moveTo(1860, 961)

def check_time():
    current_time = pd.Timestamp.now().time().hour # current_time = current_time pd.Timestamp.now().time
    time_period = None
    if current_time >= 5 and current_time < 12:
        time_period = "Bom dia,"
    elif current_time >= 12 and current_time < 18:
        time_period = "Boa tarde,"
    
    else:
        time_period = "Boa noite,"

    return time_period
period = check_time()
#print(f"{period}")
#
for linha in contatos.iter_rows(min_row=2):
    #parametro min_row= significa que deve começar a ler a partir da linha 2 da coluna
    #0-NOME	1-SUFIXO	2-TELEFONE	3-GOHOSHI
    nome = linha[0].value
    sufixo = linha[1].value
    telefone = linha[2].value
    gohoshi = linha[3].value
    dia = '12 de abril'

    #sexta
    texto1 = f'{period} {nome}!\nTudo bem por ai?\nAqui é o Bruno Geraldine.\nGostaria de convidá-l{sufixo} para participar do Gohoshi preparativo para a cerimonia agradecimento mensal no Jun-Dojo de Goiânia.'
    texto2 = f'E em representacao aos seus antepassados oferecer esforcos voluntarios no Gohoshi de {gohoshi} no dia {dia} (sexta-feira).'
    texto3 = f'As atividades iniciam-se as 09h e 30m da manha. Seria uma grande alegria contar com sua presenca!'
    texto4 = f'Qualquer duvida, estou a disposicao. \nSinta-se a vontade em responder por aqui.. \n Muito obrigado!'
    
     #trantando possiveis erros
    try:
        # Criando o link personalizado
        # formato padrão: https://web.whatsapp.com/send?phone=888&text="hhh"
        link_wtsp = f'https://web.whatsapp.com/send?phone={int(telefone)}&text={quote(texto1)}'
        #sleep(30)
        # Abrindo o browser no navegador
        webbrowser.open(link_wtsp)
        #print(link_wtsp)
        sleep(20)
        #pyautogui.moveTo(2953, 1631)
        pyautogui.click(x=2953, y=1631)
        sleep(2)
        #pyautogui.mouseDown(button='left')
        #pyautogui.leftClick(1.5)

        # Texto 2
        #pyautogui.moveTo(2793, 1621)
        pyautogui.click(x=2793, y=1621)
        #pyautogui.leftClick(1.5)
        #Point(x=2793, y=1621) → escrever
        sleep(2)
        pyautogui.write(texto2, interval=0.02)
        #pyautogui.moveTo(2953, 1631)
        pyautogui.click(x=2953, y=1631)
        sleep(2)

        # Text 3
        #pyautogui.moveTo(2793, 1621)
        pyautogui.click(x=2793, y=1621)
        #pyautogui.leftClick(1.5)
        pyautogui.write(texto3, interval=0.02)
        #pyautogui.moveTo(2953, 1631)
        pyautogui.click(x=2953, y=1631)
        sleep(2)
        #pyautogui.leftClick(1.5)

        # Texto 4
        #pyautogui.moveTo(2793, 1621)
        pyautogui.click(x=2793, y=1621)

        #pyautogui.leftClick(2)
        pyautogui.write(texto4, interval=0.02)
        #pyautogui.moveTo(2953, 1631)
        pyautogui.click(x=2953, y=1631)

        sleep(2)
        #pyautogui.leftClick(1.5)
        
        #pyautogui.hotkey('press')
        sleep(1)
#       
        pyautogui.hotkey('ctrl','w')
        sleep(2)
        
##
    except:
        print(f'Não foi possivel enviar mensagem para {nome}.')
        with open('erros.csv', 'a', newline='', encoding='utf-8') as arquivo:
            arquivo.write(f'{nome}, {telefone}')

pyautogui.alert("O código foi finalizado. Você já pode utilizar o computador!")
