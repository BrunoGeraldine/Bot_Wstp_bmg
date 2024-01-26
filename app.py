"""
Demanada: criar uma automatização p/ enviar mensagens aos contatos de uma lista, informando a programação semanal das atividades.
"""
import openpyxl
import pyautogui
import webbrowser

from time import sleep
from urllib.parse import quote #→ configura a mensagem para não ocorrer o formato correto de link

#acessando o web.whatsapp
webbrowser.open('https://web.whatsapp.com/')
sleep(20)
pyautogui.hotkey('ctrl','w')
sleep(2)

# Descrever os passos manuais e depois transformar isso em código
workbook = openpyxl.load_workbook('contatos.xlsx')
# acessando cada aba da planilha (se houver + de uma aba deve-se chamar o nome de cada uma)
pag_clientes = workbook['Página1']

#iterando e lendo cada linha
for linha in pag_clientes.iter_rows(min_row=2):
    #nome, telefone
    nome = linha[0].value
    telefone = linha[1].value
    aniversario = linha[2].value

    texto = f'Olá {nome}, essa é uma mensagem de teste para um projeto que estou desenvolvendo. Peço que desconsidere. Muito obrigado!'
    #texto = f'Olá {nome}, essa é uma mensagem de teste para um projeto que estou desenvolvendo para lembrar que você nascei no dia {aniversario.strftime('%d/%m/%Y')}. Peço que desconsidere. Muito obrigado!'

    #print(nome, "--", telefone)
        
    #trantando possiveis erros
    try:
        # Criando o link personalizado
        # formato padrão: https://web.whatsapp.com/send?phone=888&text="hhh"
        link_wtsp = f'https://web.whatsapp.com/send?phone={int(telefone)}&text={quote(texto)}'
        #print(link_wtsp)
        # Abrindo o browser no navegador
        webbrowser.open(link_wtsp)
        sleep(20)
#        #Clinando na seta para envio da mensagem
        #seta = pyautogui.locateCenterOnScreen('seta.png')
        pyautogui.hotkey('enter')
        sleep(4)
        #pyautogui.click(seta[0], seta[1])
        #pyautogui.hotkey('enter')
        
#        #Fechando a pagina para abrir nova pagina e enviar mensagem seguinte
        pyautogui.hotkey('ctrl','w')
        sleep(2)
#
    except:
        print(f'Não foi possivel enviar mensagem para {nome}.')
        with open('erros.csv', 'a', newline='', encoding='utf-8') as arquivo:
            arquivo.write(f'{nome}, {telefone}')
#