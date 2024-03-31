"""
Demanada: criar uma automatização p/ enviar mensagens aos contatos de uma lista, informando a programação semanal das atividades.
"""
import openpyxl
import pyautogui
import webbrowser

from time import sleep
# → configura a mensagem para não ocorrer o formato correto de link
from urllib.parse import quote

#acessando o web.whatsapp
webbrowser.open('https://web.whatsapp.com/')
sleep(20)
pyautogui.hotkey('ctrl','w')
sleep(2)

# Descrever os passos manuais e depois transformar isso em código
workbook = openpyxl.load_workbook('contatos.xlsx')
# acessando cada aba da planilha (se houver + de uma aba deve-se chamar o nome de cada uma)
pag_clientes = workbook['Página1']

# iterando e lendo cada linha
for linha in pag_clientes.iter_rows(min_row=2):
    # nome, telefone
    nome = linha[0].value
    sufixo = linha[1].value
    telefone = linha[2].value
    gohoshi = linha[3].value

    # Texto gohoshi cerimonia
    texto = f'Olá {nome}, boa noite!\nGostaria de convidá-l{sufixo} para oferecer Gohoshi de {gohoshi}, em representação aos seus antepassados, na cerimônia de agradecimento mensal do mês de março. \nSeria uma grande alegria contar com sua presença! \nQualquer dúvida, estou à disposição. \nSinta-se à vontade em responder por aqui.. \n Muito obrigado!'
    # texto = f'Olá {nome}, boa noite!\nGostaria de convidá-l{sufixo} para participar do Gohoshi de \norganização do Jun-Dojo de Goiânia, e em representação aos \nseus antepassados oferecer esforços no Gohoshi de {gohoshi}. \nSeria uma grande alegria contar com sua presença! \nQualquer dúvida, estou à disposição. \nSinta-se à vontade em responder por aqui.. \n Muito obrigado!'

    # print(nome, "--", telefone)

    # trantando possiveis erros
    try:
        # Criando o link personalizado
        # formato padrão: https://web.whatsapp.com/send?phone=888&text="hhh"
        link_wtsp = f'https://web.whatsapp.com/send?phone={int(telefone)}&text={quote(texto)}'
        # print(link_wtsp)
        # Abrindo o browser no navegador
        webbrowser.open(link_wtsp)
        sleep(20)
#        #Clinando na seta para envio da mensagem
        # seta = pyautogui.locateCenterOnScreen('seta.png')
        pyautogui.hotkey('enter')
        sleep(4)
        # pyautogui.click(seta[0], seta[1])
        # pyautogui.hotkey('enter')

#        #Fechando a pagina para abrir nova pagina e enviar mensagem seguinte
        pyautogui.hotkey('ctrl', 'w')
        sleep(2)
#
    except:
        print(f'Não foi possivel enviar mensagem para {nome}.')
        with open('erros.csv', 'a', newline='', encoding='utf-8') as arquivo:
            arquivo.write(f'{nome}, {telefone}')
#
