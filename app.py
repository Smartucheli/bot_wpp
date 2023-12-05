import openpyxl 
from urllib.parse import quote
import webbrowser
from time import sleep
import pyautogui

planilha = openpyxl.load_workbook('planilha_exemplo.xlsx')

pagina_numeros = planilha['Planilha1']

for cliente in pagina_numeros.iter_rows(min_row=2):

    nome = cliente[0].value
    telefone = cliente[1].value
    email = cliente[2].value

    if nome is None and telefone is None and email is None:
        print('Não existe mais contatos processo encerrado')
        break

    mensagem = f'Olá esse é um teste do nosso bot, você {nome} com o telefone {telefone} e email {email} teve a sorte de fazer parte do nosso curso!!!'

    try:
        link_mensagem_wpp = f'https://web.whatsapp.com/send?phone={telefone}&text={quote(mensagem)}'

        webbrowser.open(link_mensagem_wpp)
        sleep(20)
        
        seta = pyautogui.locateCenterOnScreen('seta.png')
        sleep(2)

        pyautogui.click(seta[0], seta[1])
        sleep(2)

        pyautogui.hotkey('ctrl', 'w')
        sleep(3)
    except:
        print(f'não enviou mensagem para {nome}')
        with open('erros.csv', 'a', newline='', encoding='utf-8') as arquivo:
            arquivo.write(f'{telefone} -- {nome}')