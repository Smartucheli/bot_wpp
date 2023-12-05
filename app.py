import openpyxl 
from urllib.parse import quote
import webbrowser

planilha = openpyxl.load_workbook('planilha_exemplo.xlsx')

pagina_numeros = planilha['Planilha1']

for cliente in pagina_numeros.iter_rows(min_row=2):

    nome = cliente[0].value
    telefone = cliente[1].value
    email = cliente[2].value

    mensagem = f'Olá esse é um teste do nosso bot, você {nome} com o telefone {telefone} e email {email} teve a sorte de fazer parte do nosso curso!!!'

    link_mensagem_wpp = f'https://web.whatsapp.com/send?phone={telefone}&text={quote(mensagem)}'

    webbrowser.open(link_mensagem_wpp)

    input('')