import openpyxl
from urllib.parse import quote
import webbrowser
from time import sleep
import pyautogui


sleep(2)

workbook = openpyxl.load_workbook('clientes.xlsx')
pagina_cliente = workbook['Sheet1']


for linha in pagina_cliente.iter_rows(min_row=2):
      #nome, telefone, vencimento
      nome = linha[0].value
      telefone = linha[1].value
      vencimento = linha[2].value

      mensagem = f'Olá {nome}, bom final de semana!!'

      #link personalizado, com base nos dados da planilha
      link_mensgem_whatsapp = f'https://web.whatsapp.com/send?phone={telefone}&text={quote(mensagem)}'
      webbrowser.open(link_mensgem_whatsapp)
      sleep(10)
      try:
            seta = pyautogui.locateOnScreen('seta.png')
            sleep(10)
            pyautogui.click(seta)
            sleep(5)
            pyautogui.hotkey('ctrl', 'w')
            sleep(10)
      except:
            print(f'Não foi possivel enviar mensagem para {nome}')
            with open('erros.csv', 'a', newline='', encoding='utf-8') as arquivo:
                  arquivo.write(f'{nome},  {telefone}')
