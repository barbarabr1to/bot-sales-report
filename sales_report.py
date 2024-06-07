from   datetime  import date
import pyautogui as auto
import pandas    as pd
import pyperclip
import time 
import os.path


auto.PAUSE = 1

def open_browser(path_browser):
    auto.hotkey('WIN', 'r')
    auto.write(path_browser, interval=0.01)
    return auto.press('ENTER')


def access_link(link):
    time.sleep(2)
    pyperclip.copy(link)
    auto.hotkey('CTRL', 'v')
    return auto.hotkey('ENTER')
   

def enter_folder():
    time.sleep(2)
    return auto.doubleClick(x=535, y=405)


def download_file():
    time.sleep(2)
    auto.click(x=540, y=500, button='right')
    return auto.click(x=710, y=900)


def calculate(path_file):

    while(not os.path.exists(path_file)):
        time.sleep(1)
        print('\n::::: Aguardando download de arquivo')
        
    spreadsheet = pd.read_excel(path_file)
    quantity = spreadsheet['Quantidade'].sum()
    revenues = spreadsheet['Valor Final'].sum()
    
    print('\nQuantidade de produtos:', quantity)
    print(f'Faturamento: R$ {revenues},00\n')
    
    return quantity, revenues


def access_gmail(link):
    auto.hotkey('CTRL', 't')
    return access_link(link)


def write_email(email, collaborator, info):
    time.sleep(2)
    auto.click(x=55, y=260)
    auto.write(email)
    auto.press('TAB')
    auto.press('TAB')
    current_date = date.today().strftime('%d/%m/%Y')
    pyperclip.copy(f'RELATÓRIO DE VENDAS - {current_date}')
    auto.hotkey('CTRL', 'v')
    auto.press('TAB')
    pyperclip.copy(f"""
    Prezados, 
    Abaixo estão as informações das vendas diária
   
    Quantidade de produtos vendidos: {info[0]}
    Faturamento: R$ {'{0:,}'.format(info[1]).replace(',','.')},00

    At.te, {collaborator}
    """)
    auto.hotkey('CTRL', 'v')
    return auto.hotkey('CTRL', 'ENTER')

# Testando comandos git
print("Relário de vendas")