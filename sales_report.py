from   datetime  import date
import pyautogui as auto
import pandas    as pd
import pyperclip
import time 
import os.path


auto.PAUSE = 1

def open_browser():
    auto.hotkey('WIN', 'r')
    auto.write('C:\Program Files\Google\Chrome\Application\chrome.exe', interval=0.01)
    return auto.press('ENTER')


def access_link(link):
    time.sleep(2)
    pyperclip.copy(link)
    auto.hotkey('CTRL', 'v')
    return auto.hotkey('ENTER')
   

def enter_folder():
    time.sleep(2)
    return auto.doubleClick(x=510, y=385)


def download_file():
    time.sleep(2)
    auto.click(x=500, y=385, button='right')
    return auto.click(x=670, y=900)


def calculate():
    file = r'C:\Users\Bárbara\Downloads\Vendas - Dez.xlsx'

    while(not os.path.exists(file)):
        time.sleep(1)
        print('\n::::: Aguardando download de arquivo')
        
    spreadsheet = pd.read_excel(file)
    quantity = spreadsheet['Quantidade'].sum()
    revenues = spreadsheet['Valor Final'].sum()
    
    print('\nQuantidade de produtos:', quantity)
    print(f'Faturamento: R$ {revenues},00\n')
    
    return quantity, revenues


def access_gmail(link):
    auto.hotkey('CTRL', 't')
    return access_link(link)


def write_email(email, collaborator):
    time.sleep(2)
    auto.click(x=130, y=260)
    auto.write(email)
    auto.press('TAB')
    auto.press('TAB')
    current_date = date.today().strftime('%d/%m/%Y')
    pyperclip.copy(f'RELATÓRIO DE VENDAS - {current_date}')
    auto.hotkey('CTRL', 'v')
    auto.press('TAB')
    info = calculate()
    pyperclip.copy(f"""
    Prezados, 
    Abaixo estão as informações das vendas diária
   
    Quantidade de produtos vendidos: {info[0]}
    Faturamento: R$ {'{0:,}'.format(info[1]).replace(',','.')},00

    At.te, {collaborator}
    """)
    auto.hotkey('CTRL', 'v')
    return auto.hotkey('CTRL', 'ENTER')