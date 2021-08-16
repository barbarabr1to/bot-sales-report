import pyautogui as auto
import pandas    as pd
import pyperclip
import time 


# Passo 1: Abrir navegador
# Passo 2: Acessar link - https://drive.google.com/drive/folders/149xknr9JvrlEnhNWO49zPcw0PW5icxga?usp=sharing
# Passo 3: Entrar na pasta "exportar"
# Passo 4: Baixa arquivo - Vendas - Dez
# Passo 5: Calcular a quantiade e o faturamento dos produtos 
# Passo 6: Enviar email para diretoria


auto.PAUSE = 1

# MELHORIA: deixar resiliente a qualquer SO e browser
def open_browser():
    auto.hotkey('WIN', 'r')
    auto.write('C:\Program Files\Google\Chrome\Application\chrome.exe', interval=0.05)
    return auto.press('ENTER')

# MELHORIA: deixar resiliente caso não tenha conta google logada
def access_link():
    time.sleep(2)
    pyperclip.copy('https://drive.google.com/drive/folders/149xknr9JvrlEnhNWO49zPcw0PW5icxga?usp=sharing')
    auto.hotkey('CTRL', 'v')
    return auto.hotkey('ENTER')

def enter_folder():
    time.sleep(2)
    return auto.doubleClick(x=520, y=455)

# MELHORIA: certificar se o arquivo baixou
def download_file():
    time.sleep(2)
    auto.click(x=550, y=540, button='right')
    return auto.click(x=720, y=950)

# MELHORIA: identificar caminho de forma dinânmica
def calculate():
    spreadsheet = pd.read_excel(r'C:\Users\Bárbara\Downloads\Vendas - Dez (1).xlsx')
    quantity = spreadsheet['Quantidade'].sum()
    revenues = spreadsheet['Valor Final'].sum()
    print('\nQuantidade de produtos:', quantity)
    print(f'Faturamento: R$ {revenues},00\n')
    return quantity, revenues