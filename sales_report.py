import pyautogui as auto
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
    auto.write("C:\Program Files\Google\Chrome\Application\chrome.exe", interval=0.05)
    return auto.press('ENTER')

# MELHORIA: deixar resiliente caso não tenha conta google logada
def access_link():
    time.sleep(2)
    pyperclip.copy("https://drive.google.com/drive/folders/149xknr9JvrlEnhNWO49zPcw0PW5icxga?usp=sharing")
    auto.hotkey('CTRL', 'v')
    return auto.hotkey('ENTER')

