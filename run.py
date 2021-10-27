import pyautogui as auto
import sales_report
import time

sales_report.open_browser('C:\Program Files\Google\Chrome\Application\chrome.exe')
sales_report.access_link('https://drive.google.com/drive/folders/1IhqBF3wQ37Ay2f1fkF2YQYgfRNwuUUMj?usp=sharing')
sales_report.enter_folder()
sales_report.download_file()
info = sales_report.calculate(r'C:\Users\Bárbara\Downloads\Vendas - Dez.xlsx')
sales_report.access_gmail('https://mail.google.com/')
sales_report.write_email('brittobsm@hotmail.com', 'Bárbara Brito', info)

# time.sleep(5)
# position = auto.position()
# print('>>>', position)

