import pyautogui as auto
import sales_report
import time

sales_report.open_browser()
sales_report.access_link('https://drive.google.com/drive/folders/149xknr9JvrlEnhNWO49zPcw0PW5icxga?usp=sharing')
sales_report.enter_folder()
sales_report.download_file()
sales_report.calculate()
sales_report.access_gmail('https://mail.google.com/')
sales_report.write_email('brittobsm@hotmail.com', 'Bárbara Brito')

# time.sleep(3)
# position = auto.position()
# print('>>>', position)


