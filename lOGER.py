import os
import logging
import glob
import re
from pathlib import Path, WindowsPath
from datetime import datetime, timedelta
import smtplib
from email.mime.multipart import MIMEMultipart      # Многокомпонентный объект
from email.mime.text import MIMEText                # Текст/HTML
from email.mime.image import MIMEImage              # Изображения
from email.header    import Header
from email.mime.base import MIMEBase
from email import encoders
import time
from os.path import basename
from email.mime.application import MIMEApplication
import mimetypes
import email.mime.application
import xlsxwriter
import pandas as pd

try:
    os.remove('results.xlsx')
except Exception as e: logging.exception(str(e))
workbook = xlsxwriter.Workbook('C:/Users/o.zadonskii/Desktop/lOGER/results.xlsx')
worksheet = workbook.add_worksheet()
#Почтовые адреса от кого и кому
addr_from = "o.zadonskiy@teh.expert"                # Адресат
#addr_to1 = "i.popov@teh.expert"
addr_to2 = "olegan.zadonskiy1@gmail.com"
#Пароль к почте
passwordMail  = "E6JCSx5uxK6KFEtukv7N"
#Актуальная дата (формат дата)
current_date = datetime.now().strftime('%d.%m.%Y')
#Формат для логирования
format = "%(asctime)s: %(message)s"
logging.basicConfig(filename = "logs/logs_" + current_date + ".txt", format=format, level=logging.INFO, datefmt="%H:%M:%S")
date_mask = r'\d{2}.\d{2}.\d{4}'
# Слова для поиска
DEST_words = [] #['error', 'ошибка', 'не загружен', 'не могу', 'not working', 'not configured', 'нет в схеме подключения', 'Не найден файл', 'не в стандартном месте']
# Путь к папке, в которой нужно искать файлы
root_dir = Path('D:/Systems/Kodeks-Intranet_SmallBase/logs')
current_time = int(time.time())
# Список для хранения результатов поиска
results = []
FILENAME = "C:/Users/o.zadonskii/Desktop/lOGER/results.txt"
Header = 'Отчет по сбору логов'
message = 'В логах найдены ошибки'
period = 3
filename = 'mask.txt'
filename1 = 'mask2.txt'

test_text = input ("Выберите маску 1 или 2 и нажмите enter:\n 1 - Поиск всех ошибок в системе \n 2 - Поиск ошибок авторизации\n")
test_number = int(test_text)
if test_number == 1:
    try:
        with open(filename, 'r', encoding='UTF-8') as file:
            for line in file:
                DEST_words.append(line.replace('\n', ''))
    except Exception as e: logging.exception(str(e))

else:
    polz = input ("Введите login пользователя и нажмите enter:")
    try:
        with open(filename1, 'r', encoding='UTF-8') as file:
            for line in file:
                DEST_words.append(line.replace('\n', ''))
    except Exception as e: logging.exception(str(e))
    DEST_words.append(polz)
    


print(DEST_words)
#Отправка почтового письма
def mailer(addr_from, passwordMail, addr_to, message, Header, path_to_pdf):
    try:
        #Формат сообщения
        msg = MIMEMultipart()                             # Создаем сообщение
        msg['From']    = addr_from                          # Адресат
        msg['To']      = addr_to                            # Получатель
        msg['Subject'] = Header                             # Заголовок    
        # Добавляем файл
        #part = MIMEText(message, 'plain')
        #msg.attach(part)
        msg.attach(MIMEText(message, "plain"))

        
        with open(path_to_pdf, "rb") as f:
            #attach = email.mime.application.MIMEApplication(f.read(),_subtype="pdf")
            attach = MIMEApplication(f.read(),_subtype="pdf")
        # encode into base64
        encoders.encode_base64(attach)
        attach.add_header('Content-Disposition','attachment',filename=str(path_to_pdf))
        msg.attach(attach)
        print(path_to_pdf)
        server = smtplib.SMTP('smtp.mail.ru', 587)           # Создаем объект SMTP
        #server.set_debuglevel(1)                         # Включаем режим отладки - если отчет не нужен, строку можно закомментировать
        server.starttls()                                   # Начинаем шифрованный обмен по TLS
        server.login(addr_from, passwordMail)
        
        server.sendmail(addr_from, addr_to, msg.as_string())  #Отправляем письмо
        server.quit()                                           #Покидаем сервер
        logging.info('Email sent!')                             #Успешная отправка
    except Exception as e: logging.exception(str(e))

try:
    with open("DestPath.txt", encoding='UTF-8') as file_in:
        lines = []
        for line in file_in:
            newline = line.replace('root_path = ','').replace('\n','').replace('email = ','').replace('logs_period = ','')                                                                                                              
            lines.append(newline)
    root_dir = Path(lines[0])
    addr_to2 = (lines[1])
    period = int((lines[2]))

except Exception as e: logging.exception(str(e))
print(period)
# Функция для рекурсивного поиска файлов
def find_files(path):
    for file in glob.iglob(os.path.join(path, '**'), recursive=True):
        full_path = os.path.join(path, file)
        if os.path.isfile(file):
            last_modified = os.path.getmtime(full_path)
            seconds_since_last_modification = current_time - last_modified
            # Проверяем, прошло ли менее суток с момента последней модификации
            if seconds_since_last_modification < (24 * 60 * 60 * period):
                with open(file, 'r', encoding='utf-8', errors='replace') as f:
                    content = f.readlines()

                    # Поиск слов в содержимом файла
                    for index, line in enumerate(content):
                        for word in DEST_words:
                            if re.search(word, line, flags=re.IGNORECASE):
                                match = re.search(date_mask, line)
                                if match:
                                    # Извлекаем дату из строки
                                    date_string = match.group()
                                    try:
                                        # Преобразуем строку в объект datetime
                                        date = datetime.strptime(date_string, '%d.%m.%Y')
                                    except ValueError:
                                        # Если преобразование не удалось, пропускаем строку
                                        continue

                                    # Вычитаем два дня из текущей даты
                                    two_days_ago = datetime.now() - timedelta(days=period)
                                    formatted_date = two_days_ago.strftime('%d.%m.%Y %H:%M:%S')
                                    formatted_date = datetime.strptime(formatted_date,'%d.%m.%Y %H:%M:%S')
                                    
                                    # Проверяем, меньше ли дата в строке, чем два дня назад
                                    print(date, "f")
                                    print(formatted_date, "c")
                                    if date >= formatted_date:
                                        print('true')
                                        results.append((file, word, index + 1, line))
                                        logging.info("We found an error on logs\n")
                                        break
                                else:
                                    results.append((file, word, index + 1, line))
                                    logging.info("We found an error on logs\n")
                                    break

# Запуск функции поиска файлов
find_files(root_dir)

# Выводим результаты в текстовый файл
with open('results.txt', 'w', encoding='utf-8') as outfile:
    for result in results:
        outfile.write(f'{result[0]} - error found at line {result[2]}: {result[3]}\n')

time.sleep(1)
#mailer(addr_from, passwordMail, addr_to2, message, Header, FILENAME)
worksheet.set_column('A:A', 20)


# Записываем данные в ячейки
worksheet.write(0, 0, "Path")
worksheet.write(0, 1, "Key")
worksheet.write(0, 2, "Line")
worksheet.write(0, 3, "Error")
row = 1
for result in results:
    for i in range(len(result)):
        worksheet.write(row, i, result[i])
    row += 1
row += 1
worksheet.freeze_panes('B2')
#for col_num in range(4):
#    worksheet.set_column(col_num, col_num, None, {'border': 1})
#rower = 'A1:D' + str(row + 1)
#worksheet.range(rower).set_border(top=True, bottom=True, left=True, right=True, diagonal=False, outline=True)
workbook.close()
#
#    try:

#   except Exception as e:
#        logging.error(e)

logging.info("Operation complete\n\n\n")
         











