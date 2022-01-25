import imaplib
import smtplib
import pymorphy2
import os
import re
import email
from email.header import decode_header
from email.header import Header
from email.mime.multipart import MIMEMultipart
from email.mime.text import MIMEText
from email.mime.image import MIMEImage
from email.utils import formataddr

# учетные данные
username = ""
password = ""

# возврат слова к начальной форме и заполнение списка ими
def norm(x):
    morph = pymorphy2.MorphAnalyzer()
    p = morph.parse(x)[0]
    return p.normal_form
# ÷èñòûé òåêñò äëÿ ñîçäàíèÿ ïàïêè
def clean(text):
    return "".join(c if c.isalnum() else "_" for c in text)

bodyspisok = set()
spisokotpravky = set()

imap = imaplib.IMAP4_SSL("imap.gmail.com") # cоздание класса IMAP4 с SSL
imap.login(username, password) # аутентификаци
# создать сервер
server = smtplib.SMTP('smtp.gmail.com: 587')
server.starttls()
server.login(username, password)

# выбор €щика писем
status, messages = imap.select('INBOX')
message = imap.status('INBOX', '(UNSEEN)') # провер€ем наличие непрочитанных
unreadcount = re.findall(r'\d+', str(message)) # ищем количество непрочитанных
unreadcount = int(unreadcount[0])
messages = int(messages[0])

for i in range(messages, messages-unreadcount, -1):
    res, msg = imap.fetch(str(i), "(RFC822)")
    for response in msg:
        if isinstance(response, tuple):
            # получение текста
            msg = email.message_from_bytes(response[1])
            # расшифровка текста
            subject, encoding = decode_header(msg.get("Subject"))[0]
            if isinstance(subject, bytes):
                # перевод в текст
                subject = subject.decode(encoding)
                subsplit = subject.split()
                for i in subsplit:
                    bodyspisok.add(norm(i))
            # расшифровка отправител€
            From, encoding = decode_header(msg.get("From"))[0]
            if isinstance(From, bytes):
                From = From.decode(encoding)
                # перебор частей письма
            for part in msg.walk():
                # получение типа содержимого письма
                content_type = part.get_content_type()
                content_disposition = str(part.get("Content-Disposition"))
                if content_type == 'text/plain':
                    body = part.get_payload(decode=True).decode()
                    # удаление знаков препинани€
                    bodysplit = re.findall(r'\w+', body)
                    for i in bodysplit:
                        bodyspisok.add(norm(i))
                if "attachment" in content_disposition:
                    # скачать вложение
                    filename = part.get_filename()
                    if filename:
                        folder_name = clean(subject)
                        if not os.path.isdir(folder_name):
                            # сделать папку дл€ этого письма, названную как тема
                            os.mkdir(folder_name)
                        filepath = os.path.join(folder_name, filename)

    break_out_flag = False
    # проверка на ключевые слова
    nazvpap = r'' # расположение папки с отделами
    for nazvpap1 in os.listdir(nazvpap): # ищем папку с отделом
        for adresa in os.listdir(os.path.join(nazvpap, nazvpap1)): # ищем название файла с ключевыми словами
            file = os.path.join(nazvpap, nazvpap1, adresa) # записываем путь до файла с ключевыми словами
            adresa1 = list(adresa.split('; ')) # раздел€ем название на им€ и адрес с форматом файла
            adresa1 = adresa1[1].split('.') # раздел€ем адрес на адрес и формат файла
            adresa1.pop(2) # удал€ем формат файла
            adresa1 = '.'.join(adresa1) # превращаем список в строку
            adresa2 = adresa1.split('@')
            adresa2.pop(1) # нужно дл€ проверки нахождени€ адреса человека в сообщении
            adresa2 = ''.join(adresa2) 
            with open (file, 'r') as fp:
                file1 = (fp.read().split(', '))
                file1.append(adresa2)
            for i in bodyspisok: # ищем ключевые слова и заполн€ем spisokotpravky адресами
                slovo = i
                if any(i in file1 for proverka in bodyspisok):
                    spisokotpravky.add(adresa1)
                if slovo == adresa2:
                    spisokotpravky.clear()
                    spisokotpravky.add(adresa1)
                    break_out_flag = True
                    break
            if break_out_flag:
                break

    # создание объекта сообщени€
    msg = MIMEMultipart()
    # установка параметра сообщени€
    msg['From'] = formataddr((str(Header(From, 'utf-8')), 'from@mywebsite.com'))
    msg['To'] = ','.join(spisokotpravky)
    msg['Subject'] = subject
    # работа со вложением
    try:
        with open (filepath, 'rb') as fp:
            file = fp.read()
            msg.attach(MIMEImage(file, name=filepath))
    except:
        pass
    msg.attach(MIMEText(body)) # добавить текст
    server.sendmail(msg['From'], msg['To'].split(','), msg.as_string()) # отправить сообщение через сервер
    bodyspisok.clear()
    spisokotpravky.clear()
server.quit()
imap.close()
imap.logout()