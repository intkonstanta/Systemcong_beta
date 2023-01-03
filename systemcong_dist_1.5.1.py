# -*- coding: utf-8 -*-
import pandas # библиотека для работы с данными
import pathlib # библиотека для работы с путями
import time # библиотека для работы с внутренним временем на устройстве
import smtplib # библиотека для работы с сетевым протоколом предназначенным для передачи электронной почты в сетях TCP/IP, (ESMTP)
import ssl # библиотека для работы с SSL протоколом (для безопасной передачи данных почты)
import ntplib # библиотека для работы с серверами реального (сетевого) времени
import logging # библиотека логирования
from pathlib import Path  # модуль создания пути до файла
from time import ctime # модуль для работы с данными времени (модль ctime нужен для преобразования количества секунд с начала эпохи в дату)
from email.mime.multipart import MIMEMultipart # Распаковка подклассов для работы с электронными письмами (дополнительные файлы (gif, media and other))
from email.mime.text import MIMEText # Распаковка подклассов для работы с электронными письмами (текст)from email.mime.image import MIMEImage # Распаковка подклассов для работы с электронными письмами (изображения)



def convert_data_format_DMY(time_data_unformat): # функция для конвертирования реального времени в полностью цифровой формат
    data_time_now = [] # создание списка для записи разделенных элементов строки с датой и временем из функции time.ctime()
    time_format = [] # создание списка для записи разделенных элементов строки со временем из списка data_time_now[]
    data_time_now = time_data_unformat.split() # разделяем элементы на отдельные через split
    day_convert = data_time_now[2] # выделяем день из списка data_time_now
    if len(day_convert) == 1: # проверяем, если строка с числом в виде 'x', то припишем к нем спереди ноль, чтобы было '0x'
        day_convert = '0' + day_convert #пример: '1' --> '01'
    month_convert = num_of_month[data_time_now[1]] # выделяем месяц из списка data_time_now, и конвертируем в формат числа через словарь num_of_month
    year_convert = data_time_now[4] # выделяем год из списка data_time_now
    time_unformat = (data_time_now[3]).split(':') # не отформатированное  время (чч:мм:сс), разделяем на отдельные элементы при помощи split
    hour_convert = time_unformat[0] # выделяем часы из списка time_unformat
    min_conver = time_unformat[1] # выделяем минуты из списка time_unformat
    sec_convert = time_unformat[2] # выделяем секунды из списка time_unformat
    logging.debug("[Succeess]-def-convert_data_format_DMY")

    return [day_convert, month_convert, year_convert, hour_convert, min_conver, sec_convert] # функция возвращает день, месяц, год, час, мин, сек в виде целых чисел


def user_data_convert_format_DMY(unformat_user_data): # конвертируем дату из таблицы в эксель, в список для более удобной работы
    unformat_user_data = str(unformat_user_data) # переводим тип данных type_data, в тип строка
    format_user_data = [] # подгтавливаем пустой список для записи в него чистых значений, без наков препинания и т.п.
    format_user_data = unformat_user_data.split() #отделяем время от даты
    format_user_data = format_user_data[0] # рассмотрим элемент даты
    format_user_data = format_user_data.split('-') # отделим день, месяц, год
    user_data_year = format_user_data[0] # запишем год
    user_data_month = format_user_data[1] # запишем месяц
    user_data_day = format_user_data[2] # запишем день
    logging.debug("[Succeess]-def-user_data_convert_format_DMY")

    return [user_data_day, user_data_month, user_data_year] # возвращаем список [день, месяц, год]


def real_time_request(delta): # функция
    try:
        response = c.request('time.windows.com')  # указываем путь (ссылку) до сервера ntp и создаём запрос
        e_time = ctime(response.tx_time)
    except Exception:
        e_time = ctime(time.time()+delta)
        logging.warning("-ntp server not answered")
    return e_time


num_of_month = { # словарь для преобразования названий месяцев в цифровое значение
    'Jan': '01',
    'Feb': '02',
    'Mar': '03',
    'Apr': '04',
    'May': '05',
    'Jun': '06',
    'Jul': '07',
    'Aug': '08',
    'Sep': '09',
    'Oct': '10',
    'Nov': '11',
    'Dec': '12'

}

dir_path = pathlib.Path.cwd() # получаем путь до папки с проектом
path = Path(dir_path, 'users_info.xlsx') # дополняем путь до файла с данными о пользователях
path = str(path) # переводим в формат строки
main_table = pandas.read_excel(path, sheet_name='list1') # таблица с информацией о пользователях, указываем страницу с которой работаем и путь до файла с таблицей
flag_sent = 0 # задаем стартовое значение флагу, что мы не отправляли письма сегодняв
msg_theme = 'Test8 msg theme' # присвоение экземпляру пиьсма его темы
body = '' # основной текст письма(не используем, т.к. есть html)
html = '' # переменная, в которую будем записывать html с поздравлениями
logging.basicConfig(level=logging.DEBUG, filename="logging_info.log", filemode="w",
                    format="%(asctime)s %(levelname)s %(message)s")

with open("time_to_send.txt", "r") as file:#откроем файл и посмотрем время введённое пользователем (в часах)
    time_to_send_mails = file.read() # время (час), в который начинать отправлять письма с поздравлением (вводить в формате строки, в виде двух значного числа)
    if len(time_to_send_mails) == 1:  # проверяем, если строка с числом в виде 'x', то припишем к нем спереди ноль, чтобы было '0x'
        time_to_send_mails = '0' + time_to_send_mails
    file.close()

logging.info(f"\nInput data:\nTime to sent:{time_to_send_mails} hours\n")


#html_woman - переменная содержащая html с gif файлом для поздравления женщин
html_woman = """\
<html>
  <body>
    <p class="aligncenter">
       <img src=https://s4.gifyu.com/images/gif_for_woman.gif width="1280" height="800" alt="woman">
    </p>
    <style>
    .aligncenter {
        margin: 0 auto;
        text-align: center;
    }
    </style>
  </body>
</html>
"""

#html_man - переменная содержащая html с gif файлом для поздравления мужчин
html_man = """\
<html>
  <body>
    <p class="aligncenter">
       <img src=https://s4.gifyu.com/images/gif_for_man.gif width="1280" height="800" alt="man">
    </p>
    <style>
    .aligncenter {
        margin: 0 auto;
        text-align: center;
    }
    </style>
  </body>
</html>
"""

#html_anniversary - переменная содержащая html с gif файлом для юбилейного поздравления
html_anniversary = """\
<html>
  <body>
    <p class="aligncenter">
       <img src=https://s4.gifyu.com/images/anniversary.gif width="1280" height="800" alt="anniversary">
    </p>
    <style>
    .aligncenter {
        margin: 0 auto;
        text-align: center;
    }
    </style>
  </body>
</html>
"""

password = "Alex.x1234" #пароль для входа в почту отправителя
addr_from = "testmailforprogs@yandex.ru" # почта отправителя
c = ntplib.NTPClient() # указываем, что работаем с ntp протоколом
response = c.request('time.windows.com') # указываем путь (ссылку) до сервера ntp
time_delta = response.tx_time - time.time()
all_user_data = main_table['data'].tolist() # список с датами пользователей
all_user_gender = main_table['gender'].tolist() # пол пользователя
all_user_mail = main_table['mail'].tolist() # список с почтами пользователей
all_user_name = main_table['name'].tolist() # список имён пользователей


while True: # запускаем бесконечный цикл, т.к. программа работает постоянно
    real_time = convert_data_format_DMY(real_time_request(time_delta)) # записываем реальное время, и преобразуем его в список

    if real_time[3] == time_to_send_mails: # посмотрим, нужно ли сейчас (в этот час) отправлять письма
        for user_num_index, user_data in enumerate(all_user_data): # рассмотрим дату пользователя из общего списка дат пользователей
            format_user_data = user_data_convert_format_DMY(user_data)  # приведём в формат DIY (тип данных список)
            if format_user_data[:2] == real_time[:2]: # если день и месяц сегодня совпадает с указанными в таблице
                logging.info(f" \nNew user info:\nUser mail:{all_user_mail[user_num_index]}\nUser data:{all_user_data[user_num_index]}\nUser name:"+all_user_name[user_num_index]+"\nUser gender:"+all_user_gender[user_num_index]+f"\nUser index:{user_num_index}\n ")


                addr_to = all_user_mail[user_num_index]# почта получателя
                msg = MIMEMultipart()  # создание экзепляра письма
                msg['From'] = addr_from  # присвоение экземпляру письма почты отправителя
                msg['To'] = addr_to # присвоение экземпляру письма почты получателя
                msg['Subject'] = msg_theme # присвоение экземпляру письма темы письма

                context = ssl.create_default_context() # подключение шифрования данных по SSL протоколу

                if (int(real_time[2]) - int(format_user_data[2])) % 5 == 0: #проверяем, юбилей или нет
                    html = html_anniversary # записываем в html файл юбилейное поздравление
                else:
                    if all_user_gender[user_num_index] == 'м': # если пользователь мужчина
                        html = html_man # записываем в html файл поздравление мужчине

                    if all_user_gender[user_num_index] == 'ж': # если пользователь женщина
                        html = html_woman

                msg.attach(MIMEText(body, 'plain'))  # загрузка содержимого в письмо
                msg.attach(MIMEText(html, 'html'))
                server = smtplib.SMTP_SSL('smtp.yandex.ru', 465) # подключение yandex почте по протоколу smtp
                server.login(addr_from, password) # вход в почту отправителя
                server.send_message(msg) # отправка сообщения
                server.quit() # закрытие сессии почты

        logging.debug('[Time]-wait: 3600s')
        time.sleep(3600)  # если да, то ждём один час, чтобы следующее условие не выполнилось до следующего дня

    else:
        real_time = convert_data_format_DMY(real_time_request(time_delta))
        logging.debug('[Time]-wait: 60s')
        time.sleep(60) # делаем ожидание одну минуту, чтобы не гонять цикл с условиями, и не тратить ресурс
#1.5.1

