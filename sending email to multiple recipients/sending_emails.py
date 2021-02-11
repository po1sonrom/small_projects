
#ИМПОРТИРУЕМ НУЖНЫЕ ПАКЕТЫ
import openpyxl, smtplib
from email.mime.multipart import MIMEMultipart
from email.mime.text import MIMEText
from email.mime.image import MIMEImage
from email.mime.base import MIMEBase
from email import encoders

#РАБОТА С EXCEL ФАЙЛОМ, КОТОРЫЙ СОДЕРЖИТ АДРЕСА ДЛЯ РАССЫЛКИ
#ОТКРЫВАЕМ НУЖНЫЙ ФАЙЛ 'test_for_emails.xlsl'.
wb = openpyxl.load_workbook('test_for_emails.xlsx')

#УКАЗЫВАЕМ НАЗВАНИЕ ЛИСТА В ФАЙЛЕ ИЗ КОТОРОГО БУДЕМ БРАТЬ ИНФОРМАЦИЮ ДЛЯ РАССЫЛКИ
sheet = wb['Лист1']

#СОЗДАЕМ СЛОВАРЬ, КУДА ВНОСИМ ОТПРАВИТЕЛЕЙ. ВАЖНО! ПЕРВЫЙ СТОЛБЕЦ: ИМЯ ОТПРАВИТЕЛЕЙ. ВТОРОЙ СТОЛБЕЦ: ЭЛЕКТРОННЫЕ АДРЕСА ОТПРАВИТЕЛЕЙ.
customers = {}
name = 1

for i, x in sheet.values:
     customers[i] = x

#ПРОВЕРКА. УДОБНО РАБОТАЕТ В "НОУТБУКЕ"
#customers

#СОЗДАНИЕ СООБЩЕНИЯ. В ДАННОМ СЛУЧАЕМ ИСПОЛЬЗУЕМ ФОРМАТ HTML.
msg = MIMEMultipart()

message = """\
<html>
  <head></head>
  <body>
    <p>Добрый вечер!<br>
       	Давно выяснено, что при оценке дизайна и композиции читаемый текст мешает сосредоточиться.<br>
       	Lorem Ipsum используют потому, что тот обеспечивает более или менее стандартное заполнение шаблона, а также реальное распределение букв и пробелов в абзацах, которое не получается при простой дубликации "Здесь ваш текст..<br>
       	Здесь ваш текст.. Здесь ваш текст..<br>
       	" Многие программы электронной вёрстки и редакторы HTML используют Lorem Ipsum в качестве текста по умолчанию, так что поиск по ключевым словам "lorem ipsum" сразу показывает, как много веб-страниц всё ещё дожидаются своего настоящего рождения.<br>
       	За прошедшие годы текст Lorem Ipsum получил много версий.<br>
       	Некоторые версии появились по ошибке, некоторые - намеренно (например, юмористические варианты).<br>
        <br>
        Дополнительная информация<br>
        <br>
        Подпись<br>
        Должность<br>
        <br>
        Телефон<br>
        Факсbr>
        <br>
        Какая-либо еще информация<br>
    </p>
  </body>
</html>

"""
 

#НАСТРОЙКИ ОТПРАВКИ СООБЩЕНИЯ
#ПАРОЛЬ ОТ АККАУНТА
password = 'PASS'

#ПОЛНОЕ УКАЗАНИЕ АДРЕСА ЭЛЕКТРОННОЙ ПОЧТЫ
msg['From'] = 'e-mail'

#ТЕМА ПИСЬМА
msg['Subject'] = 'Тема письма'

#ПРИКРЕПЛЕНИЕ ВЛОЖЕНИЯ К ПИСЬМУ
#ИМЯ ВЛОЖЕНИЯ, КОТОРОЕ БУДЕТ ВИДЕТЬ АДРЕСАТ СООБЩЕНИЯ
filename = "special_offer.pdf"

#АДРЕС МЕСТОРАСПОЛОЖЕНИЯ ВЛОЖЕНИЯ
attachment = open("special_offer.pdf", "rb")

#НАСТРОЙКА СООБЩЕНИЯ И ВЛОЖЕНИЙ
part = MIMEBase('application', 'octet-stream')
part.set_payload((attachment).read())
encoders.encode_base64(part)
part.add_header('Content-Disposition', "attachment; filename= %s" % filename)

msg.attach(part)

#ЗАДАЕМ, ЧТО БУДЕТ В ТЕЛЕ СООБЩЕНИЯ И ПЕРЕДАЕМ ФОРМАТ САМОГО СООБЩЕНИЯ
msg.attach(MIMEText(message, 'html'))
 
#СОЗДАЕМ ПОДКЛЮЧЕНИЕ. В ДАННОМ СЛУЧАЕ - ЭТО GOOGLE
server = smtplib.SMTP('smtp.gmail.com: 587')
 
server.starttls()
 
#ЛОГИНИМСЯ ДЛЯ ОТПРАВКИ СООБЩЕНИЯ
server.login(msg['From'], password)
 
#ОТПРАВКА СООБЩЕНИЯ С ПОМОЩЬЮ ПОДКЛЮЧЕННОГО СЕРВЕРА

for name, email in customers.items():
    try:
        server.sendmail(msg['From'], email, msg.as_string())
        print(f'Отправлено на адрес: {(email)}')
    except:
        print(f'--> Ошибка отправки на: {(email)}. Адресат: {(name)}.')
server.quit()
#отправка на 200 адресов, потом идет закрытие соединения