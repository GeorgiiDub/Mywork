from tkinter import *
from tkinter import messagebox as mb
from tkcalendar import DateEntry
from tkinter.ttk import Combobox
import smtplib
from email.mime.text import MIMEText
import sys
import os
from configparser import ConfigParser
import re
import win32com.client

outlook = win32com.client.Dispatch("Outlook.Application")

window = Tk()
window.title("My Work")
window.geometry("520x620")
window['background']='#477670'

# переменные
var_date = StringVar()
var_hour = StringVar()
var_minute = StringVar()
var_topic = StringVar()
var_project = StringVar()
var_duration = StringVar()
var_name_org = StringVar()
var_email = StringVar()
var_telephone = StringVar()
var_contact_org = StringVar()
var_guest_address = StringVar()
var_id_guest = StringVar()
var_pass_guest = StringVar()
var_link_guest = StringVar()
var_vcs_system = StringVar()
var_ip_conf = StringVar()
var_id_conf = StringVar()
var_pass_conf = StringVar()
var_link_conf = StringVar()
var_zoom = StringVar()
var_meet_room = StringVar()
var_content = StringVar()
var_result1 = str("")
var_answer_org = str("")
subject = str("")

times_hour = [x for x in range(24)]
times_minute = [x for x in range(0, 60, 15)]
duration = [x for x in range(00, 180, 30)]

'''поработать с объектом добавить вложения в объект "собрание" или найти другой способ формирования такого объекта 
для всех календарей 

создать отдельным файлом комнаты с сылками и использовать условием выбора идентификатора подстановкой значений в 
переменные 

обработать сбор ошибок отправки в отдельный файл написать проверку заполнения полей и проверки их ввод на 
правильность сделать запуск приложения с сервера, и путь к сетевой папке Z на папку какую-нибудь для сохранения 
файла, с условием сохранения фала в папке программы '''


# функция данные организатора одним блоком
def contacts_org():
    name_org = txt_name_org.get()
    email = txt_email.get()
    telephone = txt_telephone.get()
    if not email:
        mb.showinfo("Почта организатора", "укажите почту организаторa")
    else:
        var_contacr_org = str(name_org + '\n\tэл.почта: ' + email + '\n\tтелефон: ' + telephone)
    return var_contacr_org


# функция объединения данные подключения конференции и вывод H.323 и SIP гостевые
def connect_change_guest():
    guest_address = txt_guest_address.get()
    id_guest: str = txt_id_guest.get()
    pass_guest = txt_pass_guest.get()
    link_guest = txt_link_guest.get()
    id_guest = id_guest.replace(" ", "")
    var_z = str(
        'H.323: ' + guest_address + '##' + id_guest + '\n\tSIP: ' + id_guest + '@' + guest_address + '\n')
    var_connect = str('\n\tСсылка на подключение:\n\t' \
                      + link_guest + '\n\tИдентификатор: ' + id_guest + '\tПароль: ' \
                      + pass_guest + '\n\tАдрес: ' + guest_address + '\n\t' + var_z)
    return var_connect


# функция объединения данные подключения конференции и вывод H.323 и SIP участники компании
def connect_change():
    base_path = os.path.dirname(os.path.abspath(__file__))
    config_path = os.path.join(base_path, "addr.ini")
    if os.path.exists(config_path):
        cfg = ConfigParser()
        cfg.read(config_path)
    else:
        print("Config not found! Exiting!")
        sys.exit(1)
    vcs_system = cb_vcs_system.get()
#   ip_conf = txt_ip_conf.get()
    id_conf = txt_id_conf.get()
    pass_conf = txt_pass_conf.get()
    link_conf = txt_link_conf.get()
    addr_zoom = cfg.get("addr_zoom", "server_zoom")
    addr_yms = cfg.get("addr_yms", "server_yms")
    id_conf = id_conf.replace(" ", "")
    if vcs_system == 'ZOOM':
        var_connect = str('\nСистема ВКС: ' + vcs_system + '\n\tСсылка на подключение:\n\t' + link_conf + \
                          '\n\tИдентификатор: ' + id_conf + '\n\tПароль: ' + pass_conf + \
                          '\n\tПодключение по H.323: ' + addr_zoom)
        return var_connect
    elif vcs_system == 'YMS':
        var_z = str('H.323: ' + addr_yms + '##' + id_conf + '\n\tSIP: ' + id_conf + '@' + addr_yms + '\n')
        var_connect = str('\nСистема ВКС: ' + vcs_system + '\n\tСсылка на подключение:\n\t' + link_conf + \
                          '\n\tИдентификатор: ' + id_conf + '\tПароль: ' + pass_conf + \
                          '\n\tАдрес: ' + addr_yms + '\n\t' + var_z)
        return var_connect
    else:
        var_z = str('H.323: ' + vcs_system + '##' + id_conf + '\n\tSIP: ' + id_conf + '@' + vcs_system + '\n')
        var_connect = str('\nСистема ВКС: ' + vcs_system + '\n\tСсылка на подключение:\n\t' + link_conf + \
                          '\n\tИдентификатор: ' + id_conf + '\tПароль: ' + pass_conf + \
                          '\n\tАдрес: ' + vcs_system + '\n\t' + var_z)
        return var_connect

# def chek_input(a):


# функция дата+время шаблон "yyyy-MM-dd hh:mm"
def date_time():
    var_date = txt_date.get_date()
    var_date = var_date.strftime('%d.%m.%Y')
    var_hour = txt_hour.get()
    var_minute = txt_minute.get()
    if not var_hour:
        mb.showinfo("Некорректное время", "выберите часы")
    elif not var_minute:
        mb.showinfo("Некорректное время", "выберите минуты")
    else:
        var_datetime = str(var_date + ' ' + var_hour + ':' + var_minute)
    return var_datetime


# функция записи в файл всего совещания
def write_meet(a, b, c):  # var_result1, var_datetime, var_topic
    global subject
    # удаляет все не буквенно-цифровые символы
    e = re.sub(r'[\W_]+', '_', c)
    subject = str(b + '_' + '_' + e)
    subject = subject.replace(":", "-")
    f = open(subject + '.txt', 'w')
    f.write(a)
    f.close()


# функция результирующей строки
def save1():
    var_datetime = date_time()
    topic = txt_topic.get()
    project = txt_project.get()
    duration = txt_duration.get()
    contact_org = contacts_org()
    var_connect_guest = connect_change_guest()
    var_connect_conf = connect_change()
    meet_room = txt_meet_room.get()
    content = cb_content.get()
    if not topic:
        mb.showinfo("Тема", "Напишите тему совещания")
    elif not duration:
        mb.showinfo("Продолжительность", "выберите продолжительность совещания")
    else:
        global var_answer_org
        var_answer_org = str(f'Дата и время совещания: {var_datetime}\n'
                         f'Тема совещания: {topic}\n'
                         f'Подключение участников компании: {var_connect_conf}\n')
        global var_result1
        var_result1 = str(f'Дата и время совещания: {var_datetime}'
                      f'\nТема совещания: {topic}'
                      f'\nПроект: {project}'
                      f'\nПродолжительность: {duration} мин.'
                      f'\nЗаказчик-организатор: {contact_org}'
                      f'\nГостевое подключение: {var_connect_guest}'
                      f'\nПодключение участников компании: {var_connect_conf}'
                      f'\nПереговорные комнаты: {meet_room}'
                      f'\nДемонстрация материалов: {content}'
                      f'\n\nОтвет организатору:\n{var_answer_org}')
        write_meet(var_result1, var_datetime, topic)
        mb.showinfo("записалось", var_result1)


# функция отправки ответа организатору
def send_answer_org():
    base_path = os.path.dirname(os.path.abspath(__file__))
    config_path = os.path.join(base_path, "email.ini")
    if os.path.exists(config_path):
        cfg = ConfigParser()
        cfg.read(config_path)
    else:
        print("Config not found! Exiting!")
        sys.exit(1)
    host = cfg.get("smtp", "server")
    from_addr = cfg.get("smtp", "from_addr")
    password = cfg.get("smtp", "password")
    global var_answer_org
    email = txt_email.get()
    smtpObj = smtplib.SMTP(host, 587)
    smtpObj.starttls()
    smtpObj.login(from_addr, password)
    msg_send = MIMEText(var_answer_org, 'plain', 'utf-8')
    smtpObj.sendmail(from_addr, email, msg_send.as_string())
    smtpObj.quit()
    mb.showinfo("Ответ", var_answer_org)


# функция отправки на электронную почту админу
def send_meet():
    base_path = os.path.dirname(os.path.abspath(__file__))
    config_path = os.path.join(base_path, "email.ini")
    if os.path.exists(config_path):
        cfg = ConfigParser()
        cfg.read(config_path)
    else:
        print("Config not found! Exiting!")
        sys.exit(1)
    host = cfg.get("smtp", "server")
    from_addr = cfg.get("smtp", "from_addr")
    password = cfg.get("smtp", "password")
    addr_adm = cfg.get("smtp", "addr_adm")
    global var_result1
    global subject
    charset = 'Content-Type: text/plain; charset=utf-8'
    mime = 'MIME-Version: 1.0'
    body = "\r\n".join((f"From: {from_addr}", f"To: {addr_adm}",
                        f"Subject: {subject}", mime, charset, "", var_result1))
    smtpObj = smtplib.SMTP(host, 587)
    smtpObj.starttls()
    smtpObj.login(from_addr, password)
    smtpObj.sendmail(from_addr, addr_adm, body.encode('utf-8'))
    smtpObj.quit()


# функция создания объекта .com тип "собрание" в outlook
def creat_meeting():
    topic = txt_topic.get()
    meet_room = txt_meet_room.get()
    var_datetime = date_time()
    duration = int(txt_duration.get())
    global var_result1
    base_path = os.path.dirname(os.path.abspath(__file__))
    config_path = os.path.join(base_path, "email.ini")
    if os.path.exists(config_path):
        cfg = ConfigParser()
        cfg.read(config_path)
    else:
        print("Config not found! Exiting!")
        sys.exit(1)
    addr_adm = cfg.get("smtp", "addr_adm")
    appt = outlook.CreateItem(1)
    appt.Start = var_datetime  # '2022-04-19 08:00'
    appt.Subject = topic
    appt.Duration = duration
    appt.Location = meet_room
    appt.MeetingStatus = 1
    appt.Recipients.Add(addr_adm)
    # Установите Pattern, чтобы он повторялся каждый день в течение следующих 5 дней
    # pattern = appt.GetRecurrencePattern()
    # pattern.RecurrenceType = 0
    # pattern.Occurrences = "5"
    # appt.Save()
    appt.Send()


# функиця очистки формы
def clear_form():
    txt_date.delete("0", END)
    txt_hour.delete("0", END)
    txt_minute.delete("0", END)
    txt_topic.delete("0", END)
    txt_project.delete("0", END)
    txt_duration.delete("0", END)
    txt_name_org.delete("0", END)
    txt_email.delete("0", END)
    txt_telephone.delete("0", END)
    txt_guest_address.delete("0", END)
    txt_id_guest.delete("0", END)
    txt_pass_guest.delete("0", END)
    txt_link_guest.delete("0", END)
    cb_vcs_system.delete("0", END)
#   txt_ip_conf.delete("0", END)
    txt_id_conf.delete("0", END)
    txt_pass_conf.delete("0", END)
    txt_link_conf.delete("0", END)
    txt_meet_room.delete("0", END)
    cb_content.delete("0", END)


# интерфейс
lbl_date = Label(window, text="1. Дата и время", font=("Arial", 12))
lbl_date.place(x=20, y=20)
txt_date = DateEntry(window, width=12, textvariable=var_date, date_pattern='dd/mm/yy')
txt_date.place(x=230, y=20)
txt_hour = Combobox(window, values=times_hour, width=5)
txt_hour.place(x=334, y=20)
txt_minute = Combobox(window, values=times_minute, width=5)
txt_minute.place(x=395, y=20)

lbl_topic = Label(window, text="2. Тема совещания", font=("Arial", 12))
lbl_topic.place(x=20, y=50)
txt_topic = Entry(window, textvariable=var_topic)
txt_topic.place(x=230, y=50, width=270)

lbl_project = Label(window, text="3. Проект", font=("Arial", 12))
lbl_project.place(x=20, y=80)
txt_project = Entry(window, textvariable=var_project)
txt_project.place(x=230, y=80, width=156)

lbl_duration = Label(window, text="4. Продолжительность", font=("Arial", 12))
lbl_duration.place(x=20, y=110)
txt_duration = Combobox(window, values=duration, width=5)
txt_duration.place(x=230, y=110)

lbl_name_org = Label(window, text="5. Организатор", font=("Arial", 12))
lbl_name_org.place(x=20, y=140)
txt_name_org = Entry(window, width=28, textvariable=var_name_org)
txt_name_org.place(x=40, y=170)
lbl_email = Label(window, text="Электронная почта", font=("Arial", 12))
lbl_email.place(x=230, y=140)
txt_email = Entry(window, width=27, textvariable=var_email)
txt_email.place(x=230, y=170)
lbl_telephone = Label(window, text="Телефон", font=("Arial", 12))
lbl_telephone.place(x=410, y=140)
txt_telephone = Entry(window, textvariable=var_telephone)
txt_telephone.place(x=410, y=170, width=90)

lbl_guest_address = Label(window, text="6. Гостевое подключение", font=("Arial", 12))
lbl_guest_address.place(x=20, y=200)
txt_guest_address = Entry(window, width=28, textvariable=var_guest_address)
txt_guest_address.place(x=40, y=230)

lbl_id_guest = Label(window, text="Идентификатор", font=("Arial", 12))
lbl_id_guest.place(x=230, y=200)
txt_id_guest = Entry(window, textvariable=var_id_guest)
txt_id_guest.place(x=230, y=230, width=130)

lbl_pass_guest = Label(window, text="Пароль", font=("Arial", 12))
lbl_pass_guest.place(x=380, y=200)
txt_pass_guest = Entry(window, width=30, textvariable=var_pass_guest)
txt_pass_guest.place(x=380, y=230, width=120)

lbl_link_guest = Label(window, text="Ссылка подключения", font=("Arial", 12))
lbl_link_guest.place(x=40, y=260)
txt_link_guest = Entry(window, textvariable=var_link_guest)
txt_link_guest.place(x=40, y=290, width=460)

lbl_ip_conf = Label(window, text="7. Подключение компании", font=("Arial", 12))
lbl_ip_conf.place(x=20, y=320)
cb_vcs_system = Combobox(window, textvariable=var_vcs_system, values=['YMS', 'ZOOM', 'SKYPE', 'MS Teams', 'Другая'])
cb_vcs_system.place(x=40, y=355, width=170)

lbl_id_conf = Label(window, text="Идентификатор", font=("Arial", 12))
lbl_id_conf.place(x=230, y=320)
txt_id_conf = Entry(window, textvariable=var_id_conf)
txt_id_conf.place(x=230, y=355, width=130)

lbl_pass_conf = Label(window, text="Пароль", font=("Arial", 12))
lbl_pass_conf.place(x=380, y=320)
txt_pass_conf = Entry(window, textvariable=var_pass_conf)
txt_pass_conf.place(x=380, y=355, width=120)

lbl_link_conf = Label(window, text="Ссылка подключения", font=("Arial", 12))
lbl_link_conf.place(x=40, y=388)
txt_link_conf = Entry(window, textvariable=var_link_conf)
txt_link_conf.place(x=40, y=422, width=460,)

lbl_meet_room = Label(window, text="8. Переговорные комнаты", font=("Arial", 12))
lbl_meet_room.place(x=20, y=450)
txt_meet_room = Entry(window, textvariable=var_meet_room)
txt_meet_room.place(x=230, y=450, width=270)

lbl_content = Label(window, text="9. Демонстрация материалов", font=("Arial", 12))
lbl_content.place(x=20, y=485)
cb_content = Combobox(window, values=['ДА', 'НЕТ'], width=5)
cb_content.place(x=260, y=485)

btn_write = Button(window, text="Записать", bg="#abd9ff", command=save1, font=("Arial", 12))
btn_write.place(x=20, y=530)
btn_clear = Button(window, text="Очистить", bg="#abd9ff", command=clear_form, font=("Arial", 12))
btn_clear.place(x=20, y=570)
btn_clear = Button(window, text="Отправить админу", bg="#abd9ff", command=send_meet, font=("Arial", 12))
btn_clear.place(x=120, y=530)
btn_clear = Button(window, text="Создать собрание", bg="#abd9ff", command=creat_meeting, font=("Arial", 12))
btn_clear.place(x=120, y=570)
btn_clear = Button(window, text="Отправить организатору", bg="#abd9ff", command=send_answer_org, font=("Arial", 12))
btn_clear.place(x=280, y=530)

window.event_add('<<Paste>>', '<Control-igrave>')
window.event_add("<<Copy>>", "<Control-ntilde>")
window.mainloop()
