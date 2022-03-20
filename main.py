from tkinter import *
from tkinter import messagebox as mb
from tkcalendar import DateEntry
from tkinter.ttk import Combobox
import smtplib
from email.mime.text import MIMEText

window = Tk()
window.title("My Work")
window.geometry("1200x600")

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

times_hour = [x for x in range(24)]
times_minute = [x for x in range(0, 60, 15)]
duration = [x for x in range(1, 5)]

'''сделать подключение почты по IMAP для сохранения писем в исходящих, обработать сбор ошибок отправки в отдельный файл'''

# функция данные организатора одним блоком
def contacts_org(a,b,c):
    var_contacr_org = str(a+'\n\tэл.почта: '+b+'\n\tтелефон: '+c)
    return var_contacr_org

# функция объединения данные подключения конференции и вывод H.323 и SIP
def connect_change(a,b,c,d,e):
    var_connect = str("")
    c = c.replace(" ", "")
    if a == 'ZOOM':
        var_z = str('SIP: '+c+'@zoomcrc.com')
        var_connect = str('\n\tСсылка на подключение:\n\t' + e + '\n\tИдентификатор: ' + c + '\tПароль: ' + d + '\n\t' + var_z)
        return var_connect
    else:
        var_z = str('H.323: '+b+'##'+c+'\n\tSIP: '+c+'@'+b+'\n')
        var_connect=str('\n\tСсылка на подключение:\n\t'+e+'\n\tИдентификатор: '+c+'\tПароль: '+d+'\n\tАдрес: '+b+'\n\t'+var_z)
        return var_connect

# функция вывода часы-минуты одной строкой
def hour_minute(x, y):
    var_time = x + '-' + y
    return var_time

# функция записи в файл всего совещания
def write_meet(x, y, z, a):
    y=str(y)
    f = open(y+'_'+a+'_'+z+'_'+'.txt', 'w')
    f.write(x)
    f.close()

# функция результирующей строки
def save1():
    var_date = txt_date.get_date()
    var_date=var_date.strftime('%d.%m.%Y')
    var_hour = txt_hour.get()
    var_minute = txt_minute.get()
    var_time = hour_minute(var_hour, var_minute)
    var_topic = txt_topic.get()
    var_project = txt_project.get()
    var_duration = txt_duration.get()
    var_name_org = txt_name_org.get()
    var_email = txt_email.get()
    var_telephone = txt_telephone.get()
    var_contact_org = contacts_org(var_name_org,var_email,var_telephone)
    var_guest_address = txt_guest_address.get()
    var_id_guest = txt_id_guest.get()
    var_pass_guest = txt_pass_guest.get()
    var_link_guest = txt_link_guest.get()
    var_vcs_system = cb_vcs_system.get()
    var_connect_guest = connect_change(var_vcs_system,var_guest_address,var_id_guest,var_pass_guest,var_link_guest)
    var_ip_conf = txt_ip_conf.get()
    var_id_conf = txt_id_conf.get()
    var_pass_conf = txt_pass_conf.get()
    var_link_conf = txt_link_conf.get()
    var_connect_conf = connect_change(var_vcs_system,var_ip_conf,var_id_conf,var_pass_conf,var_link_conf)
    var_meet_room = txt_meet_room.get()
    var_content = cb_content.get()

    global var_result1
    var_result1 =str (f'Дата и время совещания: {var_date} {var_time}\n'
                      f'Тема совещания: {var_topic}\n'
                      f'Проект: {var_project}\n'
                      f'Продолжительность: {var_duration} ч.\n'
                      f'Заказчик-организатор: {var_contact_org}\n'
                      f'Гостевое подключение: {var_connect_guest}\n'
                      f'Система ВКС: {var_vcs_system}\n'
                      f'Подключение участников компании: {var_connect_conf}\n'
                      f'Переговорные комнаты: {var_meet_room}\n'
                      f'Демонстрация материалов: {var_content}\n')

    global var_answer_org
    var_answer_org = str (f'Дата и время совещания: {var_date} {var_time}\n'
                          f'Тема совещания: {var_topic}\n'
                          f'Система ВКС: {var_vcs_system}\n'
                          f'Подключение участников компании: {var_connect_conf}\n')

    write_meet(var_result1, var_date, var_topic, var_time)
    mb.showinfo("записалось", var_result1)


# функция отправки ответа организатору
def send_answer_org():
    global var_answer_org
    var_email = txt_email.get()
    smtpObj = smtplib.SMTP('', 587)
    smtpObj.starttls()
    smtpObj.login('', '')
    msg_send = MIMEText(var_answer_org, 'plain', 'utf-8')
    smtpObj.sendmail("", var_email, msg_send.as_string())
    smtpObj.quit()
    mb.showinfo("Ответ", var_answer_org)

# функция отправки на электронную почту админу
def send_meet():
    global var_result1
    smtpObj = smtplib.SMTP('smtp.mail.ru', 587)
    smtpObj.starttls()
    smtpObj.login('', '')
    msg_send = MIMEText(var_result1, 'plain', 'utf-8')
    smtpObj.sendmail("", "", msg_send.as_string())
    smtpObj.quit()

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
    txt_ip_conf.delete("0", END)
    txt_id_conf.delete("0", END)
    txt_pass_conf.delete("0", END)
    txt_link_conf.delete("0", END)
    txt_meet_room.delete("0", END)
    cb_content.delete("0", END)


lbl_date = Label(window, text="1. Дата и время", relief=GROOVE)
lbl_date.grid(column=0, row=1, ipadx=5, ipady=5, sticky=E, padx=3, pady=3)
txt_date = DateEntry(window, width=12, textvariable=var_date, date_pattern='dd/mm/yy')
txt_date.grid(column=1, row=1, ipadx=5, ipady=5, sticky=W, padx=3, pady=3)

txt_hour = Combobox(window, values=times_hour, width=5)
txt_hour.grid(column=2, row=1, ipadx=5, ipady=5, sticky=E, padx=3, pady=3)
txt_minute = Combobox(window, values=times_minute, width=5)
txt_minute.grid(column=3, row=1, ipadx=5, ipady=5, sticky=W, padx=3, pady=3)

lbl_topic = Label(window, text="2. Тема совещания", relief=GROOVE)
lbl_topic.grid(column=0, row=2, ipadx=5, ipady=5, sticky=E, padx=3, pady=3)
txt_topic = Entry(window, width=30, textvariable=var_topic)
txt_topic.grid(column=1, row=2, ipadx=5, ipady=5, sticky=W, padx=3, pady=3)

lbl_project = Label(window, text="3. Проект", relief=GROOVE)
lbl_project.grid(column=0, row=3, ipadx=5, ipady=5, sticky=E, padx=3, pady=3)
txt_project = Entry(window, width=30, textvariable=var_project)
txt_project.grid(column=1, row=3, ipadx=5, ipady=5, sticky=W, padx=3, pady=3)

lbl_duration = Label(window, text="4. Продолжительность", relief=GROOVE)
lbl_duration.grid(column=0, row=4, ipadx=5, ipady=5, sticky=E, padx=3, pady=3)
txt_duration = Combobox(window, values=duration, width=5)
txt_duration.grid(column=1, row=4, ipadx=5, ipady=5, sticky=W, padx=3, pady=3)

lbl_name_org = Label(window, text="5. Организатор", relief=GROOVE)
lbl_name_org.grid(column=0, row=5, ipadx=5, ipady=5, padx=3, pady=3, sticky=E)
txt_name_org = Entry(window, width=30, textvariable=var_name_org)
txt_name_org.grid(column=0, row=6, ipadx=5, ipady=5, padx=3, pady=3, sticky=E)

lbl_email = Label(window, text="Электронная почта", relief=GROOVE)
lbl_email.grid(column=1, row=5, ipadx=5, ipady=5, padx=3, pady=3)
txt_email = Entry(window, width=30, textvariable=var_email)
txt_email.grid(column=1, row=6, ipadx=5, ipady=5, padx=3, pady=3)

lbl_telephone = Label(window, text="Телефон", relief=GROOVE)
lbl_telephone.grid(column=2, row=5, ipadx=5, ipady=5, padx=3, pady=3)
txt_telephone = Entry(window, width=30, textvariable=var_telephone)
txt_telephone.grid(column=2, row=6, ipadx=5, ipady=5, padx=3, pady=3)

lbl_guest_address = Label(window, text="6. Гостевое подключение", relief=GROOVE)
lbl_guest_address.grid(column=0, row=7, ipadx=5, ipady=5, sticky=E, padx=3, pady=3)
txt_guest_address = Entry(window, width=30, textvariable=var_guest_address)
txt_guest_address.grid(column=0, row=8, ipadx=5, ipady=5, padx=3, pady=3, sticky=E)

lbl_id_guest = Label(window, text="Идентификатор", relief=GROOVE)
lbl_id_guest.grid(column=1, row=7, ipadx=5, ipady=5, padx=3, pady=3)
txt_id_guest = Entry(window, width=30, textvariable=var_id_guest)
txt_id_guest.grid(column=1, row=8, ipadx=5, ipady=5, padx=3, pady=3)

lbl_pass_guest = Label(window, text="Пароль", relief=GROOVE)
lbl_pass_guest.grid(column=2, row=7, ipadx=5, ipady=5, padx=3, pady=3)
txt_pass_guest = Entry(window, width=30, textvariable=var_pass_guest)
txt_pass_guest.grid(column=2, row=8, ipadx=5, ipady=5, padx=3, pady=3)

lbl_link_guest = Label(window, text="Ссылка подключения", relief=GROOVE)
lbl_link_guest.grid(column=0, row=9, ipadx=5, ipady=5, padx=3, pady=3, sticky=E)
txt_link_guest = Entry(window, width=50, textvariable=var_link_guest)
txt_link_guest.grid(column=1, row=9, ipadx=5, ipady=5, padx=3, pady=3, columnspan=2, sticky=EW)

lbl_vcs_system = Label(window, text="7. Система ВКС", relief=GROOVE)
lbl_vcs_system.grid(column=0, row=10, ipadx=5, ipady=5, sticky=E, padx=3, pady=3)
cb_vcs_system = Combobox(window, textvariable=var_vcs_system, values=['YMS','ZOOM','SKYPE','MS Teams'], width=5)
cb_vcs_system.grid(column=1, row=10, ipadx=5, ipady=5, sticky=W, padx=3, pady=3)

'''сделать подстановку из выбора уже готовых шаблонов через условиче "Если YMS" '''

lbl_ip_conf = Label(window, text="8. Подключение участников компании", relief=GROOVE)
lbl_ip_conf.grid(column=0, row=11, ipadx=5, ipady=5, sticky=E, padx=3, pady=3)
txt_ip_conf = Entry(window, width=30, textvariable=var_ip_conf)
txt_ip_conf.grid(column=0, row=12, ipadx=5, ipady=5, padx=3, pady=3, sticky=E)

lbl_id_conf = Label(window, text="Идентификатор", relief=GROOVE)
lbl_id_conf.grid(column=1, row=11, ipadx=5, ipady=5, padx=3, pady=3)
txt_id_conf = Entry(window, width=30, textvariable=var_id_conf)
txt_id_conf.grid(column=1, row=12, ipadx=5, ipady=5, padx=3, pady=3)

lbl_pass_conf = Label(window, text="Пароль", relief=GROOVE)
lbl_pass_conf.grid(column=2, row=11, ipadx=5, ipady=5, padx=3, pady=3)
txt_pass_conf = Entry(window, width=30, textvariable=var_pass_conf)
txt_pass_conf.grid(column=2, row=12, ipadx=5, ipady=5, padx=3, pady=3)

lbl_link_conf = Label(window, text="Ссылка подключения", relief=GROOVE)
lbl_link_conf.grid(column=0, row=13, ipadx=5, ipady=5, padx=3, pady=3, sticky=E)
txt_link_conf = Entry(window, width=50, textvariable=var_link_conf)
txt_link_conf.grid(column=1, row=13, ipadx=5, ipady=5, padx=3, pady=3, columnspan=2, sticky=EW)

lbl_meet_room = Label(window, text="9. Переговорные комнаты", relief=GROOVE)
lbl_meet_room.grid(column=0, row=14, ipadx=5, ipady=5, sticky=E, padx=3, pady=3)
txt_meet_room = Entry(window, width=50, textvariable=var_meet_room)
txt_meet_room.grid(column=1, row=14, ipadx=5, ipady=5, padx=3, pady=3, columnspan=2, sticky=EW)

lbl_content = Label(window, text="10. Демонстрация материалов", relief=GROOVE)
lbl_content.grid(column=0, row=15, ipadx=5, ipady=5, sticky=E, padx=3, pady=3)
cb_content = Combobox(window, values=['ДА','НЕТ'], width=5)
cb_content.grid(column=1, row=15, ipadx=5, ipady=5, sticky=W, padx=3, pady=3)

btn_write = Button(window, text="Записать", bg="#abd9ff", command=save1)
btn_write.grid(column=0, row=16, ipadx=5, ipady=5, sticky=E, padx=3, pady=3)

btn_clear = Button(window, text="Очистить форму", bg="#abd9ff", command=clear_form)
btn_clear.grid(column=1, row=16, ipadx=5, ipady=5, padx=3, pady=3)
btn_clear = Button(window, text="Отправить email админу", bg="#abd9ff", command=send_meet)
btn_clear.grid(column=2, row=16, ipadx=5, ipady=5, padx=3, pady=3)
btn_clear = Button(window, text="Отправить ответ организатору", bg="#abd9ff", command=send_answer_org)
btn_clear.grid(column=3, row=16, ipadx=5, ipady=5, padx=3, pady=3)

window.event_add('<<Paste>>', '<Control-igrave>')
window.event_add("<<Copy>>", "<Control-ntilde>")
window.mainloop()

