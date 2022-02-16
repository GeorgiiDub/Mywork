from tkinter import *
from tkinter import messagebox as mb
from tkcalendar import DateEntry
from tkinter.ttk import Combobox

''' работа над полями ввода'''

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
var_guest_address = StringVar()
var_id_guest = StringVar()
var_pass_guest = StringVar()
var_link_guest = StringVar()
var_ip_conf = StringVar()
var_id_conf = StringVar()
var_pass_conf = StringVar()
var_link_conf = StringVar()
# var_data_connect = StringVar()
var_vcs = StringVar()
var_meet_room = StringVar()
var_content = StringVar()

times_hour = [x for x in range(24)]
times_minute = [x for x in range(0, 60, 15)]
duration = [x for x in range(1, 5)]


def hour_minute(x, y):
    var_time = x + '-' + y
    return var_time

def write_meet(x, y, z, a):
    y=str(y)
    f = open(y+'_'+a+'_'+z+'_'+'.txt', 'w')
    f.write(x)
    f.close()

def save1():
    var_date = txt_date.get_date()
    var_date=var_date.strftime('%m.%d.%Y')
    var_hour = txt_hour.get()
    var_minute = txt_minute.get()
    var_topic = txt_topic.get()
    var_project = txt_project.get()
    var_duration = txt_duration.get()
    var_name_org = txt_name_org.get()
    var_email = txt_email.get()
    var_telephone = txt_telephone.get()
    var_data_connect = txt_data_connect.get()
    var_vcs = txt_vcs.get()
    var_meet_room = txt_meet_room.get()
    var_content = txt_content.get()

    var_time = hour_minute(var_hour, var_minute)

    result1 =(f'{lbl_date["text"]}: {var_date} {var_time}\n{lbl_topic["text"]}: {var_topic}\n{lbl_project["text"]}: {var_project}\n\
{lbl_duration["text"]}: {var_duration}\n{lbl_name_org["text"]}: {var_name_org}\t{var_email}\t\
{var_telephone}\n{lbl_data_connect["text"]}: {var_data_connect}\n\
{lbl_vcs["text"]}: {var_vcs}\n{lbl_meet_room["text"]}: {var_meet_room}\n{lbl_content["text"]}: {var_content}\n')
    write_meet(result1, var_date, var_topic, var_time)
    mb.showinfo("записалось", result1)

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

'''переименовать данные на подключение входные и выходные добавить поля ссылка, идентификатор, пароль, IP адрес'''

lbl_guest_address = Label(window, text="6. Гостевое подключение IP", relief=GROOVE)
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


lbl_vcs = Label(window, text="7. Система ВКС", relief=GROOVE)
lbl_vcs.grid(column=0, row=10, ipadx=5, ipady=5, sticky=E, padx=3, pady=3)
rb_yms = Radiobutton(window, text='YMS', value = 1)
rb_yms.grid(column=1, row=10, ipadx=5, ipady=5, padx=3, pady=3)
rb_zoom = Radiobutton(window, text='ZOOM', value = 2)
rb_zoom.grid(column=2, row=10, ipadx=5, ipady=5, padx=3, pady=3)


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

'''сделать чек баттон выбора переговорных комнат'''

lbl_meet_room = Label(window, text="8. Переговорные комнаты", relief=GROOVE)
lbl_meet_room.grid(column=0, row=14, ipadx=5, ipady=5, sticky=E, padx=3, pady=3)
chb_meet_room2 = Checkbutton(window, text="С №2")
chb_meet_room2.grid(column=1, row=14, ipadx=5, ipady=5, padx=3, pady=3, sticky=E)
chb_meet_room3 = Checkbutton(window, text="С №3")
chb_meet_room3.grid(column=1, row=14, ipadx=5, ipady=5, padx=3, pady=3, sticky=W)
chb_meet_roomv = Checkbutton(window, text="VIP")
chb_meet_roomv.grid(column=1, row=14)


'''сделать радиобаттон'''

lbl_content = Label(window, text="9. Демонстрация материалов", relief=GROOVE)
lbl_content.grid(column=0, row=15, ipadx=5, ipady=5, sticky=E, padx=3, pady=3)
txt_content = Entry(window, width=30, textvariable=var_content)
txt_content.grid(column=1, row=15, ipadx=5, ipady=5, sticky=W, padx=3, pady=3)



btn_write = Button(window, text="Записать", bg="#abd9ff", command=save1)
btn_write.grid(column=0, row=16, ipadx=5, ipady=5, sticky=E, padx=3, pady=3)


window.event_add('<<Paste>>', '<Control-igrave>')
window.event_add("<<Copy>>", "<Control-ntilde>")
window.mainloop()

