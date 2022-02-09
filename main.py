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
var_contacts = StringVar()
var_data_connect = StringVar()
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
    var_contacts = txt_contacts.get()
    var_data_connect = txt_data_connect.get()
    var_vcs = txt_vcs.get()
    var_meet_room = txt_meet_room.get()
    var_content = txt_content.get()

    var_time = hour_minute(var_hour, var_minute)

    result1 =(f'{lbl_date["text"]}: {var_date} {var_time}\n{lbl_topic["text"]}: {var_topic}\n{lbl_project["text"]}: {var_project}\n\
{lbl_duration["text"]}: {var_duration}\n{lbl_contacts["text"]}: {var_contacts}\n{lbl_data_connect["text"]}: {var_data_connect}\n\
{lbl_vcs["text"]}: {var_vcs}\n{lbl_meet_room["text"]}: {var_meet_room}\n{lbl_content["text"]}: {var_content}\n')
    write_meet(result1, var_date, var_topic, var_time)
    mb.showinfo("записалось", result1)

lbl_date = Label(window, text="1. Дата и время", relief=GROOVE)
lbl_date.grid(column=0, row=1, ipadx=5, ipady=5, sticky=E, padx=3, pady=3)
txt_date = DateEntry(window, width=12, textvariable=var_date, date_pattern='dd/mm/yy')
txt_date.grid(column=1, row=1, ipadx=5, ipady=5, sticky=W, padx=3, pady=3)

txt_hour = Combobox(window, values=times_hour, width=5)
txt_hour.grid(column=3, row=1, ipadx=5, ipady=5, sticky=W, padx=3, pady=3)
txt_minute = Combobox(window, values=times_minute, width=5)
txt_minute.grid(column=4, row=1, ipadx=5, ipady=5, sticky=W, padx=3, pady=3)

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

'''Добавить контакты организатора приславшего данные, контакты нашего организатора'''

lbl_contacts = Label(window, text="5. Контакты организатора", relief=GROOVE)
lbl_contacts.grid(column=0, row=5, ipadx=5, ipady=5, sticky=E, padx=3, pady=3)
txt_contacts = Entry(window, width=30, textvariable=var_contacts)
txt_contacts.grid(column=1, row=5, ipadx=5, ipady=5, sticky=W, padx=3, pady=3)

'''переименовать данные на подключение гостевые'''

lbl_data_connect = Label(window, text="6. Данные на покдлючение входные", relief=GROOVE)
lbl_data_connect.grid(column=0, row=6, ipadx=5, ipady=5, sticky=E, padx=3, pady=3)
txt_data_connect = Entry(window, width=30, textvariable=var_data_connect)
txt_data_connect.grid(column=1, row=6, ipadx=5, ipady=5, sticky=W, padx=3, pady=3)

'''сделать выбор радио баттон'''

lbl_vcs = Label(window, text="7. Система ВКС", relief=GROOVE)
lbl_vcs.grid(column=0, row=7, ipadx=5, ipady=5, sticky=E, padx=3, pady=3)
txt_vcs = Entry(window, width=30, textvariable=var_vcs)
txt_vcs.grid(column=1, row=7, ipadx=5, ipady=5, sticky=W, padx=3, pady=3)

''' добавить поля нашего подключения к нашей системе ВКС, может быть добавить рядом с гостевыми приглашениями,
Поля добавить отдельно сссылка, идентификатор, пароль для конференции'''

lbl_meet_room = Label(window, text="8. Переговорные комнаты", relief=GROOVE)
lbl_meet_room.grid(column=0, row=8, ipadx=5, ipady=5, sticky=E, padx=3, pady=3)
txt_meet_room = Entry(window, width=30, textvariable=var_meet_room)
txt_meet_room.grid(column=1, row=8, ipadx=5, ipady=5, sticky=W, padx=3, pady=3)

lbl_content = Label(window, text="9. Демонстрация материалов", relief=GROOVE)
lbl_content.grid(column=0, row=9, ipadx=5, ipady=5, sticky=E, padx=3, pady=3)
txt_content = Entry(window, width=30, textvariable=var_content)
txt_content.grid(column=1, row=9, ipadx=5, ipady=5, sticky=W, padx=3, pady=3)



btn_write = Button(window, text="Записать", bg="#abd9ff", command=save1)
btn_write.grid(column=0, row=10, ipadx=5, ipady=5, sticky=E, padx=3, pady=3)


window.event_add('<<Paste>>', '<Control-igrave>')
window.event_add("<<Copy>>", "<Control-ntilde>")
window.mainloop()

