
from tkinter import *
from tkinter.ttk import Combobox
from tkinter import filedialog as fd
from tkinter.messagebox import showerror, showinfo

import tk as tk
from openpyxl import load_workbook
import os
import sys
from docxtpl import DocxTemplate
from petrovich.main import Petrovich
from petrovich.enums import Case, Gender
from docx import settings
from docx import Document
import shutil



lines = []

if not os.path.exists("folder_save.txt"):
    open("folder_save.txt", 'w').close()


if not os.path.exists("information.txt"):
    open("information.txt", 'w').close()



if os.stat("folder_save.txt").st_size == 0:
    file = open("folder_save.txt", "w")
    file.write(" ")
    file.close()

with open(r"folder_save.txt", "r") as file:
    for line in file:
        file_directory = line
file.close()



p = Petrovich()
doc = DocxTemplate('Шаблон.docx')

def insert_text():
    global file_directory
    global path_save_label
    file_directory = fd.askdirectory()
    path_save_label.destroy()

    path_save_label = Label(top, text=f"{file_directory}", font='Times 16', wraplength=450)
    path_save_label.place(x=220, y=400)

    file = open("folder_save.txt", "w")
    file.write(f"{file_directory}")
    file.close()






def uvd_writer(str):
    if str == 'Зареченский':
        strUVD = "«Зареченский» УМВД России по г. Туле"

        return strUVD
    elif str == 'Ильинское':
        strUVD = "«Ильинское» УМВД России по г. Туле"
        return strUVD

    elif str == 'Косогорское':
        strUVD = "«Косогорское» УМВД России по г. Туле"
        return strUVD

    elif str == 'Криволученский':
        strUVD = "«Криволученский» УМВД России по г. Туле"
        return strUVD

    elif str == 'Ленинский':
        strUVD = "«Ленинский» УМВД России по г. Туле"
        return strUVD

    elif str == 'Привокзальный':
        strUVD = "«Привокзальный» УМВД России по г. Туле"
        return strUVD

    elif str == 'УМВД':
        strUVD = "УМВД России по г. Туле"
        return strUVD

    elif str == 'Скуратовский':
        strUVD = "«Скуратовский» УМВД России по г. Туле"
        return strUVD

    elif str == 'Советский':
        strUVD = "«Советский» УМВД России по г. Туле"
        return strUVD

    elif str == 'Центральный':
        strUVD = "«Центральный» УМВД России по г. Туле"
        return strUVD


def print_doc():
    global lines
    global lname
    global fname
    global mname
    global birthday
    global address
    global alert
    global uvd
    # lines = [lname, fname, mname, birthday, address, alert, uvd]

    if os.stat("folder_save.txt").st_size == 0 or file_directory == " ":
        showerror(title="Ошибка", message="Не указан путь сохранения !")
        return

    if os.stat("information.txt").st_size == 0:
        lname = lname_entry.get()
        fname = fname_entry.get()
        mname = mname_entry.get()
        birthday = birthday_entry.get()
        address = address_entry.get()
        alert = alert_entry.get()
        uvd = combo_UVD.get()
        lines = [lname, fname, mname, birthday, address, alert, uvd]
        with open(r"information.txt", "w") as file:
                file.write(f"""{lines[0]}
{lines[1]}
{lines[2]}
{lines[3]}
{lines[4]}
{lines[5]}
{lines[6]}""")

    else:
        with open(r"information.txt", "r") as file:
            for line in file:
                lines.append(line)
        lname = lines[0][:-2]
        fname = lines[1][:-2]
        mname = lines[2][:-2]
        birthday = lines[3][:-2]
        address = lines[4][:-2]
        alert = lines[5][:-2]
        uvd = lines[6]

    strUVD = uvd_writer(uvd)

    if lname == "" or fname == "" or mname == "" or birthday == "" or address == "" or alert == "":
        showerror(title="Ошибка", message="Заполните все поля!")
        return


  


    cased_lastname = p.lastname(lname, Case.GENITIVE, Gender.MALE)
    cased_firstname = p.firstname(fname, Case.GENITIVE, Gender.MALE)
    cased_middlename = p.middlename(mname, Case.GENITIVE, Gender.MALE)
    cased_last_first_middle = f"{cased_lastname} {cased_firstname} {cased_middlename}"


    abbreviated_lastname = f"{lname} {fname[0].upper()}.{mname[0].upper()}."
    cased_abbreviated_lastname = f"{cased_lastname} {fname[0].upper()}.{mname[0].upper()}."

    pathDoc = f'{abbreviated_lastname}' + '.docx'

    contex = {'strUVD': strUVD,
                'cased_last_first_middle': cased_last_first_middle,
                'abbreviated_lastname': abbreviated_lastname,
                'cased_abbreviated_lastname': cased_abbreviated_lastname,
                'strData': birthday,
                'place_registration': address,
                'notification_date': alert
                }
    doc.render(contex)
    try:
        doc.save(f"{file_directory}/{pathDoc}")
    except:
        print("!!!!!!!")
    if os.path.exists(f"{file_directory}/{pathDoc}"):
        showinfo(title="Информация", message="Документ успешно сохранен!")

    
top = Tk()
top.geometry("640x640")


# folder = tk.StringVar() # Receiving user's folder_name selection

with open(r"information.txt", "r") as file:
    for line in file:
        lines.append(line)


sv = StringVar()

def callback():
    if os.stat("information.txt").st_size != 0:
        lines1 = []
        with open("information.txt", 'r') as f:
            lines1 = f.readlines()
        text = str(sv.get())
        lines1[0] = text

        i = 1
        while i < len(lines1):
            lines1[i] = lines1[i][:-1]
            if i == 5:
                break
            i += 1

        with open(r"information.txt", "w") as file:
            for line in lines1:
                file.write(line + '\n')
        return True
    else:
        return


name_label = Label(top, text="Фамилия", font='Times 18').place(x=30, y=50)
lname_entry = Entry(top, width=30, font='Times 18', textvariable=sv, validate="focusout", validatecommand=callback)
lname_entry.place(x=220, y=50)
if os.stat("information.txt").st_size != 0:
    lname_entry.insert(0, str(lines[0]))

first_name_label = Label(top, text="Имя", font='Times 18').place(x=30, y=100)
fname_entry = Entry(top, width=30, font='Times 18')
fname_entry.place(x=220, y=100)
if os.stat("information.txt").st_size != 0:
    fname_entry.insert(0, str(lines[1]))

middle_name_label = Label(top, text="Отчество", font='Times 18')
middle_name_label.place(x=30, y=150)
mname_entry = Entry(top, width=30, font='Times 18')
mname_entry.place(x=220, y=150)
if os.stat("information.txt").st_size != 0:
    mname_entry.insert(0, str(lines[2]))

birthday_label = Label(top, text="Дата рождения", font='Times 18')
birthday_label.place(x=30, y=200)
birthday_entry = Entry(top, width=30, font='Times 18')
birthday_entry.place(x=220, y=200)
if os.stat("information.txt").st_size != 0:
    birthday_entry.insert(0, str(lines[3]))

address_label = Label(top, text="Адрес", font='Times 18')
address_label.place(x=30, y=250)
address_entry = Entry(top, width=30, font='Times 18')
address_entry.place(x=220, y=250)
if os.stat("information.txt").st_size != 0:
    address_entry.insert(0, str(lines[4]))

alert_label = Label(top, text="Дата оповещения", font='Times 18')
alert_label.place(x=30, y=300)
alert_entry = Entry(top, width=30, font='Times 18')
alert_entry.place(x=220, y=300)
if os.stat("information.txt").st_size != 0:
    alert_entry.insert(0, str(lines[5]))

combo_UVD = Combobox(top, width=30, font='Times 18')
combo_UVD['values'] = ('УМВД', 'Зареченский', 'Криволученский', 'Привокзальный',
                       'Советский', 'Центральный', 'Ленинский', 'Ильинское',
                       'Косогорское', 'Скуратовский')

if os.stat("information.txt").st_size != 0:
    combo_UVD.set(lines[6])
else:
    combo_UVD.set('УМВД')

UVD_label = Label(top, text="УВД", font='Times 18').place(x=30, y=350)
combo_UVD.place(x=220, y=350)

btn = Button(top, text='Выбрать\n путь сохранения\n документа', command=insert_text, width=16, height=3, font='Times 14')
btn.place(x=33, y=400)

path_save_label = Label(top, text=f"{file_directory}", font='Times 16', wraplength=450)
path_save_label.place(x=220, y=400)

btn = Button(top, text='Создать документ',
             command=print_doc, width=14, font='Times 18')
btn.place(x=220, y=500)




top.mainloop()
