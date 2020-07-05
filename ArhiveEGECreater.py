import tkinter
import os
from tkinter import filedialog, END
from tkinter import messagebox as mb
from zipfile import ZipFile
import re


def esc(event=None):
    window.destroy()
    quit()


def ChooseVedPPE(event=None):
    global VedPPE_dir, VedPPE
    VedPPE_dir = filedialog.askdirectory()
    VedPPE = os.listdir(VedPPE_dir)
    VedomostyPPE_list.delete(0, END)
    trash = []
    for i in VedPPE:
        if re.fullmatch(r'70_PPE_\d{4}_\d{4}.[0-1][0-9].[0-3][0-9].*pdf', i):
            VedomostyPPE_list.insert(END, i[7:11])
        else:
            trash.append(i)
    if trash != []:
        mb.showwarning(title='Это не ведомость', message='\n'.join(trash))
    for i in trash:
        VedPPE.remove(i)


def ChooseRasPPE(event=None):
    global RasPPE_dir, RasPPE
    RasPPE_dir = filedialog.askdirectory()
    RasPPE = os.listdir(RasPPE_dir)
    RassadkaPPE_list.delete(0, END)
    trash = []
    for i in RasPPE:
        if re.fullmatch(r'\d{4}-[0-1][0-9]-[0-3][0-9]=\w{3}=\d\d-\d{4}.*pdf', i):
            RassadkaPPE_list.insert(END, i[18:22])
        else:
            trash.append(i)
    if trash != []:
        mb.showwarning(title='Это не рассадка', message='\n'.join(trash))
    for i in trash:
        RasPPE.remove(i)


def ChooseBoth(event=None):
    global BothPPE_dir, BothPPE
    BothPPE_dir = filedialog.askdirectory()
    BothPPE = os.listdir(BothPPE_dir)
    Both_list.delete(0, END)
    trash = []
    for i in BothPPE:
        if re.fullmatch(r'\d{4}-[0-1][0-9]-[0-3][0-9]=\w{3}=\d\d-\d{4}.*pdf', i) or re.fullmatch(
                r'70_PPE_\d{4}_\d{4}.[0-1][0-9].[0-3][0-9].*pdf', i):
            sms = mb.askyesno('Уверены?', 'Добавить файл %s в список загружаемых во все архивы?' % i)
            if sms == True:
                Both_list.insert(END, i)
            else:
                trash.append(i)
        else:
            Both_list.insert(END, i)
    for i in trash:
        BothPPE.remove(i)


def check(event=None):
    test_list.delete(0, END)
    test_list.insert(END, 'Ведомости есть, рассадки нет:')
    VedPPE_i = []
    RasPPE_i = []
    count = 0
    for i in VedPPE:
        VedPPE_i.append(i[7:11])
    for i in RasPPE:
        RasPPE_i.append(i[18:22])
    for i in VedPPE_i:
        if not i in RasPPE_i:
            test_list.insert(END, i)
            count = count + 1
    if count == 0:
        test_list.delete(0, END)
    test_list.insert(END, 'Рассадка есть, ведомостей нет:')
    for i in RasPPE_i:
        if not i in VedPPE_i:
            test_list.insert(END, i)
            count = count + 1
    if count == 0:
        test_list.delete(0, END)
        test_list.insert(END, 'Расхождений нет')


def del_file(event=None):
    file = Both_list.get(Both_list.curselection()[0])
    BothPPE.remove(file)
    Both_list.delete(0, END)
    for i in BothPPE:
        Both_list.insert(END, i)


def arhive(event=None):
    arhive_dir = filedialog.askdirectory()
    if arhive_dir == VedPPE_dir:
        sms = mb.askyesno(title='Атата', message='Нельзя сохранять архивы \nв папку с ведомостями '
                                                 '\n\nВыбрать новую папку?')
        if sms == True:
            arhive()
    elif arhive_dir == RasPPE_dir:
        sms = mb.askyesno(title='Атата', message='Нельзя сохранять архивы \nв папку с рассадкой '
                                                 '\n\nВыбрать новую папку?')
        if sms == True:
            arhive()
    elif arhive_dir == BothPPE_dir:
        sms = mb.askyesno(title='Атата', message='Нельзя сохранять архивы \nв папку с общими документами '
                                                 '\n\nВыбрать новую папку?')
        if sms == True:
            arhive()
    else:
        for i in VedPPE:
            with ZipFile(arhive_dir + '/70' + i[7:11] + '.zip', 'w') as newarh:
                for k in BothPPE:
                    newarh.write(BothPPE_dir + '/' + k, k)
                for j in RasPPE:
                    if i[7:11] == j[18:22]:
                        newarh.write(RasPPE_dir + '/' + j, j)
                for j in VedPPE:
                    if i[7:11] == j[7:11]:
                        newarh.write(VedPPE_dir + '/' + j, j)
        sms = mb.showinfo(title='Успех', message='Архивы готовы')


BothPPE = []
BothPPE_dir = []
RasPPE = []
RasPPE_dir = []
window = tkinter.Tk()
window.title('ArhiveEGECreater')

d1 = tkinter.Label(height="3", width="5")
d1.grid(row=0, column=0)

ChooseVedPPE_button = tkinter.Button(window, text='Папка с ведомостями', command=ChooseVedPPE)
ChooseVedPPE_button.grid(row=0, column=1, columnspan=2)

d1 = tkinter.Label(height="3", width="5")
d1.grid(row=0, column=3)

ChooseRasPPE_button = tkinter.Button(window, text='Папка с рассадкой', command=ChooseRasPPE)
ChooseRasPPE_button.grid(row=0, column=4, columnspan=2)

d1 = tkinter.Label(height="3", width="5")
d1.grid(row=0, column=6)

ChooseBoth_button = tkinter.Button(window, text='Папка с общими файлами', command=ChooseBoth)
ChooseBoth_button.grid(row=0, column=7, columnspan=2)

# поле вывода списка ППЭ, для которых найдены ведомости
VedomostyPPE_list = tkinter.Listbox(height=15, width=20)
VedPPE_scroll = tkinter.Scrollbar(command=VedomostyPPE_list.yview)
VedomostyPPE_list.config(yscrollcommand=VedPPE_scroll.set)
VedomostyPPE_list.grid(row=1, column=1)
VedPPE_scroll.grid(row=1, column=2)
# поле вывода списка ППЭ, для которых найдены файлы рассадки
RassadkaPPE_list = tkinter.Listbox(height=15, width=20)
RasPPE_scroll = tkinter.Scrollbar(command=RassadkaPPE_list.yview)
RassadkaPPE_list.config(yscrollcommand=VedPPE_scroll.set)
RassadkaPPE_list.grid(row=1, column=4)
RasPPE_scroll.grid(row=1, column=5)
# поле вывода списка общих для всех ППЭ файлов, которые нужны в архиве
Both_list = tkinter.Listbox(height=15, width=100)
Both_scroll = tkinter.Scrollbar(command=Both_list.yview)
Both_list.config(yscrollcommand=Both_scroll.set)
Both_list.grid(row=1, column=7)
Both_scroll.grid(row=1, column=8)

# Удалить файл из списка добавляемых во все архивы файлов
del_botton = tkinter.Button(window, text='Удалить', command=del_file)
del_botton.grid(row=2, column=7, columnspan=2)

# Проверка наличия файла рассадки при наличие файла ведомостей
d1 = tkinter.Label(height="3", width="25")
d1.grid(row=2, column=0, columnspan=5)
test_list = tkinter.Listbox(height=5, width=30)
test_scroll = tkinter.Scrollbar(command=test_list.yview)
test_list.config(yscrollcommand=test_scroll.set)
test_list.grid(row=3, column=1)
test_scroll.grid(row=3, column=2)

test_button = tkinter.Button(window, text='Проверка', command=check)
test_button.grid(row=3, column=3, columnspan=2)

# Кнопка генерации архивов
arhive_button = tkinter.Button(window, text='Создать архивы', command=arhive)
arhive_button.grid(row=3, column=5, columnspan=3)

window.protocol('WM_DELETE_WINDOW', esc)
window.mainloop()
