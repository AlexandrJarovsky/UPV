# coding: utf8

from tkinter import *
from tkinter.scrolledtext import ScrolledText
from tkinter import ttk
from datetime import datetime, timedelta
import time
import pyodbc
import sqlite3
import os, sys, subprocess
import tep_print
import opros, tep_print, restoredata
import importlib

if __name__ == '__main__':

    try:

        def on_exit():

            open(os.curdir + '\\report_error.txt', 'a').write(
                'Close programm : ' + datetime.today().strftime("%d.%m.%Y %H:%M:%S") + '\n' + '=' * 30 + '\n')
            root.destroy()

        def set_day():
            Count_day.after(3600000, set_day)
            Count_day.set(time.localtime()[2])
            label2['text'] = 'День в ' + time.strftime('%B')

        def tick():
            label.after(1000, tick)
            label['text'] = str(datetime.now()).split('.')[0]

        def tick_opros():
            label_opros.after(15000, tick_opros)
            start = time.perf_counter()
            opros.opros(int(conf_dict['rasch_proc']), int(conf_dict['lab_density']))
            st_temp = time.perf_counter() - start
            label_opros['text'] = 'Last update   ---  ' + str(opros.LastDate).split('.')[0] + '  (' + '{:.4f}'.format(
                st_temp) + ')'

        def Lclick(event):
            if Count_mounth.get() != '':
                Count_day1['state'] = 'enable'
                But_print1['state'] = 'active'
                But_print1_f['state'] = 'active'
            else:
                Count_day1['state'] = 'disable'
                But_print1['state'] = 'disable'
                But_print1_f['state'] = 'disable'

        def check_print():
            if len(cursor.execute("SELECT * FROM seting").fetchall()) != 0:
                sql = "UPDATE seting SET print_check= %d" % int(cvar1.get())
            else:
                sql = "INSERT seting VALUES ('%s')" % int(cvar1.get())
            cursor.execute(sql)
            cnn.commit()

        def changeMonth():
            cursor.execute(r"SELECT name FROM sqlite_master WHERE type='table' and name like 'Archiv_%' order by name desc")
            Count_mounth['value']= [row[0] for row in cursor.fetchall()]

        with open(os.curdir + '\\config.ini') as conf:
            conf_list = []
            for c in conf:
                if c != '\n':
                    conf_list.append(c.strip().split('='))
            conf_dict = dict(conf_list)

        root = Tk()
        root.title(conf_dict['Title'])
        root.geometry('500x400')
        root.resizable(FALSE, False)

        open(os.curdir + '\\report_error.txt', 'a').write(
            'Start programm : ' + datetime.today().strftime("%d.%m.%Y %H:%M:%S") + '\n')

        try:

            prog = [line.split() for line in
                    subprocess.check_output('tasklist /FI "WINDOWTITLE eq ' + root.title() + '"').splitlines()]
            if len(prog) > 1:
                open(os.curdir + '\\report_error.txt', 'a').write(
                    'Dable start : ' + datetime.today().strftime("%d.%m.%Y %H:%M:%S") + '\n')
                sys.exit()

        except Exception as err:
            print(err)

        cnn = sqlite3.connect(os.curdir + conf_dict['DATABASE'])
        cursor = cnn.cursor()

        restoredata.restoredata(int(conf_dict['rasch_proc']))
        cursor.execute("SELECT * FROM %s order by sortorder" % (conf_dict['Table_work']))
        for rows in cursor.fetchall():
            ii = 0
            for row in rows:
                if row == None:
                    if cursor.execute("pragma table_info(Table1)").fetchall()[ii][2] == 'TEXT':
                        sql = "UPDATE %s SET %s= %s WHERE Shifr='%s'" % (
                        conf_dict['Table_work'], cursor.execute("pragma table_info(Table1)").fetchall()[ii][1], str("""''"""), rows[0])
                    else:
                        sql = "UPDATE %s SET %s= %s WHERE Shifr='%s'" % (
                        conf_dict['Table_work'], cursor.execute("pragma table_info(Table1)").fetchall()[ii][1], 0, rows[0])
                    cursor.execute(sql)
                    cnn.commit()
                ii += 1

        label = Label(font='sans 10')
        label_opros = Label(font='sans 10')

        label1 = Label(root, font='sans 10')
        label1.pack()
        label1['text'] = str(datetime.now()).split('.')[0]  # + str(datetime.now().year)

        list1 = []
        for i in range(1, 32):
            list1.append(str(i))

        f_top = LabelFrame(root, height=50, bg='green', fg='white', text="Печать")
        f_top.config(relief=RAISED, borderwidth=1, font=('courier', 12, 'normal'))
        f_top['width'] = len(f_top['text'])
        f_top.pack(expand=YES, fill=BOTH)

        f_now = LabelFrame(f_top, height=50, bg='green', fg='white', text="Текущий месяц")
        f_now.config(relief=RAISED, borderwidth=1, font=('courier', 12, 'normal'))
        f_now.pack(side=LEFT, expand=YES, fill=BOTH)

        f_now1 = Frame(f_now, bg='green')
        f_now1.config(borderwidth=1)
        f_now1.pack(side=BOTTOM, expand=NO, fill=BOTH)

        f_arch = LabelFrame(f_top, height=50, bg='green', fg='white', text="Из архива")
        f_arch.config(relief=RAISED, borderwidth=1, font=('courier', 12, 'normal'))
        f_arch.pack(side=LEFT, expand=YES, fill=BOTH)

        f_arch1 = Frame(f_arch, bg='green')
        f_arch1.config(borderwidth=1)
        f_arch1.pack(side=BOTTOM, expand=NO, fill=BOTH)

        label2 = Label(f_now, font='sans 10')
        label2.pack(side=TOP, anchor=NW)
        label2['text'] = r'День в ' + time.strftime('%B')

        Count_day = ttk.Combobox(f_now, height=1, width=15, value=list1)
        Count_day.pack(side=TOP, anchor=NW)
        Count_day.after_idle(set_day)

        label3 = Label(f_arch, font='sans 10')
        label3.pack(side=TOP, anchor=NW)
        label3['text'] = 'Выбор архива'

        Count_mounth = ttk.Combobox(f_arch, height=1, width=15, postcommand=changeMonth)
        Count_mounth.pack(side=TOP, anchor=NW)

        label3 = Label(f_arch, font='sans 10')
        label3.pack(side=TOP, anchor=NW)
        label3['text'] = 'День'

        Count_day1 = ttk.Combobox(f_arch, height=1, width=15, value=list1, state="disable")
        Count_day1.pack(side=TOP, anchor=NW)
        Count_day1.bind('<Button-1>', Lclick)

        But_print = Button(f_now1, text='Печать', command=lambda: tep_print.prin_pr(day_prn=Count_day.get()))
        But_print.pack(side=LEFT)

        But_print_f = Button(f_now1, text='Печать в EXCEL',
                             command=lambda: tep_print.prin_pr(day_prn=Count_day.get(), flag_prn=False))
        But_print_f.pack(side=RIGHT)

        cvar1 = BooleanVar()
        set_print = cursor.execute("SELECT * FROM seting").fetchone()
        if len(cursor.execute("SELECT * FROM seting").fetchall()) != 0 and set_print[0] != None:
            cvar1.set(set_print[0])
        else:
            cvar1.set(1)
        c1 = Checkbutton(f_now1, text="Печать на принтер", variable=cvar1, onvalue=1, offvalue=0,
                         command=lambda: check_print())
        c1.pack()

        But_print1 = Button(f_arch1, text='Печать',
                            command=lambda: tep_print.prin_pr(TableName_prn=Count_mounth.get(), day_prn=Count_day1.get(),
                                                             flag='1'), state="disable")
        But_print1.pack(side=LEFT)

        But_print1_f = Button(f_arch1, text='Печать в EXCEL',
                              command=lambda: tep_print.prin_pr(TableName_prn=Count_mounth.get(),
                                                               day_prn=Count_day1.get(), flag='1', flag_prn=False),
                              state="disable")
        But_print1_f.pack(side=RIGHT)

        label_opros.pack()
        label_opros.after_idle(tick_opros)

        label.pack()
        label.after_idle(tick)

        menubar = Menu(root)
        root.config(menu=menubar)

        fileMenu = Menu(menubar)
        menubar.add_cascade(label="File", menu=fileMenu)

        fileMenu.add_command(label="opros", command=lambda: importlib.reload(opros))
        fileMenu.add_command(label="fast_update", command=lambda: importlib.reload(opros.fast_update))
        fileMenu.add_command(label="print", command=lambda: importlib.reload(tep_print))
        fileMenu.add_command(label="print_excel", command=lambda: importlib.reload(tep_print.print_exel))
        fileMenu.add_command(label="Exit", command=on_exit)

        root.protocol("WM_DELETE_WINDOW", on_exit)
        root.mainloop()

    except Exception as err:
        print(err)
