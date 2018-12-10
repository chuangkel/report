import calendar
import datetime
import threading
from tkinter import *
import tkinter
from tkinter import ttk
import tkinter.messagebox

from readfile import readConfig
from class_exe_config_rep import class_exe_report
class class_calendar:
    def __init__(self):
        self.window = tkinter.Tk()
        self.window.title('报表生成工具')
        self.frame_head = Frame(self.window, bd=4)
        self.frame_right = Frame(self.window, bd=5)
        self.frame_bottom = Frame(self.window, bd=10)

        self.frame_left = Frame(self.window, bd=2)
        self.now_time = datetime.datetime.now()
        self.now_arr = str(self.now_time).split("-")
        self.year, self.month = self.now_arr[0], self.now_arr[1]
        self.e = StringVar()
        self.e2 = StringVar()
        self.e_left = StringVar()
        self.e_head = StringVar()

        self.label_head = tkinter.Label(self.frame_head, textvariable=self.e_head).pack()
        self.label_left =tkinter.Text(self.frame_left)
        self.label_left.insert(END,"欢迎使用报表生成小工具")
        self.label_left.pack()

        self.readconfig = readConfig()

        self.e_head.set(self.year + "年" + self.month + "月")
        self.frame_init()
        self.show_now()
        self.window.mainloop()

    def refresh(self):
        self.e_left.set('开始生成'  + '报表...')
        self.label_left.after(1, self.refresh())

    def showMsg(self,new_date):
        msg = '确认生成确认日为' + new_date + '的报表？'
        result = tkinter.messagebox.askokcancel('提示', msg)
        if result:
            print(new_date)

            self.e_left.set('开始生成' + '报表...')

            e_report = class_exe_report()

            e_report.setConfig(self.readconfig)

            e_report.select_day = new_date
            e_report.label_left = self.label_left
            e_report.start()

    def show_calendar(self,year, month):
        week_arr = [['一', '二', '三', '四', '五', '六', '日'], ]
        add_null = [[0,0, 0, 0, 0, 0, 0],]

        c_arr = calendar.monthcalendar(int(year), int(month))
        print((6-len(c_arr))*add_null)
        c_arr = c_arr + (6-len(c_arr))*add_null
        c_arr = week_arr + c_arr
        print(c_arr)
        for i in range(len(c_arr)):
            for j in range(len(c_arr[0])):

                temp = c_arr[i][j]
                if temp == 0:
                    temp = "  "
                label = Label(self.frame_bottom, text=temp)
                label.grid(row=i, column=j, sticky=W + E + N + S, padx=15, pady=15)
                if c_arr[i][j] != 0:
                    label.bind('<ButtonRelease-1>', self.labelAction)


    def Show(self):
        self.year = self.e.get()
        self.month = self.e2.get()
        self.e_head.set(self.year + "年" + self.month + "月")
        self.show_calendar(str(self.year), str(self.month))

    def center_window(root, width, height):
        screenwidth = root.winfo_screenwidth()
        screenheight = root.winfo_screenheight()
        size = '%dx%d+%d+%d' % (width, height, (screenwidth - width) / 2, (screenheight - height) / 2)
        print(size)
        root.geometry(size)

    def show_now(self):
        self.show_calendar(self.year, self.month)

    def frame_init(self):
        label1 = tkinter.Label(self.frame_right, text="月份").pack()

        numberChosen = ttk.Combobox(self.frame_right, width=12, textvariable=self.e)
        numberChosen['values'] = list(range(2018, 2077, 1))  # 设置下拉列表的值
        numberChosen.current(0)
        numberChosen.bind("<<ComboboxSelected>>")
        numberChosen.pack()

        label2 = tkinter.Label(self.frame_right, text="年份").pack()
        numberChosen1 = ttk.Combobox(self.frame_right, width=12, textvariable=self.e2)
        numberChosen1['values'] = list(range(1, 13, 1))  # 设置下拉列表的值
        numberChosen1.current(0)
        numberChosen1.bind("<<ComboboxSelected>>")  # 绑定事件
        numberChosen1.pack()
        button = tkinter.Button(self.frame_right, text="确定", anchor='c', width=6, height=1, command=self.Show).pack()
        self.frame_head.pack(side=TOP)
        self.frame_left.pack(side=LEFT)
        self.frame_right.pack(side=RIGHT)
        self.frame_bottom.pack(side=BOTTOM)

    def labelAction(self,event):
        day = self.year.zfill(4) + self.month.zfill(2) + str(event.widget.cget('text')).zfill(2)
        self.showMsg(day)

#开始执行
c = class_calendar()