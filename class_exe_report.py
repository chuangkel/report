# coding=UTF-8
import sys
import threading

import pymysql, xlwt
import os
import time
import re
import tkinter.messagebox

class class_exe_report(threading.Thread):
    def __init__(self):
        super(class_exe_report, self).__init__()  # 重写父类属性
        self.dict = {
            "C:\报表\基金资金轧差清算报表_分集合并": ["3.sql", "sheet1"],
            "C:\报表\销售商确认汇总报表": ["5.sql", "sheet1"],
            "C:\报表\资金交收报表（305）": ["10.305.sql", "sheet1"],
            "C:\报表\资金交收报表（926）": ["10.926.sql", "sheet1"],
            "C:\报表\资金交收报表（汇总）": ["10.sql", "sheet1"],
            "C:\报表\货币份额流入流出统计": ["1.sql", "sheet1"],
            "C:\报表\基金的资金清算报表(基金)": ["4.sql", "sheet1"],
            "C:\报表\JY基金申赎及基金投资人结构日报表": ["2.3.sql"
                                      "|2.4.sql"
                                      "|2.1.sql"
                                      "|2.5.sql"
                                      "|2.2.sql", "保有份额|保有净值|基金交易|认申购|账户"],
        }

        self.path_one = r"C:\sqls\sql_one"
        self.path_two = r"C:\sqls\sql_two"
        self.path_report = 'C:\\报表'

        self.host, self.user, self.passwd, self.db = '127.0.0.1', 'root', 'root', 'yunta'
        self.conn = pymysql.connect(user=self.user, host=self.host, port=3306, passwd=self.passwd, db=self.db, charset='utf8')

        self.borders = xlwt.Borders()
        self.borders.bottom = xlwt.Borders.THIN
        self.borders.left = xlwt.Borders.THIN
        self.borders.right = xlwt.Borders.THIN
        self.borders.top = xlwt.Borders.THIN
        self.style = xlwt.XFStyle()
        self.style.borders = self.borders
        
    #查询申请日期和确认日期
    def select_resAndConfirmDate(self,selectDate):
        cur = self.conn.cursor()
        requestdate_sql =  'select getjobDate('+selectDate+',0) from dual;'
        confirmdate_sql = 'select getjobDate('+selectDate+',1) from dual;'
        cur.execute(requestdate_sql)
        self.requestdate = cur.fetchall()[0][0]
        cur.execute(confirmdate_sql)
        self.confirmdate = cur.fetchall()[0][0]
        print("申请日期为",self.requestdate)
        print("确认日期为",self.confirmdate)
        
    def tab_2_excel(self,path, sqls, sheets, name):
        book = xlwt.Workbook(encoding='utf-8', style_compression=0)
        for id in range(len(sqls)):

            each = sqls[id]
            cur = self.conn.cursor()
            os.chdir(path)

            sheet = book.add_sheet(sheets[id])

            sql = ""
            with open(each, "r", encoding="utf-8") as f:
                for each_line in f.readlines():
                    sql += each_line
                if sql:
                    cur.execute(sql)
                    fields = cur.description
                    all_data = cur.fetchall()

                    for col in range(0, len(fields)):
                        sheet.write(0, col, fields[col][0], self.style)

                    row = 1
                    for data in all_data:
                        for col, field in enumerate(data):
                            sheet.write(row, col, field, self.style)
                        row += 1
            f.close()

        os.chdir(self.path_report)
        book.save('%s' % name)
        self.export_info("生成"+ "%s.xls" % name)

    def replace_date(self,b):
        paths = [self.path_one, self.path_two]
        for i in paths:
            os.chdir(i)
            for file in os.listdir(i):
                with open(file, '+r', encoding="utf-8") as f:
                    t = f.read()
                    regex = re.compile(r"(20\d{6})")
                    list = regex.findall(t)
                    dict = {}
                    if None != list:
                        for li in list:
                            dict.setdefault(li, 0)
                            dict[li] += 1
                    for key in dict.keys():
                        self.export_info(i+file+ "中替换日期"+ str(key) + str(dict[key])+"次")
                        t = t.replace(key, b)
                    f.seek(0, 0)
                    f.write(t)

    def export_info(self, msg):
        self.label_left.insert(tkinter.END,'\n'+msg)
    def run(self):
        # print("输入日期".center(40, "="))
        # self.confirmdate = self.input_control()
        self.select_resAndConfirmDate(self.select_day)
        self.export_info("替换日期".center(40, "="))
        self.replace_date(self.confirmdate)
        self.export_info("删除报表".center(40, "="))
        for i in os.listdir(self.path_report):
            path_file = os.path.join(self.path_report, i)
            if os.path.isfile(path_file):
                self.export_info("删除文件" + path_file)
                os.remove(path_file)
        self.export_info("生成报表".center(40, "="))

        self.startwith_sql()

        for k, v in self.dict.items():
            sqls = v[0].split("|")
            sheets = v[1].split("|")
            if k == 'C:\报表\JY基金申赎及基金投资人结构日报表':
                filename = k + self.requestdate + ".xls"
            else:
                filename = k + self.confirmdate + ".xls"
            self.tab_2_excel(self.path_one, sqls, sheets, filename)
        tkinter.messagebox.askokcancel('提示', "报表生成成功！")

    def sqlname2excelname(self,name, sqlfilename):
        book = xlwt.Workbook(encoding='utf-8', style_compression=0)

        each = sqlfilename
        sheet = "sheet1"

        cur = self.conn.cursor()
        os.chdir(self.path_two)

        sheet = book.add_sheet(sheet)

        sql = ""
        with open(each, "r", encoding="utf-8") as f:
            for each_line in f.readlines():
                sql += each_line
            if sql:
                cur.execute(sql)
                fields = cur.description
                all_data = cur.fetchall()

                for col in range(0, len(fields)):
                    sheet.write(0, col, fields[col][0], self.style)

                row = 1
                for data in all_data:
                    for col, field in enumerate(data):
                        sheet.write(row, col, field, self.style)
                    row += 1
        f.close()
        os.chdir(self.path_report)
        book.save('%s.xls' % name)
        self.export_info("生成" + self.path_report + "%s.xls" % name)

    def startwith_sql(self):
        for i in os.listdir(self.path_two):
            whole_path = os.path.join(self.path_two, i)
            if os.path.isfile(whole_path):
                name = i.split(".")[0]
                self.sqlname2excelname(name + self.confirmdate, i)