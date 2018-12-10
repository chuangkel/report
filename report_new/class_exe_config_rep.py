# coding=UTF-8
import threading
import pymysql, xlwt
import os
import re
import tkinter.messagebox

class class_exe_report(threading.Thread):
    def __init__(self):
        super(class_exe_report, self).__init__()  # 重写父类属性

        self.borders = xlwt.Borders()
        self.borders.bottom = xlwt.Borders.THIN
        self.borders.left = xlwt.Borders.THIN
        self.borders.right = xlwt.Borders.THIN
        self.borders.top = xlwt.Borders.THIN
        self.style = xlwt.XFStyle()
        self.style.borders = self.borders

    def setConfig(self,readconfig):
        self.dict = readconfig.dict
        self.path_one = readconfig.sql_path
        self.path_report = readconfig.sql_excel
        print(readconfig.user)
        self.conn = pymysql.connect(user=readconfig.user, host=readconfig.host, port=3306, passwd=readconfig.passwd, db=readconfig.db, charset='utf8')


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
        paths = [self.path_one]
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
        self.label_left.see(tkinter.END)
    def run(self):
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

        for k, v in self.dict.items():
            sqls = v[0].split("|")
            sheets = v[1].split("|")
            if k == 'JY基金申赎及基金投资人结构日报表':
                filename = k + self.requestdate + ".xls"
            else:
                filename = k + self.confirmdate + ".xls"
            self.tab_2_excel(self.path_one, sqls, sheets, filename)
        tkinter.messagebox.askokcancel('提示', "报表生成成功！")
