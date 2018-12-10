# -*- coding:utf-8 -*-
import configparser
class readConfig:
    def __init__(self):
        self.cf = configparser.ConfigParser()
        self.cf.read('report.config',encoding='utf-8')

        self.dict = {}
        self.kvs = self.cf.items('filename')

        for key in self.kvs:
            arr = key[1].split(",")
            self.dict.setdefault(key[0],arr)

        self.sql_path = self.cf.get('path','sql_path')
        self.sql_excel =self.cf.get('path','excel_path')
        self.host =self.cf.get('database','host')
        self.user =self.cf.get('database','user')
        self.passwd =self.cf.get('database','passwd')
        self.db =self.cf.get('database','db')
        print(self.host)
readConfig()