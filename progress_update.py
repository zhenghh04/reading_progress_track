#!/usr/bin/env python
# -*- coding: utf-8 -*- 
import pandas as pd
import gspread
from oauth2client.service_account import ServiceAccountCredentials
import csv
from datetime import datetime

class Progress:
    def __init__(self, fstr=''):
        self.data=[]
        self.num_readers=0
        self.readers=[]
        self.sh = None
        self.sheet=fstr
        if fstr!='':
            self.getRemote(fstr)
    def getNumPages(self, name="", t=["11/01/2019", "12/01/2019"]):
        dat = self.data[self.data["您的姓名"]==name]
        dat=dat[dat.Timestamp>t[0]]
        dat=dat[dat.Timestamp<t[1]]
        p = 0
        for rec in dat["文集页数 [如: 20-100]"].values:
            tmp = rec.split('-')
            p += int(tmp[1]) - int(tmp[0])+1
        return p
    def loadCSV(self, fstr):
        self.data=pd.read_csv(fstr)
    def getReaders(self):
        self.readers=list(set(self.data["您的姓名"].values))
        self.num_readers = len(self.readers)
        return self.readers
    def getRemote(self, sheet=""):
        scope = ['https://spreadsheets.google.com/feeds',
         'https://www.googleapis.com/auth/drive']
        credentials = ServiceAccountCredentials.from_json_keyfile_name('/Users/zhenghh/Documents/Christ/CWWL/cwwl-midwest.json', scope)
        gc = gspread.authorize(credentials)
        sh = gc.open(sheet)
        self.sh = sh
        worksheet = sh.get_worksheet(0)
        list_of_lists = worksheet.get_all_values()
        headers=list_of_lists[0]
        self.data = pd.DataFrame(list_of_lists[1:], columns=headers)
    def getProgress(self, select="month", verbose=1):
        record={}
        today = datetime.today()
        if (select=="month"):
            t="%s/01/%s-%s/01/%s"%(today.month, today.year, today.month+1, today.year)
        elif (select=="year"):
            t="01/01/%s-12/31/%s"%(today.year, today.year)
        elif (select=="all"):
            t="01/01/0000-12/31/3100"
        else:
            t=select
        t=t.split('-')
        if (verbose==1):
            print("时间：%s - %s"%(t[0], t[1]))
            print("读者数: ", len(self.getReaders()))
            print("  姓名  页数")
        for user in self.getReaders():
            p=self.getNumPages(user, t=t)
            if (verbose==1):
                print("%3s %5d"%(user, p))
            record[user]=p
        return record
    def reportProgress(self):
        self.getRemote(self.sheet)
        record = {}
        today = datetime.today()
        readers = sorted(self.getReaders())

        for u in readers:
            record[u]=[]
        y = today.year
        for m in range(1,13):
            t="%02d/01/%s-%02d/01/%s"%(m, y, m+1, y)
            r = self.getProgress(select=t, verbose=0)
            for u in readers:
                record[u].append(r[u])
        rt = self.getProgress(select="all", verbose=0)
        try:
            worksheet = self.sh.worksheet("%s年每月进度统计"%y)
        except:
            worksheet=self.sh.add_worksheet(title="%s年每月进度统计"%y, rows=self.num_readers+1000, cols=20)
        worksheet.update_cell(1, 1, "%s年每月阅读页数统计(文集总页数逾十万)"%y)
        worksheet.update_cell(2, 1, "最新更新:%s"%today)
        worksheet.update_cell(4, 1, '姓名')
        worksheet.update_cell(4, 16, '姓名')
        for i in range(12):
            worksheet.update_cell(4, 2+i, '%s月'%(i+1))

        worksheet.update_cell(4, 14, '%s年'%(y))
        worksheet.update_cell(4, 15, '总共')
        r=5
        for u in readers:
            worksheet.update_cell(r, 1, u)
            worksheet.update_cell(r, 16, u)
            for i in range(12):
                if (record[u][i]!=0):
                     worksheet.update_cell(r, 2+i, record[u][i])
            worksheet.update_cell(r, 14, sum(record[u][:12]))
            worksheet.update_cell(r, 15, rt[u])
            r=r+1

        return record

def main():
    rp = Progress("cwwl-midwest-reading-progress")
    rp.reportProgress()

if __name__=="__main__":
    main()
