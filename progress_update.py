#!/opt/intel/intelpython3/bin/python
# -*- coding: utf-8 -*- 
import pandas as pd
import gspread
from oauth2client.service_account import ServiceAccountCredentials
import csv
from datetime import date, datetime
from apscheduler.schedulers.blocking import BlockingScheduler
from time import strptime

class Progress:
    def __init__(self, fstr=''):
        self.data=[]
        self.num_readers=0
        self.readers=[]
        self.sh = None
        self.sheet=fstr
        self.tt = None
        if fstr!='':
            self.getRemote(fstr)
    def getNumPages(self, name="", t=["2019/11/01", "2019/12/01"]):
        p = 0
        n = len(self.data)
        i = 0
        for i in range(n):
            rec = self.data["文集页数 [如: 20-100]"].values[i]
            if (self.date.values[i]<=t[1]) and (self.date.values[i]>=t[0]) and self.data["您的姓名"][i]==name:
                try:
                    a = int(rec.split('-')[0])
                    tmp = rec.split('-')
                    p += int(tmp[1]) - int(tmp[0])+1
                except:
                    print("Correcting for %s" %rec)
                    a, b, c, d = rec.split()
                    l = strptime(b,'%b').tm_mon
                    r = int(c)+1
                    p += r - l + 1
        return p
    def getChurch(self):
        self.church={}
        for u in self.getReaders():
            dat = self.data[self.data["您的姓名"]==u]
            self.church[u] = dat["召会 [如: Chicago]"].values[0].strip()
        return self.church
    def loadCSV(self, fstr):
        self.data=pd.read_csv(fstr)
        n = len(self.data)
        t={}
        t["date"] = []
        for i in range(n):
            m, d, y = self.data.Timestamp[i].split()[0].split("/")
            t["date"].append("%s/%s/%s" %(y, m, d))
        self.date = pd.DataFrame(data=t)
        
    def getReaders(self):
        self.readers=list(set([n.strip() for n in self.data["您的姓名"].values]))
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
        n = len(self.data)
        t={}
        t["date"] = []
        for i in range(n):
            m, d, y = self.data.Timestamp[i].split()[0].split("/")
            t["date"].append("%s/%s/%s" %(y, m, d))
        self.date = pd.DataFrame(data=t)
    def getProgress(self, t, verbose=0):
        record={}
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
    def reportProgress(self, y=0):
        self.getRemote(self.sheet)
        record = {}
        today = datetime.today()
        readers = sorted(self.getReaders())
        for u in readers:
            record[u]=[]
        if (y==0):
            y = today.year
        for m in range(1,13):
            t="%s/%d/1-%s/%d/31"%(y, m, y, m)
            r = self.getProgress(t, verbose=0)
            for u in readers:
                record[u].append(r[u])
        rt = self.getProgress(t="0000/0/0-2200/0/0", verbose=0)
        try:
            worksheet = self.sh.worksheet("%s年每月进度统计"%y)
        except:
            worksheet=self.sh.add_worksheet(title="%s年每月进度统计"%y, rows=self.num_readers+1000, cols=20)
        c = self.getChurch()
        #print(c)
        worksheet.update_cell(1, 1, "%s年每月阅读页数统计(文集总页数逾十万)"%y)
        worksheet.update_cell(2, 1, "最新更新:%s (每小时更新一次)"%today)
        worksheet.update_cell(4, 1, '召会')
        worksheet.update_cell(4, 2, '姓名')
        worksheet.update_cell(4, 17, '姓名')
        for i in range(12):
            worksheet.update_cell(4, 3+i, '%s月'%(i+1))

        worksheet.update_cell(4, 15, '%s年'%(y))
        worksheet.update_cell(4, 16, '总共')
        r=5
        for u in readers:
            worksheet.update_cell(r, 1, c[u])
            worksheet.update_cell(r, 2, u)
            worksheet.update_cell(r, 17, u)
            for i in range(12):
                if (record[u][i]!=0):
                     worksheet.update_cell(r, 3+i, record[u][i])
            worksheet.update_cell(r, 15, sum(record[u][:12]))
            worksheet.update_cell(r, 16, rt[u])
            r=r+1

        return record
    
def main():
    rp = Progress("cwwl-midwest-reading-progress")
    rp.reportProgress(y=2019)
    rp.reportProgress()

if __name__=="__main__":
    main()
#    scheduler = BlockingScheduler()
#    scheduler.add_job(main, 'interval', minutes=1)
#    scheduler.start()
