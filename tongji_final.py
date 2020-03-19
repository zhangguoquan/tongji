#-*- coding:utf-8 -*-
from tkinter import *
from tkinter.filedialog import askdirectory
from tkinter.filedialog import askopenfilename
import os
import re
import pandas as pd
import numpy as np

def sdata(a):
    return(a[0:10])
def s1data(a):
    return(a[0:20])
def srcselectPath(): 
    filename = askopenfilename(title='打开考勤数据文件', filetypes=[('excel', '*.xlsx'), ('All Files', '*')],initialdir='d:\\SourceCode')
    srcpath.set(filename)
    
def getallPath():
    df=pd.read_excel(srcpath.get(),sheet_name="车辆")
    data = df.loc[:,['姓名','入场时间']].values
    data1 = df.loc[:,['姓名','出场时间']].values

  
    writer = pd.ExcelWriter("d:\\SourceCode\\key1.xlsx")
    df1 = pd.DataFrame(data,columns=["姓名","打卡时间"])
    df1.to_excel(writer,index = False)
    start = df1.shape[0]
    df2 = pd.DataFrame(data1,columns=["姓名","打卡时间"])
    df2.to_excel(writer,index = False,startrow=start,header=False)
    
    start = start+df2.shape[0]
    df=pd.read_excel(srcpath.get(),sheet_name="桂平门禁")
    data = df.loc[:,['姓名','时间']].values
    df3 = pd.DataFrame(data,columns=["姓名","打卡时间"])
    df3.to_excel(writer,index = False,startrow=start,header=False)
    start = df3.shape[0] +start
    df=pd.read_excel(srcpath.get(),sheet_name="紫竹门禁")
    data = df.loc[:,['姓名','时间']].values
    df3 = pd.DataFrame(data,columns=["姓名","打卡时间"])
    df3.to_excel(writer,index = False,startrow=start,header=False)
    writer.save()
   
   
    df=pd.read_excel("d:\\SourceCode\\key1.xlsx")
    df["打卡时间"] = df["打卡时间"].astype("str")
    df['日期'] = df.apply(lambda x:sdata(x['打卡时间']), axis=1)
    df1=df.groupby(['姓名','日期']).agg([{"打卡时间":[('上班打卡','min'),('下班打卡','max')]}]).reset_index()

    data = df1.iloc[:,[0,1,2,3]].values
    df2 = pd.DataFrame(data,columns=["姓名","日期","上班打卡","下班打卡"])
    

   
    df2['下班打卡'] = df2.apply(lambda x:s1data(x['下班打卡']), axis=1)
    df2['上班打卡'] = df2.apply(lambda x:s1data(x['上班打卡']), axis=1)

    df2['上班打卡'] = pd.to_datetime(df2['上班打卡'])
    df2['下班打卡'] = pd.to_datetime(df2['下班打卡'])
    my_timedelta = np.timedelta64(1,'h')
    df2['上班时间'] = (df2['下班打卡'] - df2['上班打卡']).values/my_timedelta
    
    writer = pd.ExcelWriter("d:\\SourceCode\\key5.xlsx")
    df2.to_excel(writer,index = False)
    writer.save()
 

root = Tk()


srcpath = StringVar()

root.title("考勤数据整理")
root.geometry("445x80+400+300")      #宽 x 高 + 左边距 + 上边距
root['bg'] =  "lightblue"
Label(root,bg = 'lightblue', text = "数据源").grid(row = 0, column = 0)
Entry(root, textvariable = srcpath,width = 50).grid(row = 0, column = 1)
Button(root,bg = 'lightblue',text = "选择", command = srcselectPath).grid(row = 0, column = 2)

bttn = Button(root,text = "开始处理", width = 20,bg = "lightblue",command = getallPath).grid(row = 1,column = 1)

root.mainloop()