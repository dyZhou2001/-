# 好好学习 天天向上
# {2023/11/14} {21:42}
from tkinter import *
from tkinter.scrolledtext import ScrolledText
import tkinter as tk
import os
import sheetplot

def input_good():
    desktop_path = os.path.join(os.environ['USERPROFILE'], 'Desktop')
    def load():
        with open(desktop_path+"\\商品表.txt", 'r') as file:
            contents.delete('1.0', END)
            contents.insert(INSERT, file.read())

    def save():
        with open(desktop_path+"\\商品表.txt", 'a+') as file:
            idx=lb.curselection()
            file.write('\n'+str(idx[0])+" "+good_info.get())
        load()
        good_info.delete(0,"end")
    top = Tk()
    top.title("input_good")
    lb=tk.Listbox(top)
    lb.insert(0,'按面积计算')
    lb.insert(1, '按个计算')
    lb.pack()
    contents = ScrolledText(top)
    contents.pack(side=BOTTOM, expand=True, fill=BOTH)
    good_info = Entry(top)
    good_info.pack(side=LEFT, expand=True, fill=X)
    Button(top,text='保存', command=save).pack(side=LEFT)
    mainloop()
def delete_good():
    desktop_path = os.path.join(os.environ['USERPROFILE'], 'Desktop')
    allGoodList=[]
    window = Tk()
    window.geometry("300x300")
    lb1 = tk.Listbox(window)
    lb1.grid(row=0, column=0,sticky="w",padx = 10, pady = 5)
    with open(desktop_path + "\\商品表.txt", 'r') as file:
        for line in file:
            if line == "\n": continue
            lb1.insert("end", line)
            allGoodList.append(line)
    def deletelist():
        del allGoodList[lb1.curselection()[0]]
        lb1.delete(lb1.curselection(),lb1.curselection())
        with open(desktop_path + "\\商品表.txt", 'w') as file:
            for line in allGoodList:
                file.writelines(line)
    Button(window, text='删除', command=deletelist).grid(row=2, column=0,sticky="w",padx = 10, pady = 5)
    mainloop()
class Customer:
    def __init__(self,cosname,teleph,address,date,goodlist):
        self.name=cosname
        self.phonenum=teleph
        self.address=address
        self.date=date
        self.goodlist=goodlist
class good_for_PDF:
    def __init__(self,gtp,gnm,gpr,per):
        self.goodtp = gtp
        self.goodnm = gnm
        self.goodpr = gpr
        self.danwei=per
        self.description=""
        self.number=0
        self.allmoney=0
class AreaGood(good_for_PDF):
    def __init__(self,gtp,gnm,gpr,per,long,width):
        super().__init__(gtp,gnm,gpr,per)
        self.long=long
        self.width=width
        self.number=round(self.long*self.width,4)
        self.allmoney=round(float(self.number)*float(self.goodpr),2)
class SingleGood(good_for_PDF):
    def __init__(self,gtp,gnm,gpr,per,num):
        super().__init__(gtp,gnm,gpr,per)
        self.number=num
        self.allmoney=round(float(self.goodpr)*float(self.number),2)
def outputExcel():
    goodlist=[]
    desktop_path = os.path.join(os.environ['USERPROFILE'], 'Desktop')
    def setGood():
        def pushsgood():
            a=SingleGood(*(tmp.split()), int(numEntry.get()))
            goodlist.append(a)
            a.description=str(a.goodtp)+" "+a.goodnm+" "+str(a.goodpr)+" "+str(a.number)+str(a.danwei)
            lb2.insert(END,a.description)
            numEntry.destroy()
            l1.destroy()
            b1.destroy()
        def pushagood():
            b=AreaGood(*(tmp.split()), float(longEntry.get()),float(widthEntry.get()))
            goodlist.append(b)
            b.description=str(b.goodtp)+" "+b.goodnm + " " + str(b.goodpr) + " " + str(b.number)+ "m²"
            lb2.insert(END, b.description)
            longEntry.destroy()
            widthEntry.destroy()
            l1.destroy()
            l2.destroy()
            b1.destroy()
        tmp=str(lb1.get(lb1.curselection(),lb1.curselection())[0])
        if tmp[0]=="1":
            numEntry = tk.Entry(window)
            numEntry.grid(row = 5, column = 1, padx = 10, pady = 5)
            l1=tk.Label(window, text="个数:")
            l1.grid(row=5)
            b1=tk.Button(window, text="确定", width=10, command=pushsgood)
            b1.grid(row=5, column=2, padx=10, pady=5)
        if tmp[0]=="0":
            longEntry = tk.Entry(window)
            longEntry.grid(row=5, column=1, padx=10, pady=5)
            widthEntry = tk.Entry(window)
            widthEntry.grid(row=6, column=1, padx=10, pady=5)
            l1=tk.Label(window, text="长度:")
            l1.grid(row=5)
            l2=tk.Label(window, text="宽度:")
            l2.grid(row=6)
            b1=tk.Button(window, text="确定", width=10, command=pushagood)
            b1.grid(row=5, column=2, padx=10, pady=5)



    def createExcel():
        customerobj=Customer(e1.get(),e2.get(),e3.get(),e4.get(),goodlist)
        sheetplot.setsheet(e1.get()+e3.get())
        sheetplot.completeSheet(customerobj)
        window.destroy()
    def cancelgood():
        tmp2=str(lb2.get(lb2.curselection(),lb2.curselection())[0])
        lb2.delete(lb2.curselection(),lb2.curselection())
        for i in goodlist:
            if i.description==tmp2:
                goodlist.remove(i)

    window = tk.Tk()
    tk.Label(window,text = "客户名:").grid(row=0)
    tk.Label(window,text = "联系方式:").grid(row=1)
    tk.Label(window, text="地址:").grid(row=2)
    tk.Label(window, text="日期:").grid(row=3)
    lb1=tk.Listbox(window)
    lb2 = tk.Listbox(window)
    allGoodList=[]
    with open(desktop_path+"\\商品表.txt", 'r') as file:
        for line in file:
            if line=="\n": continue
            lb1.insert("end",line)
            Mygood=good_for_PDF(*(line.split()))
            allGoodList.append(Mygood)
    lb1.grid(row=4, column=0,padx = 10, pady = 5)
    lb2.grid(row=4, column=1, padx=10, pady=5)
    e1 = tk.Entry(window)
    e2 = tk.Entry(window)
    e3 = tk.Entry(window)
    e4 = tk.Entry(window)
    e1.grid(row=0,column = 1,padx = 10,pady = 5)
    e2.grid(row=1,column = 1, padx = 10,pady = 5)
    e3.grid(row=2,column = 1, padx = 10,pady = 5)
    e4.grid(row=3,column = 1, padx = 10,pady = 5)
    tk.Button(window, text="输出", width=10, command=createExcel).grid(row=9, column=0,sticky="w",padx = 10, pady = 5)
    tk.Button(window, text="选择", width=10, command=setGood).grid(row=9, column=1, padx = 10, pady = 5)
    tk.Button(window, text="删除", width=10, command=cancelgood).grid(row=9,column=2,padx = 10, pady = 5)

    mainloop()