import tkinter as tk
import command
window=tk.Tk()
window.title('出货单自动生成软件')
window.geometry('300x300')
b1=tk.Button(window,text="录入商品",width=20,height=3,command=command.input_good).pack()
b2=tk.Button(window,text="删除商品",width=20,height=3,command=command.delete_good).pack()
b3=tk.Button(window,text="输出表格",width=20,height=3,command=command.outputExcel).pack()
window.mainloop()


