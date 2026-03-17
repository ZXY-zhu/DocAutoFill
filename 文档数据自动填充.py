import tkinter as tk

###GUI界面
##主界面
root=tk.Tk()
root.title("文档自动填充")
root.geometry("400x200")

#点击函数
def on_begin_clicked():
    print("开始")
def on_pexcel_clicked():
    print("选择e")
    text_var_excel.set("test")
def on_pword_clicked():
    print("选择w")
    text_var_word.set("test")

#创建按钮
begin=tk.Button(root,text="开始",
                width=15,
                command=on_begin_clicked,activebackground="gray",
                cursor="hand2")
begin.place(x=140,y=150)
p_excel=tk.Button(root,text="选择e",
                  width=10,
                  command=on_pexcel_clicked,activebackground="gray",
                  cursor="hand2")
p_word=tk.Button(root,text="选择w",
                 width=10,
                 command=on_pword_clicked,activebackground="gray",
                 cursor="hand2")
p_excel.place(x=300,y=27)
p_word.place(x=300,y=87)

#创建标签
l_excel=tk.Label(root,text="Excel文件：")
l_excel.place(x=20,y=30)
l_word=tk.Label(root,text="Word模板：")
l_word.place(x=20,y=90)

#创建地址框
#Excel
f_excel=tk.Frame(root,
                 width=200,height=25,
                 relief="sunken",bd=2,bg="white")
f_excel.place(x=90,y=30)
f_excel.pack_propagate(False) #禁止自动改变边框大小
text_var_excel=tk.StringVar()
text_var_excel.set("")
text_address_excel=tk.Label(f_excel,textvariable=text_var_excel,
                            bg="white",fg="black",font=("宋体",12),
                            anchor="w",padx=3)
text_address_excel.pack(expand=True,fill="both")

#Word
f_word=tk.Frame(root,
                 width=200,height=25,
                 relief="sunken",bd=2,bg="white")
f_word.place(x=90,y=90)
f_word.pack_propagate(False) #禁止自动改变边框大小
text_var_word=tk.StringVar()
text_var_word.set("")
text_address_word=tk.Label(f_word,textvariable=text_var_word,
                            bg="white",fg="black",font=("宋体",12),
                           anchor="w",padx=3)
text_address_word.pack(expand=True,fill="both")


# 启动主事件循环
root.mainloop()