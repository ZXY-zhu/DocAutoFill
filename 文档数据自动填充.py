import tkinter as tk
from tkinter import filedialog, ttk, messagebox
import pandas as pd
import os
import docxtpl as doc
from docxtpl import DocxTemplate
import win32com.client

from matplotlib.pyplot import title


class App:
    def __init__(self,root):
        self.root=root
        self.e_file_path=""
        self.w_file_path=""

        self.text_var_excel=tk.StringVar()
        self.text_var_word=tk.StringVar()

        self.setup_ui()

    # 点击函数
    def on_begin_clicked(self):
        final_w_file_path = self.w_file_path

        if not os.path.exists(self.e_file_path):
            self.setup_warn("❌ 错误：Excel文件路径无效！")
            return
        if not os.path.exists(self.w_file_path):
            self.setup_warn("❌ 错误：Word文件路径无效！")
            return

        try:
            df=pd.read_excel(self.e_file_path)
            file_dir=os.path.dirname(self.e_file_path)
            new_folder=os.path.join(file_dir,"处理结果") #将目录路径和文件夹名组合成完整的文件夹路径
            if not os.path.exists(new_folder):
                os.mkdir(new_folder)

            #使用win32com.client处理兼容问题
            if self.w_file_path.endswith(".doc"):
                zhuan = tk.Toplevel(self.root)
                zhuan.title("提示")
                zhuan.geometry("250x100")
                wl = tk.Label(zhuan, text="检测到模板为doc文件，\n正在转化为docx文件……")
                wl.place(x=60, y=10)

                #不确定进度条
                progress_zhuan=ttk.Progressbar(zhuan, length=200, mode='indeterminate')
                progress_zhuan.place(x=25, y=50)
                progress_zhuan.start()

                # 窗口置顶
                zhuan.transient(self.root)
                zhuan.grab_set()

                # 窗口居中显示
                zhuan.update_idletasks()
                x = (zhuan.winfo_screenwidth() - zhuan.winfo_width()) // 2
                y = (zhuan.winfo_screenheight() - zhuan.winfo_height()) // 2
                zhuan.geometry(f"+{x}+{y}")

                #启动word
                word=win32com.client.Dispatch("Word.Application")
                word.Visible=False
                #打开.doc
                doc=word.Documents.Open(os.path.abspath(self.w_file_path))
                #创建临时.docx路径
                final_w_file_path=os.path.join(new_folder,"doc转化docx后模板.docx")
                #保存为.docx
                doc.SaveAs(os.path.abspath(final_w_file_path),FileFormat=16)
                doc.Close()
                word.Quit()

                # 转换窗口完成后销毁
                zhuan.destroy()
                #zhuan.update()

            doing=tk.Toplevel(self.root)
            doing.title("提示")
            doing.geometry("250x100")
            wl=tk.Label(doing, text="正在生成中……")
            wl.place(x=80, y=20)

            # 确定进度条
            progress_doing = ttk.Progressbar(doing, length=300, mode='determinate')
            total_rows = len(df)
            progress_doing.place(x=20, y=50)
            progress_doing['maximum'] = total_rows

            # 窗口置顶
            doing.transient(self.root)
            doing.grab_set()

            # 窗口居中显示
            doing.update_idletasks()
            x = (doing.winfo_screenwidth() - doing.winfo_width()) // 2
            y = (doing.winfo_screenheight() - doing.winfo_height()) // 2
            doing.geometry(f"+{x}+{y}")
            #使用docxtpl进行模板替换操作
            tpl=DocxTemplate(final_w_file_path) #加载Word模板

            for index,row in df.iterrows():
                context=row.to_dict()
                first_value=df.iloc[index,0]
                tpl.render(context)
                tpl.save(os.path.join(new_folder,f"{first_value}.docx"))
                # 更新进度条
                progress_doing['value'] = index + 1
                root.update_idletasks()

            if doing.winfo_exists():
                doing.destroy()
            messagebox.showinfo("成功", f"处理完成！\n文件已保存至：{new_folder}")

        except Exception as e:
            print("详细错误:", e)
            self.setup_warn("❌ 处理文件时出错")

    def on_pexcel_clicked(self):
        self.e_file_path = tk.filedialog.askopenfilename(title="选择Excel文件",
                                                    filetypes=[("Excel文件", "*.xlsx;*.xls")])  # 返回字符串路径
        self.text_var_excel.set(self.e_file_path)

    def on_pword_clicked(self):
        self.w_file_path = tk.filedialog.askopenfilename(title="选择Word模板",
                                                    filetypes=[("Word文件", "*.docx;*.doc")])
        self.text_var_word.set(self.w_file_path)

    def setup_ui(self):
        # 创建按钮
        begin = tk.Button(self.root, text="开始",
                          width=15,
                          command=self.on_begin_clicked, activebackground="gray",
                          cursor="hand2")
        begin.place(x=140, y=150)
        p_excel = tk.Button(self.root, text="选择",
                            width=10,
                            command=self.on_pexcel_clicked, activebackground="gray",
                            cursor="hand2")
        p_word = tk.Button(self.root, text="选择",
                           width=10,
                           command=self.on_pword_clicked, activebackground="gray",
                           cursor="hand2")
        p_excel.place(x=300, y=27)
        p_word.place(x=300, y=87)

        # 创建标签
        l_excel = tk.Label(self.root, text="Excel文件：")
        l_excel.place(x=20, y=30)
        l_word = tk.Label(self.root, text="Word模板：")
        l_word.place(x=20, y=90)

        # 创建地址框
        # Excel
        f_excel = tk.Frame(self.root,
                           width=200, height=25,
                           relief="sunken", bd=2, bg="white")
        f_excel.place(x=90, y=30)
        f_excel.pack_propagate(False)  # 禁止自动改变边框大小

        text_address_excel = tk.Label(f_excel, textvariable=self.text_var_excel,
                                      bg="white", fg="black", font=("宋体", 12),
                                      anchor="w", padx=3)
        text_address_excel.pack(expand=True, fill="both")

        # Word
        f_word = tk.Frame(self.root,
                          width=200, height=25,
                          relief="sunken", bd=2, bg="white")
        f_word.place(x=90, y=90)
        f_word.pack_propagate(False)  # 禁止自动改变边框大小

        text_address_word = tk.Label(f_word, textvariable=self.text_var_word,
                                     bg="white", fg="black", font=("宋体", 12),
                                     anchor="w", padx=3)
        text_address_word.pack(expand=True, fill="both")

    def setup_warn(self,wl_text):
        warn = tk.Toplevel(self.root)
        warn.title("错误")
        warn.geometry("250x100")
        wl = tk.Label(warn, text=wl_text)
        wl.place(x=40, y=30)

        # 窗口置顶
        warn.transient(self.root)
        warn.grab_set()

        # 窗口居中显示
        warn.update_idletasks()
        x = (warn.winfo_screenwidth() - warn.winfo_width()) // 2
        y = (warn.winfo_screenheight() - warn.winfo_height()) // 2
        warn.geometry(f"+{x}+{y}")



if __name__ == '__main__':
    ##主界面
    root = tk.Tk()
    root.title("文档自动填充")
    root.geometry("400x200")
    # 窗口居中显示
    root.update_idletasks()
    x = (root.winfo_screenwidth() - root.winfo_width()) // 2
    y = (root.winfo_screenheight() - root.winfo_height()) // 2
    root.geometry(f"+{x}+{y}")

    app = App(root)

    # 启动主事件循环
    root.mainloop()