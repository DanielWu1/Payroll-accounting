import time
import tkinter as tk
from tkinter import filedialog, messagebox
from newclass2 import combine as nc
from newclass import collect as na

class newclassGUI:

    def __init__(self):
        self.start_wn = tk.Tk()
        self.enter_wn = tk.Tk()
        self.fileLocation = ''
        self.counterjiegou = None
        self.counterkaoqin = None
        self.counterbutie = None
        self.loading_index = 0



    def center(self, top, width = 500, height = 300):
        top.update_idletasks()
        w = top.winfo_screenwidth()
        h = top.winfo_screenheight()
        x = (w - width) // 2
        y = (h - height) // 2
        '''format = width x height + left_top_width + left_top_height'''
        top.geometry("{}x{}+{}+{}".format(width, height, x, y))


    def startpage(self):
        self.center(self.start_wn)
        self.enter_wn.withdraw()
        self.start_wn.title(" 工资计算0.1 ")

        startLabel = tk.Label(self.start_wn, text=" 欢迎使用李炳臻的小程序,\n如果本程序还在测试阶段，\n有任何问题请及时联系我", bd=1, relief=tk.RAISED, font = "times 12", width=40, height=3, anchor = tk.N)
        startLabel.pack()

        choseFile = tk.Label(self.start_wn,text=" 请点击下方按钮选择您的Excel文件\n(请注意本程序暂时只支持读取xlsx文件)", bd=1, relief=tk.RAISED, font = "times 12", width=40, height =3,anchor = tk.N)
        choseFileButton = tk.Button(self.start_wn, text = "选择您的文件", command=self.searchFile)
        choseFile.pack()
        choseFileButton.place(x = 200, y = 200)

        self.start_wn.protocol("WM_DELETE_WINDOW", self.close_control1)
        self.start_wn.mainloop()

    def searchFile(self):
        self.start_wn.withdraw()
        self.fileLocation = tk.filedialog.askopenfilename(filetypes = (("Excel file", "*.xls"), ("all files", "*.*")))
        if self.fileLocation != "":
            self.reading_wn = tk.Tk()
            self.reading_wn.title("读取界面")
            self.center(self.reading_wn, width = 360, height = 60)
            loading_label = tk.Label(self.reading_wn, text = "读取文件中...")
            loading_label.place(x = 60, y = 15)
            self.reading_wn.after(100, self.enterPage)
            self.reading_wn.mainloop()
        else: self.start_wn.deiconify()

    def enterPage(self):

        self.counterjiegou = na(self.fileLocation, '结构')
        self.counterkaoqin = na(self.fileLocation, '考勤')
        self.counterbutie = na(self.fileLocation, '补贴')
        self.counterjiegou = self.counterjiegou.run_jiegou()
        self.counterkaoqin = self.counterkaoqin.run_kaoqin()
        self.counterbutie = self.counterbutie.run_butie()
        #sheet_name_list = self.counter.get_sheet_names()

        self.center(self.enter_wn)
        self.enter_wn.title("生成界面")
        L1 = tk.Label(self.enter_wn, text="请输入生成工资表格的新名称")
        L1.place(relx = 0.0, rely = 0.1)
        self.E1 = tk.Entry(self.enter_wn, bd=5)
        self.E1.place(relx = 0.5, rely = 0.1)
        self.exlname = self.E1.get()
        #L2 = tk.Label(self.enter_wn, text="请选择要读取的列表")
        #L2.place(relx = 0.0, rely = 0.3)
        #self.E2 = tk.Listbox(self.enter_wn, selectmode = tk.SINGLE)

        '''index = 1
        for name in sheet_name_list:
            self.E2.insert(index, name)
            index += 1
        self.E2.place(relx = 0.5, rely = 0.3)'''
        returnButton = tk.Button(self.enter_wn, text ="返回", command = self.windowControl)
        returnButton.place(relx = 0.3, rely =0.9)
        enterButton = tk.Button(self.enter_wn, text = "确认", command = self.conform)
        enterButton.place(relx = 0.7, rely = 0.9)

        self.reading_wn.withdraw()
        self.reading_wn.destroy()

        self.enter_wn.deiconify()
        self.enter_wn.protocol("WM_DELETE_WINDOW", self.close_control1)
        self.enter_wn.mainloop()

    def close_control1(self):
        self.start_wn.destroy()
        self.start_wn.quit()
        self.enter_wn.destroy()
        self.enter_wn.quit()

    def windowControl(self):
        self.enter_wn.withdraw()
        self.start_wn.deiconify()

    def loading_str(self):
        loading_list = ["loading.", "loading..", "loading...", "loading....", "loading....."]
        self.loading_index += 1
        if self.loading_index == 5: self.loading_index = 0
        return loading_list[self.loading_index]


    def conform(self):
        self.enter_wn.withdraw()
        self.loading_wn = tk.Tk()
        self.loading_wn.title("生成界面")
        self.center(self.loading_wn, width = 360, height = 60)
        loading_label = tk.Label(self.loading_wn, text = "生成中...")
        loading_label.place(x = 60, y = 15)
        self.loading_wn.after(50, self.back_end_program)
        self.loading_wn.mainloop()


    def back_end_program(self):
        #self.counter.set_sheet_name(str(self.E2.get(self.E2.curselection())))
        #self.counter.set_output_sheet(str(self.E1.get()))
        #self.counter.set()
        #self.counter.run()
        self.all_people_map_final = nc(self.counterjiegou,self.counterkaoqin,self.counterbutie)
        self.all_people_map_final.run()
        self.loading_wn.withdraw()
        self.loading_wn.destroy()
        tk.messagebox.showinfo("运行结束", "您的数据统计完毕")
        self.close_control1()


x = newclassGUI()

x.startpage()