import os
import tkinter as tk
from tkinter import filedialog
import tkinter.messagebox
import matplotlib
matplotlib.use('TkAgg')
import numpy as np
from matplotlib.backends.backend_tkagg import FigureCanvasTkAgg
#from matplotlib.backends.backend_tkagg import NavigationToolbar2TkAgg
from matplotlib.font_manager import FontProperties
from matplotlib.figure import Figure
import openpyxl

class layout:
    def mainlayout(self):
        win = tk.Tk()

        win.title('边界流量计算')
        win.geometry('905x620')
        win.resizable(width=False, height=False)
        win.configure(background='cornflowerblue')

        # 定义第一个容器
        frame_left = tk.Frame(win, height=620, width=445, bg="WhiteSmoke", name="_left")
        # frame_left.place(relx=0.0, rely=0, relwidth=0.5, relheight=1)
        frame_left.place(x=5, y=0, height=620, width=445)

        self.leftlauyout(frame_left)

        # 定义第二个容器
        frame_right = tk.Frame(win, bg="WhiteSmoke", name="_right")
        # frame_right.place(relx=0.5, rely=0.0, relwidth=0.5, relheight=1)
        frame_right.place(x=455, y=0, height=620, width=445)

        self.rightlayout(frame_right)

        return win
    #左边frame
    def leftlauyout(self,frame_left):
        #第一行
        lable1 = tk.Label(frame_left,text = " 左边界:"  , bg="WhiteSmoke",font = ("宋体",13))
        lable1.place(x=7,y=2,height = 30, width=70)

        entry1 = tk.Entry(frame_left,borderwidth=2)
        entry1.place(x=80,y=2,height = 30, width=120)

        lable2 = tk.Label(frame_left,text = "  右边界:"  , bg="WhiteSmoke",font = ("宋体",13))
        lable2.place(x=220,y=2,height = 30, width=70)

        entry2 = tk.Entry(frame_left,borderwidth=2)
        entry2.place(x=300,y=2,height = 30, width=120)

        #第2行
        lable3 = tk.Label(frame_left,text = "   比降:"  , bg="WhiteSmoke",font = ("宋体",13))
        lable3.place(x=7,y=35,height = 30, width=70)

        entry3 = tk.Entry(frame_left,borderwidth=2)
        entry3.place(x=80,y=35,height = 30, width=120)

        lable4 = tk.Label(frame_left,text = "主槽糙率:"  , bg="WhiteSmoke",font = ("宋体",13))
        lable4.place(x=220,y=35,height = 30, width=70)

        entry4 = tk.Entry(frame_left,borderwidth=2)
        entry4.place(x=300,y=35,height = 30, width=120)


        #第3行
        lable5 = tk.Label(frame_left,text = "   水位:"  , bg="WhiteSmoke",font = ("宋体",13))
        lable5.place(x=7,y=68,height = 30, width=70)

        entry5 = tk.Entry(frame_left,borderwidth=2)
        entry5.place(x=80,y=68,height = 30, width=120)

        lable6 = tk.Label(frame_left,text = "    备注:"  , bg="WhiteSmoke",font = ("宋体",13))
        lable6.place(x=220,y=68,height = 30, width=70)

        entry6 = tk.Entry(frame_left,borderwidth=2)
        entry6.place(x=300,y=68,height = 30, width=120)
        #第4行
        lable7 = tk.Label(frame_left,text = "左滩边界:"  , bg="WhiteSmoke",font = ("宋体",13))
        lable7.place(x=7,y=101,height = 30, width=70)

        entry7 = tk.Entry(frame_left,borderwidth=2)
        entry7.place(x=80,y=101,height = 30, width=120)

        lable8 = tk.Label(frame_left,text = "右滩边界:"  , bg="WhiteSmoke",font = ("宋体",13))
        lable8.place(x=220,y=101,height = 30, width=70)

        entry8 = tk.Entry(frame_left,borderwidth=2)
        entry8.place(x=300,y=101,height = 30, width=120)

        #第5行
        lable9 = tk.Label(frame_left,text = "左滩糙率:"  , bg="WhiteSmoke",font = ("宋体",13))
        lable9.place(x=7,y=134,height = 30, width=70)

        entry9 = tk.Entry(frame_left,borderwidth=2)
        entry9.place(x=80,y=134,height = 30, width=120)

        lable10 = tk.Label(frame_left,text = "右滩糙率:"  , bg="WhiteSmoke",font = ("宋体",13))
        lable10.place(x=220,y=134,height = 30, width=70)

        entry10 = tk.Entry(frame_left,borderwidth=2)
        entry10.place(x=300,y=134,height = 30, width=120)
        #操作列
        frame_oprate = tk.Frame(frame_left,   bg="LightCyan")
        frame_oprate.place(x=0,y=134,height = 40, width=445)

        button1 = tk.Button(frame_left, text = "计算",background='DodgerBlue')
        button1.place(x=120,y=139,height = 30, width=50)

        button2 = tk.Button(frame_left, text = "保存计算结果与边界条件",background='DodgerBlue')
        button2.place(x=190,y=139,height = 30, width=150)

        #左槽
        group1 = tk.LabelFrame(frame_left, text="左滩", padx=5, pady=5)
        group1.place(x=10,y=172,height = 100, width=425)

        lable11 = tk.Label(group1,text = "左滩湿周:"  , bg="WhiteSmoke",font = ("宋体",13))
        lable11.place(x=7,y=3,height = 30, width=70)

        entry11 = tk.Entry(group1,borderwidth=2)
        entry11.place(x=80,y=3,height = 30, width=100)

        lable12 = tk.Label(group1,text = "左滩断面面积:"  , bg="WhiteSmoke",font = ("宋体",13))
        lable12.place(x=185,y=3,height = 30, width=120)

        entry12 = tk.Entry(group1,borderwidth=2)
        entry12.place(x=300,y=3,height = 30, width=105)


        lable13 = tk.Label(group1,text = "左滩水力半径:"  , bg="WhiteSmoke",font = ("宋体",13))
        lable13.place(x=7,y=39,height = 30, width=110)

        entry13 = tk.Entry(group1,borderwidth=2)
        entry13.place(x=120,y=39,height = 30, width=100)

        lable14 = tk.Label(group1,text = "左滩流量:"  , bg="WhiteSmoke",font = ("宋体",13))
        lable14.place(x=225,y=39,height = 30, width=80)

        entry14 = tk.Entry(group1,borderwidth=2)
        entry14.place(x=305,y=39,height = 30, width=100)

        # 右槽
        group2 = tk.LabelFrame(frame_left, text="右滩", padx=5, pady=5)
        group2.place(x=10, y=275, height=100, width=425)

        lable15 = tk.Label(group2, text="右滩湿周:", bg="WhiteSmoke", font=("宋体", 13))
        lable15.place(x=7, y=3, height=30, width=70)

        entry15 = tk.Entry(group2, borderwidth=2)
        entry15.place(x=80, y=3, height=30, width=100)

        lable16 = tk.Label(group2, text="右滩断面面积:", bg="WhiteSmoke", font=("宋体", 13))
        lable16.place(x=185, y=3, height=30, width=120)

        entry16 = tk.Entry(group2, borderwidth=2)
        entry16.place(x=300, y=3, height=30, width=105)

        lable17 = tk.Label(group2, text="右滩水力半径:", bg="WhiteSmoke", font=("宋体", 13))
        lable17.place(x=7, y=39, height=30, width=110)

        entry17 = tk.Entry(group2, borderwidth=2)
        entry17.place(x=120, y=39, height=30, width=100)

        lable18 = tk.Label(group2, text="右滩流量:", bg="WhiteSmoke", font=("宋体", 13))
        lable18.place(x=225, y=39, height=30, width=80)

        entry18 = tk.Entry(group2, borderwidth=2)
        entry18.place(x=305, y=39, height=30, width=100)

        # 主槽
        group3 = tk.LabelFrame(frame_left, text="主槽", padx=5, pady=5)
        group3.place(x=10, y=378, height=100, width=425)

        lable19 = tk.Label(group3, text="主槽湿周:", bg="WhiteSmoke", font=("宋体", 13))
        lable19.place(x=7, y=3, height=30, width=70)

        entry19 = tk.Entry(group3, borderwidth=2)
        entry19.place(x=80, y=3, height=30, width=100)

        lable20 = tk.Label(group3, text="主槽断面面积:", bg="WhiteSmoke", font=("宋体", 13))
        lable20.place(x=185, y=3, height=30, width=120)

        entry20 = tk.Entry(group3, borderwidth=2)
        entry20.place(x=300, y=3, height=30, width=105)

        lable21 = tk.Label(group3, text="主槽水力半径:", bg="WhiteSmoke", font=("宋体", 13))
        lable21.place(x=7, y=39, height=30, width=110)

        entry21 = tk.Entry(group3, borderwidth=2)
        entry21.place(x=120, y=39, height=30, width=100)

        lable22 = tk.Label(group3, text="主槽流量:", bg="WhiteSmoke", font=("宋体", 13))
        lable22.place(x=225, y=39, height=30, width=80)

        entry22 = tk.Entry(group3, borderwidth=2)
        entry22.place(x=305, y=39, height=30, width=100)

        # 主槽
        group3 = tk.LabelFrame(frame_left, text="主槽", padx=5, pady=5)
        group3.place(x=10, y=378, height=100, width=425)

        lable19 = tk.Label(group3, text="主槽湿周:", bg="WhiteSmoke", font=("宋体", 13))
        lable19.place(x=7, y=3, height=30, width=70)

        entry19 = tk.Entry(group3, borderwidth=2)
        entry19.place(x=80, y=3, height=30, width=100)

        lable20 = tk.Label(group3, text="主槽断面面积:", bg="WhiteSmoke", font=("宋体", 13))
        lable20.place(x=185, y=3, height=30, width=120)

        entry20 = tk.Entry(group3, borderwidth=2)
        entry20.place(x=300, y=3, height=30, width=105)

        lable21 = tk.Label(group3, text="主槽水力半径:", bg="WhiteSmoke", font=("宋体", 13))
        lable21.place(x=7, y=39, height=30, width=110)

        entry21 = tk.Entry(group3, borderwidth=2)
        entry21.place(x=120, y=39, height=30, width=100)

        lable22 = tk.Label(group3, text="主槽流量:", bg="WhiteSmoke", font=("宋体", 13))
        lable22.place(x=225, y=39, height=30, width=80)

        entry22 = tk.Entry(group3, borderwidth=2)
        entry22.place(x=305, y=39, height=30, width=100)

        # 总计
        group4 = tk.LabelFrame(frame_left, text="总计", padx=5, pady=5)
        group4.place(x=10, y=481, height=100, width=425)

        lable23 = tk.Label(group4, text="湿周:", bg="WhiteSmoke", font=("宋体", 13))
        lable23.place(x=7, y=3, height=30, width=70)

        entry23 = tk.Entry(group4, borderwidth=2)
        entry23.place(x=80, y=3, height=30, width=100)

        lable24 = tk.Label(group4, text="过水断面面积:", bg="WhiteSmoke", font=("宋体", 13))
        lable24.place(x=185, y=3, height=30, width=120)

        entry24 = tk.Entry(group4, borderwidth=2)
        entry24.place(x=300, y=3, height=30, width=105)

        lable25 = tk.Label(group4, text="水力半径:", bg="WhiteSmoke", font=("宋体", 13))
        lable25.place(x=7, y=39, height=30, width=110)

        entry25 = tk.Entry(group4, borderwidth=2)
        entry25.place(x=120, y=39, height=30, width=100)

        lable26 = tk.Label(group4, text="流量:", bg="WhiteSmoke", font=("宋体", 13))
        lable26.place(x=225, y=39, height=30, width=80)

        entry26 = tk.Entry(group4, borderwidth=2)
        entry26.place(x=305, y=39, height=30, width=100)

        #备注
        lable27 = tk.Label(frame_left,text=" 备注:", bg="WhiteSmoke",font = ("宋体",13))
        lable27.place(x=7,y=584,height = 30, width=70)

        entry27 = tk.Entry(frame_left,borderwidth=2)
        entry27.place(x=80,y=584,height = 30, width=350)

    #右边frame
    def rightlayout(self,frame_right):

        #canvas
        canvas1 = tk.Frame(frame_right, bg="WhiteSmoke", name="_right")
        canvas1.place(x=0, y=40, height=580, width=445)
        fig = Figure(figsize=(5, 4), dpi=100, facecolor='WhiteSmoke')
        canvas1.ax = fig.add_subplot(111)
        canvas1.canvas = FigureCanvasTkAgg(fig, master=canvas1)
        canvas1.canvas.get_tk_widget().pack(side=tk.TOP, fill=tk.BOTH, expand=1)
        canvas1.canvas._tkcanvas.pack(side=tk.TOP, fill=tk.BOTH, expand=1)
        # toolbar = NavigationToolbar2TkAgg(frame_right.canvas, frame_right)
        # toolbar.update()
        # 操作列
        button1 = tk.Button(frame_right,
                            text="选择断面数据",
                            command=lambda: self.openfile(frame_right,canvas1)
                            )
        button1.place(x=10, y=2, height=30, width=120)
        button1 = tk.Button(frame_right,
                            text="计算断面面积",
                            command=self.calarea
                            )
        button1.place(x=140, y=2, height=30, width=120)
        #self.draw(canvas1)
    #画图
    def draw(self,canvas1):
            '''绘图逻辑'''
            font_set = FontProperties(fname=r"c:\windows\fonts\simsun.ttc", size=12)

            # x = range(2, 26, 2)
            # y = [15, 13, 14.5, 17, 20, 25, 26, 26, 24, 22, 18, 15]
            # y1 = [17, 17, 17, 17, 17,17, 17, 17, 17, 17, 17, 17]
            # 数据在y周的位置是一个可迭代的对象
            # x轴 y轴的数据一起组成了所有要绘制出的图标
            # self.fig.clf() # 方式一：①清除整个Figure区域
            # self.ax = self.fig.add_subplot(111) # ②重新分配Axes区域
            canvas1.ax.clear() # 方式二：①清除原来的Axes区域
            # 添加描述信息
            canvas1.ax.set_xlabel('起点距', fontproperties=font_set )
            canvas1.ax.set_ylabel('高程', fontproperties=font_set )
            canvas1.ax.plot(self.x, self.y,linewidth=3,color='blue',marker='.',markerfacecolor='red',markersize=10)  # 传入x和y 通过plot绘制出折线图
            canvas1.ax.plot(self.x, self.y1,linewidth=3,color='red')  # 传入x和y 通过plot绘制出折线图
            canvas1.canvas.draw()
    #读取excel
    def openfile(self,frame_right,canvas1):
        my_filetypes = [('text files', '.xlsx')]#('all files', '.*'), ('text files', '.xls'),
        answer = filedialog.askopenfilename(parent=frame_right,
                                            initialdir=os.getcwd(),
                                            title="Please select a file:",
                                            filetypes=my_filetypes)
        try:
            wb = openpyxl.load_workbook(answer)
            ws1 = wb.get_sheet_by_name("断面")
            ws2 = wb.get_sheet_by_name("水位")
            self.waterlevel = ws2.cell(row=1, column=1).value
            # print(ws1.cell(row=1, column=2).value)
            self.x = []
            self.y = []
            self.y1 = []
            self.xy = []
            for num in range(2, ws1._current_row):
                self.x.append(ws1.cell(row=num, column=1).value)
                self.y.append(ws1.cell(row=num, column=2).value)
                self.y1.append(self.waterlevel)
                self.xy.append([ws1.cell(row=num, column=1).value,ws1.cell(row=num, column=2).value])
            self.draw(canvas1)
        except Exception as e:
            #print(e)  # 打印所有异常到屏幕
            tk.messagebox.showerror(title='错误', message="excel文档错误："+e)
    #计算面积
    def calarea(self):
        #获取所有交点
        allpoint=[];
        for i in range(len(self.y)):
            if(i>0 and (self.y[i]-self.waterlevel)*(self.y[i-1]-self.waterlevel)<=0):
                # 一般式 Ax+By+C=0
                a = self.y[i] - self.y[i-1]
                b = self.x[i-1] - self.x[i]
                c = self.x[i] * self.y[i-1] - self.x[i-1] * self.y[i]
                xpoint = ((0-b*self.waterlevel)-c)/a
                allpoint.append([xpoint,self.waterlevel])
        if(len(allpoint)<2):
            tk.messagebox.showerror(title='错误', message="交点少于两个，无法计算面积")
            return
        #获取相聚最远的两个相邻焦点
        index = 0
        maxlengh = 0;
        for i in range(len(allpoint)):
            if(i>0 and allpoint[i][0] -  allpoint[i-1][0]>maxlengh):
                index = i
                maxlengh = allpoint[i][0] -  allpoint[i-1][0]
        rightpoint = allpoint[index]#右交点
        leftpoint = allpoint[index-1]#左交点
        #将焦点加入xy数据
        self.xy.append(leftpoint)
        self.xy.append(rightpoint)
        self.xy.sort()
        area = 0#面积
        for i  in range(len(self.xy)):
            if(self.xy[i][0]>leftpoint[0] and self.xy[i][0] <= rightpoint[0] ):
                area += (self.xy[i][0]-self.xy[i-1][0])*(self.waterlevel- self.xy[i][1])
        tk.messagebox.showinfo(title='面积为', message=area)


win = layout()
win.mainlayout().mainloop()