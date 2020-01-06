import tkinter as tk
from tkinter import filedialog
from tkinter import StringVar
import sys
from os import walk
import glob
import os
from pathlib import Path
from datetime import datetime
import OriginExt as Origin
from OriginExt import *
import time
import csv


class App(tk.Frame):
    def __init__(self, master=None):
        super().__init__(master)
        self.master = master
        self.pack()
        self.data_folder = StringVar()
        self.comp_order = StringVar()
        self.test_order = StringVar()
        self.create_widgets()

    def create_widgets(self):
        self.folder = tk.Entry(self, textvariable=self.data_folder, width=20)
        self.folder.grid(row=0, column=0, columnspan=2)
        self.browse = tk.Button(self)
        self.browse["text"] = "选择数据目录..."
        self.browse["command"] = self.select_folder
        self.browse.grid(row=0, column=2)
        self.comp_label = tk.Label(self, text="沉积文件序号")
        self.comp_label.grid(row=1, column=0)
        self.comp_order_input = tk.Entry(
            self, textvariable=self.comp_order, width=5)
        self.comp_order_input.grid(row=1, column=1)
        self.comp_label = tk.Label(self, text="测试文件序号")
        self.comp_label.grid(row=2, column=0)
        self.test_order_input = tk.Entry(
            self, textvariable=self.test_order, width=5)
        self.test_order_input.grid(row=2, column=1)
        self.process = tk.Button(self)
        self.process["text"] = "导入数据"
        self.process["width"] = 12
        self.process["command"] = self.import_data
        self.process.grid(row=1, column=2)

        self.plot = tk.Button(self)
        self.plot["text"] = "作图"
        self.plot["width"] = 12
        self.plot["command"] = self.plot_data
        self.plot.grid(row=2, column=2)

        self.comp_order.set("1")
        self.test_order.set("3")

    def select_folder(self):
        self.data_folder.set(filedialog.askdirectory())
        print(self.data_folder.get())
        return

    def import_data(self):
        # file_names = []
        # for (dirpath, dirnames, filenames) in walk(self.data_folder.get()):
        #     file_names.extend(filenames)
        # for file_name in file_names:
        #     id = str.split(str.split(file_name,'.')[1],']')[0]
        #     print(id)
        # files = list(filter(os.path.isfile, glob.glob(self.data_folder.get() + "/*")))
        # files.sort(key=lambda x: os.path.getmtime(x))
        # print(files)
        # print(len(files))
        self.process["text"] = "正在导入数据..."
        self.process["state"] = "disabled"
        self.process.update()

        self.oapp = Origin.Application()
        self.oapp.NewProject()

        self.datasheet_page = self.oapp.WorksheetPages.Add()
        self.datasheet_page.Name = "数据"
        self.datasheet = self.datasheet_page.Layers(0)
        self.datasheet.Name = "原始数据"
        self.datasheet_page.AddLayer()
        self.tenerysheet = self.datasheet_page.Layers(1)
        self.tenerysheet.Name = "三元"

        files = list(filter(os.path.isfile, Path(
            self.data_folder.get()).rglob("*.csv")))
        files.sort(key=lambda x: os.path.getmtime(x))
        previous = 0
        expset = 0
        i = 0
        # datasheet.PutCols(66)
        # print(datasheet.Columns)
        # print(len(datasheet.Columns))
        tech = ""
        el_info = ""
        mark_x = []
        el_infos = []
        el_list = []
        for file in files:
            current = str.split(str.split(file.name, '.')[1], ']')[0]
            if current != previous:
                expset = 1
                # print(file.name)
            else:
                expset += 1

            if str(expset) == self.comp_order.get():
                with open(file, newline='', encoding="utf8") as csvfile:
                    csvreader = csv.reader(csvfile)
                    el_infoline = 0
                    el_info = []
                    for row in csvreader:
                        if row[0].lstrip('-').replace('.', '', 1).isdigit():
                            break
                        if el_infoline > 0:
                            el_name = row[0][:2]
                            if el_name not in el_list:
                                el_list.append(el_name)
                            el_info.append(el_name + ":" + row[1])
                        elif "电解液" in row[0]:
                            el_infoline = csvreader.line_num + 1
                    el_infos.append(el_info)

            #print(str.split(str.split(file.name,'.')[1],']')[0]+":::"+datetime.fromtimestamp(os.path.getmtime(file)).strftime('%Y-%m-%d %H:%M:%S'))
            if str(expset) == self.test_order.get():
                x, y, data = [], [], []
                with open(file, newline='', encoding="utf8") as csvfile:
                    csvreader = csv.reader(csvfile)
                    startrow = 0
                    for row in csvreader:
                        if tech == "" and csvreader.line_num == 2:
                            tech = str.split(row[0], ':', 1)[1]
                        if startrow > 0:
                            x.append(float(row[0]))
                            y.append(float(row[1]))
                        elif row[0].lstrip('-').replace('.', '', 1).isdigit():
                            startrow = csvreader.line_num
                            x.append(float(row[0]))
                            y.append(float(row[1]))
                data.append(x)
                data.append(y)
                j = len(y) - 1
                while j >= 0:
                    if y[j] < 0.01:
                        break
                    j -= 1
                mark_x.append(x[j] + (x[j]-x[j-1]) /
                              (y[j]-y[j-1])*(0.01-y[j]))
                print(y[j])

                # datasheet.Columns(i+1).SetName("I")#Name属性不可更改
                self.datasheet.SetData(data, 0, i)

                self.datasheet.Columns(i).Type = COLTYPE_DESIGN_X
                self.datasheet.Columns(i+1).Type = COLTYPE_DESIGN_Y
                self.datasheet.Columns(i).DataFormat = Origin.COLFORMAT_NUMERIC
                self.datasheet.Columns(
                    i+1).DataFormat = Origin.COLFORMAT_NUMERIC
                self.datasheet.Columns(i).Comments = current
                self.datasheet.Columns(i+1).Comments = ("".join(el_info)).replace(":","")
                self.datasheet.Columns(i).LongName = "Voltage"
                self.datasheet.Columns(i+1).LongName = "Current"
                i += 2
            previous = current
        # self.oapp.Exit()
        el_cont = [[0.0 for h in range(int(i/2))] for k in range(len(el_list))]
        l = 0
        while l < len(el_infos):
            for els in el_infos[l]:
                el_name = els.split(":")[0]
                el_cont[el_list.index(el_name)][l] = float(els.split(":")[1])
            l += 1
        el_cont.append(mark_x)
        self.tenerysheet.SetData(el_cont, 0 , 0)
        self.tenerysheet.Columns(0).Type = COLTYPE_DESIGN_X
        self.tenerysheet.Columns(1).Type = COLTYPE_DESIGN_Y
        self.tenerysheet.Columns(2).Type = COLTYPE_DESIGN_Z
        self.tenerysheet.Columns(0).LongName = el_list[0]
        self.tenerysheet.Columns(1).LongName = el_list[1]
        self.tenerysheet.Columns(2).LongName = el_list[2]
        self.tenerysheet.Columns(3).LongName = "电势 at 10 mA"
        #self.tenerysheet.Columns()
        self.process["state"] = "normal"
        self.process["text"] = "导入数据"
        self.process.update()
        self.datasheet.PutLabelVisible(Origin.LABEL_COMMENTS)
        self.datasheet.PutLabelVisible(Origin.LABEL_LONG_NAME)
        self.datasheet.Name += tech
        self.oapp.Visible = 1
#        print(dir(datasheet.Columns(0)))
        return

    def plot_data(self):
        self.graphlayer = self.oapp.GraphPages.Add("Line").Layers(0)
        self.datarange = self.datasheet.NewDataRange(0, 0, -1, -1)
        #self.dataplot = self.graphlayer.DataPlots.Add(self.datarange)
        self.dataplot = self.graphlayer.AddPlot(self.datarange, 200)
        # Origin.GraphLayer.
        self.trirange = self.tenerysheet.NewDataRange(0, 0, -1, 2)
        #self.dataplot = self.graphlayer.DataPlots.Add(self.datarange)
        for i in range(300):
            self.trigraphlayer = self.oapp.GraphPages.Add("tenery").Layers(0)
            self.triplot = self.trigraphlayer.AddPlot(self.trirange, i)#245


root = tk.Tk()
root.title('数据处理')
app = App(master=root)
app.mainloop()
try:
    app.oapp.Exit()
except Exception:
    pass
sys.exit(0)
