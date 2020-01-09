import tkinter as tk
from tkinter import filedialog
from tkinter import StringVar
from tkinter import IntVar
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
import math


class App(tk.Frame):
    def __init__(self, master=None):
        super().__init__(master)
        self.master = master
        self.pack()
        self.data_folder = StringVar()
        self.comp_order = StringVar()
        self.test_order = StringVar()
        self.set_correct = IntVar()
        self.area = StringVar()
        self.resistance = StringVar()
        self.benchmark = StringVar()
        self.overpotent = StringVar()
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

        self.benchmark_label = tk.Label(self, text="性能比较线")
        self.benchmark_label.grid(row=3, column=0, sticky = "w")
        self.benchmark_input = tk.Entry(
            self, textvariable=self.benchmark, width=5)
        self.benchmark_input.grid(row=3, column=1)
        self.benchmark_label2 = tk.Label(self, text="mA （注意正负）", anchor="w")
        self.benchmark_label2.grid(row=3, column=2, sticky = "w")

        self.overpotent_label = tk.Label(self, text="过电势计算")
        self.overpotent_label.grid(row=4, column=0, sticky = "w")
        self.overpotent_input = tk.Entry(
            self, textvariable=self.overpotent, width=5)
        self.overpotent_input.grid(row=4, column=1)
        self.overpotent_label = tk.Label(self, text="V", anchor="w")
        self.overpotent_label.grid(row=4, column=2, sticky = "w")

        self.check_correct = tk.Checkbutton(self, variable=self.set_correct)
        self.check_correct["text"] = "修正原始数据"
        self.check_correct.grid(row=5,column=0)

        self.area_label = tk.Label(self, text="电极面积")
        self.area_label.grid(row=6, column=0, sticky = "w")
        self.area_input = tk.Entry(
            self, textvariable=self.area, width=5)
        self.area_input.grid(row=6, column=1)
        self.area_label2 = tk.Label(self, text="平方厘米", anchor="w")
        self.area_label2.grid(row=6, column=2, sticky = "w")

        
        self.resist_label = tk.Label(self, text="电解液电阻")
        self.resist_label.grid(row=7, column=0, sticky = "w")
        self.resist_input = tk.Entry(
            self, textvariable=self.resistance, width=5)
        self.resist_input.grid(row=7, column=1)
        self.resist_label2 = tk.Label(self, text="欧姆", anchor="w")
        self.resist_label2.grid(row=7, column=2, sticky = "w")




        #self.y_max = 0.0
        #self.y_min = 0.0
        self.x_max = 0.0
        self.x_min = 0.0
        self.mark_x_max = 0.0
        self.mark_x_min = 0.0
        self.cmark_x_max = 0.0
        self.cmark_x_min = 0.0

        self.comp_order.set("1")
        self.test_order.set("3")
        self.area.set("1")
        self.resistance.set("2")
        self.benchmark.set("10")
        self.overpotent.set("0")

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
        self.ternarysheet = self.datasheet_page.Layers(1)
        self.ternarysheet.Name = "原始数据三元"

        if self.set_correct.get() == 1:
            self.datasheet_page.AddLayer()
            self.cdatasheet = self.datasheet_page.Layers(2)
            self.cdatasheet.Name = "修正数据"

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
        cmark_x = []
        el_infos = []
        el_list = []
        x_range_set = False
        bench = float(self.benchmark.get())/1000
        resist = float(self.resistance.get())
        area = float(self.area.get())
        overpot = float(self.overpotent.get())

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

            # print(str.split(str.split(file.name,'.')[1],']')[0]+":::"+datetime.fromtimestamp(os.path.getmtime(file)).strftime('%Y-%m-%d %H:%M:%S'))
            if str(expset) == self.test_order.get():
                x, y, data = [], [], []
                cx, cy, cdata = [], [], [] #修正后的数据
                with open(file, newline='', encoding="utf8") as csvfile:
                    csvreader = csv.reader(csvfile)
                    startrow = 0
                    for row in csvreader:
                        if tech == "" and csvreader.line_num == 2:
                            tech = str.split(row[0], ':', 1)[1]
                        if not x_range_set and csvreader.line_num == 3:
                            self.x_min = float(row[1])
                        if not x_range_set and csvreader.line_num == 4:
                            self.x_max = float(row[1])
                            x_range_set = True
                        if startrow > 0:
                            x.append(float(row[0]))
                            y.append(float(row[1]))
                            # if float(row[0]) > self.x_max:
                            #     self.x_max=float(row[0])
                            # if float(row[0]) < self.x_min:
                            #     self.x_min=float(row[0])
                            # if float(row[1]) > self.y_max:
                            #     self.y_max=float(row[1])
                            # if float(row[1]) < self.y_min:
                            #     self.y_min=float(row[1])
                            if self.set_correct.get() == 1:
                                cx.append(float(row[0])-float(row[1])*resist+overpot)
                                cy.append(float(row[1])/area)
                        elif row[0].lstrip('-').replace('.', '', 1).isdigit():
                            startrow=csvreader.line_num
                            x.append(float(row[0]))
                            y.append(float(row[1]))
                            if self.set_correct.get() == 1:
                                cx.append(float(row[0])-float(row[1])*resist+overpot)
                                cy.append(float(row[1])/area)
                            # if float(row[0]) > self.x_max:
                            #     self.x_max=float(row[0])
                            # if float(row[0]) < self.x_min:
                            #     self.x_min=float(row[0])
                            # if float(row[1]) > self.y_max:
                            #     self.y_max=float(row[1])
                            # if float(row[1]) < self.y_min:
                            #     self.y_min=float(row[1])
                data.append(x)
                data.append(y)
                
                if self.set_correct.get() == 1:
                    cdata.append(cx)
                    cdata.append(cy)
                    cj=len(cy) - 1
                    while cj >= 0:
                        if cy[cj] < bench:
                            break
                        cj -= 1
                    cmark=cx[cj] + (cx[cj]-cx[cj-1]) / (cy[cj]-cy[cj-1])*(bench-cy[cj])
                    cmark_x.append(cmark)



                j=len(y) - 1
                while j >= 0:
                    if y[j] < bench:
                        break
                    j -= 1
                mark=x[j] + (x[j]-x[j-1]) / (y[j]-y[j-1])*(bench-y[j])
                mark_x.append(mark)
                # if mark > self.mark_x_max:
                #     self.mark_x_max = mark
                # if mark < self.mark_x_min:
                #     self.mark_x_min = mark
                # print(y[j])

                # datasheet.Columns(i+1).SetName("I")#Name属性不可更改
                self.datasheet.SetData(data, 0, i)

                self.datasheet.Columns(i).Type=COLTYPE_DESIGN_X
                self.datasheet.Columns(i+1).Type=COLTYPE_DESIGN_Y
                self.datasheet.Columns(i).DataFormat=Origin.COLFORMAT_NUMERIC
                self.datasheet.Columns(
                    i+1).DataFormat=Origin.COLFORMAT_NUMERIC
                self.datasheet.Columns(i).Comments=current
                self.datasheet.Columns(
                    i+1).Comments=("".join(el_info)).replace(":", "")
                self.datasheet.Columns(i).LongName="Voltage"
                self.datasheet.Columns(i+1).LongName="Current"

                if self.set_correct.get() == 1:
                    self.cdatasheet.SetData(cdata, 0, i)

                    self.cdatasheet.Columns(i).Type=COLTYPE_DESIGN_X
                    self.cdatasheet.Columns(i+1).Type=COLTYPE_DESIGN_Y
                    self.cdatasheet.Columns(i).DataFormat=Origin.COLFORMAT_NUMERIC
                    self.cdatasheet.Columns(
                        i+1).DataFormat=Origin.COLFORMAT_NUMERIC
                    self.cdatasheet.Columns(i).Comments=current
                    self.cdatasheet.Columns(
                        i+1).Comments=("".join(el_info)).replace(":", "")
                    self.cdatasheet.Columns(i).LongName="Voltage"
                    self.cdatasheet.Columns(i+1).LongName="Current"

                i += 2
            previous=current
        # self.oapp.Exit()
        # self.x_max = round_up(self.x_max,1)
        # self.x_min = round_down(self.x_min,1)
        # self.y_max = round_up(self.y_max,1)
        # self.y_min = round_down(self.y_min,1)
        if self.x_max < self.x_min:
            temp = self.x_max
            self.x_max = self.x_min
            self.x_min = temp

        self.mark_x_max = min([max(mark_x),self.x_max])
        self.mark_x_min = max([round_down(min(mark_x),2),self.x_min])
        temp_max = self.mark_x_min
        self.num_levels = 0
        while temp_max < self.mark_x_max:
            temp_max += 0.01
            self.num_levels += 1
        self.mark_x_max = round(temp_max,2)

        if self.set_correct.get() == 1:
            self.cmark_x_max = min([max(cmark_x),self.x_max])
            self.cmark_x_min = max([round_down(min(cmark_x),2),self.x_min])
            temp_cmax = self.cmark_x_min
            self.cnum_levels = 0
            while temp_cmax < self.cmark_x_max:
                temp_cmax += 0.01
                self.cnum_levels += 1
            self.cmark_x_max = round(temp_cmax,2)

        el_cont=[[0.0 for h in range(int(i/2))] for k in range(len(el_list))]
        l=0
        while l < len(el_infos):
            for els in el_infos[l]:
                el_name=els.split(":")[0]
                el_cont[el_list.index(el_name)][l]=float(els.split(":")[1])
            l += 1
        el_cont.append(mark_x)
        if self.set_correct.get() == 1:
            el_cont.append(cmark_x)
        self.ternarysheet.SetData(el_cont, 0, 0)
        self.ternarysheet.Columns(0).Type=COLTYPE_DESIGN_X
        self.ternarysheet.Columns(1).Type=COLTYPE_DESIGN_Y
        self.ternarysheet.Columns(2).Type=COLTYPE_DESIGN_Z
        self.ternarysheet.Columns(0).LongName=el_list[0]
        self.ternarysheet.Columns(1).LongName=el_list[1]
        self.ternarysheet.Columns(2).LongName=el_list[2]
        self.ternarysheet.Columns(3).LongName=f"电势 at {self.benchmark_input.get()} mA"
        if self.set_correct.get() == 1:
            self.ternarysheet.Columns(4).LongName=f"修正电势 at {self.benchmark_input.get()} mA"

        # self.ternarysheet.Columns()
        self.process["state"]="normal"
        self.process["text"]="导入数据"
        self.process.update()
        self.datasheet.PutLabelVisible(Origin.LABEL_COMMENTS)
        self.datasheet.PutLabelVisible(Origin.LABEL_LONG_NAME)
        self.datasheet.Name += tech
        self.oapp.Visible=1
#        print(dir(datasheet.Columns(0)))
        return

    def plot_data(self):
        self.graphlayer=self.oapp.GraphPages.Add("Line").Layers(0)
        self.datarange=self.datasheet.NewDataRange(0, 0, -1, -1)
        # self.dataplot = self.graphlayer.DataPlots.Add(self.datarange)
        self.dataplot=self.graphlayer.AddPlot(self.datarange, 200)
        # Origin.GraphLayer.
        color_high = 1
        color_low = 11
        if self.mark_x_max < 0:
            color_high = 11
            color_low = 1
        labtalk=";".join(["page.dimension.unit=1",
                            "page.dimension.width=7200",
                            "page.dimension.height=4800",
                            "layer.dimension.unit=1",
                            "layer.width=80",
                            "layer.height=80",
                            "layer.top=5",
                            "layer.left=16",
                            "layer.y.MajorTicks=11",
                            "layer.z.MajorTicks=11",
                            "layer.x.ticks=1",
                            "layer.y.ticks=1",
                            "layer.x.showaxes=3",
                            "layer.y.showaxes=3",
                            "layer.x2.ticks=0",
                            "layer.y2.ticks=0",
                            "layer.x.thickness=2",
                            "layer.x2.thickness=2",
                            "layer.y.thickness=2",
                            "layer.y2.thickness=2",
                            "layer.x.label.pt=24",
                            "layer.x.label.bold=1",
                            "layer.x2.label.pt=24",
                            "layer.x2.label.bold=1",
                            "layer.y.label.pt=24",
                            "layer.y.label.bold=1",
                            "layer.y2.label.pt=24",
                            "layer.y2.label.bold=1",
                            "xb.text$=\"\\b(%(?X))\"",
                            "xb.fsize=36",
                            "yl.text$=\"\\b(%(?Y))\"",
                            "yl.fsize=36",
                            "set %c -w 1500",

                            f"layer.x.from={self.x_min}",
                            f"layer.x.to={self.x_max}",
                            ])
        print(labtalk)
        self.graphlayer.Execute(labtalk)

        if self.set_correct.get() == 1:
            self.cgraphlayer=self.oapp.GraphPages.Add("Line").Layers(0)
            self.cdatarange=self.cdatasheet.NewDataRange(0, 0, -1, -1)
            # self.dataplot = self.graphlayer.DataPlots.Add(self.datarange)
            self.cdataplot=self.cgraphlayer.AddPlot(self.cdatarange, 200)
            # Origin.GraphLayer.
            labtalk=";".join(["page.dimension.unit=1",
                                "page.dimension.width=7200",
                                "page.dimension.height=4800",
                                "layer.dimension.unit=1",
                                "layer.width=80",
                                "layer.height=80",
                                "layer.top=5",
                                "layer.left=16",
                                "layer.y.MajorTicks=11",
                                "layer.z.MajorTicks=11",
                                "layer.x.ticks=1",
                                "layer.y.ticks=1",
                                "layer.x.showaxes=3",
                                "layer.y.showaxes=3",
                                "layer.x2.ticks=0",
                                "layer.y2.ticks=0",
                                "layer.x.thickness=2",
                                "layer.x2.thickness=2",
                                "layer.y.thickness=2",
                                "layer.y2.thickness=2",
                                "layer.x.label.pt=24",
                                "layer.x.label.bold=1",
                                "layer.x2.label.pt=24",
                                "layer.x2.label.bold=1",
                                "layer.y.label.pt=24",
                                "layer.y.label.bold=1",
                                "layer.y2.label.pt=24",
                                "layer.y2.label.bold=1",
                                "xb.text$=\"\\b(%(?X))\"",
                                "xb.fsize=36",
                                "yl.text$=\"\\b(%(?Y))\"",
                                "yl.fsize=36",
                                "set %c -w 1500",

                                f"layer.x.from={self.x_min}",
                                f"layer.x.to={self.x_max}",
                                ])
            print(labtalk)
            self.cgraphlayer.Execute(labtalk)

        self.trirange=self.ternarysheet.NewDataRange(0, 0, -1, 2)
        # self.dataplot = self.graphlayer.DataPlots.Add(self.datarange)
        self.tripage=self.oapp.GraphPages.Add("ternary")
        self.trigraphlayer=self.tripage.Layers(0)
        self.triplot=self.trigraphlayer.AddPlot(self.trirange, 245)
        labtalk=";".join(["page.dimension.unit=1",
                            "page.dimension.width=7200",
                            "page.dimension.height=4800",
                            "layer.x.MajorTicks=11",
                            "layer.y.MajorTicks=11",
                            "layer.z.MajorTicks=11",
                            "layer.x.MinorTicks=0",
                            "layer.y.MinorTicks=0",
                            "layer.z.MinorTicks=0",

                            "set %c -k 17",
                            "set %c -z 20",
                            "set %c -cse 102",
                            "set %c -cset 2",

                            "set %c -cpal rainbow balanced",
                            f"layer.cmap.numMajorLevels={self.num_levels}",
                            "layer.cmap.color7=14",
                            f"layer.cmap.colorLow={color_low}",
                            f"layer.cmap.colorHigh={color_high}",
                            f"layer.cmap.zmax={self.mark_x_max}",
                            f"layer.cmap.zmin={self.mark_x_min}",
                            "layer.cmap.colormixmode=1",
                            "layer.cmap.fill(2)",
                            "layer.cmap.SetLevels()",
                            "layer.cmap.updateScale()",
                            "spectrum",
                            "spectrum1.top=0",
                            "spectrum1.labels.autodisp=0",
                            "spectrum1.labels.decplaces=2",
                            "spectrum1.width=1000",
                            "spectrum1.height=4000",
                            "spectrum1.top=200",
                            "spectrum1.left=6000",
                            "spectrum1.barthick=400",
                            "label -r legend",
                            
                            "layer.x.label.pt=24",
                            "layer.y2.label.pt=24",                           
                            "layer.z.label.pt=24",
                            "layer.x.label.bold=1",
                            "layer.y2.label.bold=1",
                            "layer.z.label.bold=1",
                            "layer.x.thickness=2",
                            "layer.y2.thickness=2",
                            "layer.z.thickness=2",
                            
                            "layer.x.ticks=0",
                            "layer.y2.ticks=0",
                            "layer.z.ticks=0",
                            "layer.z.label.offsetV = -90",
                            "layer.z.label.offsetH = -50",
                            "layer.y2.label.offsetH = 50",
                            "layer.x.label.offsetV = -50",
                            #"sec -p 5",

                            "xb.text$=\"\\b(%(?X))\"",
                            "xb.fsize=36",

                            "yr.text$=\"\\b(%(?Y))\"",
                            "yr.fsize=36",

                            "zf.text$=\"\\b(%(?Z))\"",
                            "zf.fsize=36",

                            "xb.rotate=0",
                            "yr.rotate=0",
                            "zf.rotate=0",

                            "sec -p 0.5",
                            "xb.top=4300",
                            "xb.left=5300",
                            "yr.top=200",
                            "yr.left=3000",
                            "zf.top=4300",
                            "zf.left=600",

                            ])
        print(labtalk)
        self.trigraphlayer.Execute(labtalk)

        if self.set_correct.get() == 1:
            self.ctrirange=self.ternarysheet.NewDataRange(0, 0, -1, 2)
            # self.dataplot = self.graphlayer.DataPlots.Add(self.datarange)
            self.ctripage=self.oapp.GraphPages.Add("ternary")
            self.ctrigraphlayer=self.ctripage.Layers(0)
            self.ctriplot=self.ctrigraphlayer.AddPlot(self.ctrirange, 245)
            labtalk=";".join(["page.dimension.unit=1",
                                "page.dimension.width=7200",
                                "page.dimension.height=4800",
                                "layer.x.MajorTicks=11",
                                "layer.y.MajorTicks=11",
                                "layer.z.MajorTicks=11",
                                "layer.x.MinorTicks=0",
                                "layer.y.MinorTicks=0",
                                "layer.z.MinorTicks=0",

                                "set %c -k 17",
                                "set %c -z 20",
                                "set %c -cse 103",
                                "set %c -cset 2",

                                "set %c -cpal rainbow balanced",
                                f"layer.cmap.numMajorLevels={self.cnum_levels}",
                                "layer.cmap.color7=14",
                                f"layer.cmap.colorLow={color_low}",
                                f"layer.cmap.colorHigh={color_high}",
                                f"layer.cmap.zmax={self.cmark_x_max}",
                                f"layer.cmap.zmin={self.cmark_x_min}",
                                "layer.cmap.colormixmode=1",
                                "layer.cmap.fill(2)",
                                "layer.cmap.SetLevels()",
                                "layer.cmap.updateScale()",
                                "spectrum",
                                "spectrum1.top=0",
                                "spectrum1.labels.autodisp=0",
                                "spectrum1.labels.decplaces=2",
                                "spectrum1.width=1000",
                                "spectrum1.height=4000",
                                "spectrum1.top=200",
                                "spectrum1.left=6000",
                                "spectrum1.barthick=400",
                                "label -r legend",
                                
                                "layer.x.label.pt=24",
                                "layer.y2.label.pt=24",                           
                                "layer.z.label.pt=24",
                                "layer.x.label.bold=1",
                                "layer.y2.label.bold=1",
                                "layer.z.label.bold=1",
                                "layer.x.thickness=2",
                                "layer.y2.thickness=2",
                                "layer.z.thickness=2",
                                
                                "layer.x.ticks=0",
                                "layer.y2.ticks=0",
                                "layer.z.ticks=0",
                                "layer.z.label.offsetV = -90",
                                "layer.z.label.offsetH = -50",
                                "layer.y2.label.offsetH = 50",
                                "layer.x.label.offsetV = -50",
                                #"sec -p 5",

                                "xb.text$=\"\\b(%(?X))\"",
                                "xb.fsize=36",

                                "yr.text$=\"\\b(%(?Y))\"",
                                "yr.fsize=36",

                                "zf.text$=\"\\b(%(?Z))\"",
                                "zf.fsize=36",

                                "xb.rotate=0",
                                "yr.rotate=0",
                                "zf.rotate=0",

                                "sec -p 0.5",
                                "xb.top=4300",
                                "xb.left=5300",
                                "yr.top=200",
                                "yr.left=3000",
                                "zf.top=4300",
                                "zf.left=600",

                                ])
            print(labtalk)
            self.ctrigraphlayer.Execute(labtalk)

def round_up(n, decimals=0): 
    multiplier = 10 ** decimals 
    return math.ceil(n * multiplier) / multiplier

def round_down(n, decimals=0):
    multiplier = 10 ** decimals
    return math.floor(n * multiplier) / multiplier

root=tk.Tk()
root.title('数据处理')
app=App(master=root)
app.mainloop()
try:
    app.oapp.Exit()
except Exception:
    pass
sys.exit(0)
