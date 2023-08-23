Sub ProcessFile()

    Dim ws As Worksheet
    Dim InputPath As String
    Dim LastElement As String
    Dim OutputPath As String
    Dim LineContent As String
    Dim Elements() As String
    Dim FileContent As String
    Dim Lines() As String
    Dim i As Long

    ' Step 1: 获取指定Sheet的第一个单元格作为输入文件的路径
    Set ws = ThisWorkbook.Sheets("Sheet1") ' 更改Sheet名称如有必要
    InputPath = ws.Cells(1, 1).Value

    ' 输出文件的路径（这里我设为同一个文件夹下，名为"Output.txt"的文件）
    OutputPath = ThisWorkbook.Path & "\Output.txt"

    ' 使用 ADODB.Stream 读取文件内容
    Dim oStream As Object
    Set oStream = CreateObject("ADODB.Stream")
    oStream.Open
    oStream.Type = 1 'Binary
    oStream.LoadFromFile InputPath

    ' 转换为 UTF-8 编码
    oStream.Type = 2 'Text
    oStream.Charset = "utf-8"
    FileContent = oStream.ReadText

    oStream.Close
    Set oStream = Nothing

    ' 将文件内容分割为行
    Lines = Split(FileContent, vbCrLf)

    ' 打开输出文件
    Dim OutputFile As Integer
    OutputFile = FreeFile
    Open OutputPath For Output As #OutputFile

    For i = LBound(Lines) To UBound(Lines)
        LineContent = Lines(i)

        ' Step 3: 把每一行内容用逗号分隔，并且获取最后一个元素的内容
        Elements = Split(LineContent, ",")
        LastElement = Elements(UBound(Elements))

        ' Step 4: 判断最后一个元素的数据是否为日期类型，如果是的话就输出到新文件中
        If IsDate(LastElement) Then
            Print #OutputFile, LineContent
        End If
    Next i

    ' 关闭文件
    Close #OutputFile

    MsgBox "处理完毕，结果已保存在: " & OutputPath

End Sub










# -*- coding: utf-8 -*-

import sys
from PyQt5.QtWidgets import QApplication, QMainWindow, QCheckBox, QHBoxLayout
from PyQt5.QtCore import QObject, pyqtSignal, pyqtSlot, Qt
from myMainWindow import *
import csv
import pyqtgraph as pg


class Communicator(QObject):
    update_filtered_curves = pyqtSignal(int, bool)

class MyMainWindow(QMainWindow, Ui_MainWindow):
    def __init__(self, parent=None):
        super(MyMainWindow, self).__init__(parent)
        self.setupUi(self)
        self.communicator = Communicator()
        # self.communicator.update_ui.connect(self.update_ui)
        self.communicator.update_filtered_curves.connect(self.update_filtered_curves)
        # self.update_filtered_curves = pyqtSignal(int, bool)
        self.pushButton.clicked.connect(self.start)
        self.curves_names = ["A", "B", "C"]
        self.invisible_curves =[False,False,False]
        self.graphicsView.showGrid(x=True, y=True, alpha=0.5)
        self.graphicsView.addLegend(size=(150, 80))

        # vLine = pg.InfiniteLine(angle=90, movable=False, )
        # hLine = pg.InfiniteLine(angle=0, movable=False, )
        # self.graphicsView.addItem(vLine, ignoreBounds=True)
        # self.graphicsView.addItem(hLine, ignoreBounds=True)
        # vb = self.graphicsView.vb
        # def mouseMoved(evt):
        #     pos = evt[0]  ## using signal proxy turns original arguments into a tuple
        #     if self.graphicsView.sceneBoundingRect().contains(pos):
        #         mousePoint = vb.mapSceneToView(pos)
        #         index = int(mousePoint.x())
        #         pos_y = int(mousePoint.y())
        #         print(index)
        #         # if 0 < index < len(data.index):
        #         #     print(xdict[index], data['open'][index], data['close'][index])
        #         #     label.setHtml(
        #         #         "<p style='color:white'>日期：{0}</p><p style='color:white'>开盘：{1}</p><p style='color:white'>收盘：{2}</p>".format(
        #         #             xdict[index], data['open'][index], data['close'][index]))
        #         #     label.setPos(mousePoint.x(), mousePoint.y())
        #         vLine.setPos(mousePoint.x())
        #         hLine.setPos(mousePoint.y())
        #
        # proxy = pg.SignalProxy(self.graphicsView.scene().sigMouseMoved, rateLimit=60, slot=mouseMoved)

        self.buttons_group = []
        for i, (c_name, c_active) in enumerate(zip(self.curves_names, self.invisible_curves)):
            cb = QCheckBox(c_name)
            cb.setChecked(not c_active)  # активный чекбокс показывает False в invisible_curves

            # не знаю как иначе передать номер чек бокса внутрь обработчика, только лямбда и замыкание...
            #  причем пробросим i внутрь лямбды через kwargs, иначе у каждой лямбды будет последнее i счетчитка
            cb.toggled.connect(lambda value, i=i: self.update_checkbox(i, value))
            self.buttons_group.append(cb)
        l = QHBoxLayout()
        for b in self.buttons_group:
            l.addWidget(b)

        self.widget.setLayout(l)

    def start(self):


        with open('./csvfile.csv', 'r', newline='') as csvf:
            spamreader = csv.reader(csvf)  # 创建reader对象

            skip = []
            data = []
            # data[0] = []
            # data[1] = []
            # data[2] = []
            #
            global xdict, stringaxis, data1
            data1 = []
            data2 = []
            data3 = []
            for i in spamreader:
                # print(i)
                temp = '\t'.join(i)
                temp_list = temp.split("\t")
                # print(type(temp_list))
                if len(temp_list) > 0 and temp_list[0] != '' and not temp_list[0] in skip:
                    skip.append(temp_list[0])
                    data1.append(int(temp_list[2]))
                    data2.append(int(temp_list[3]))
                    data3.append(int(temp_list[4]))
            print(data1)
            print(data2)
            print(data3)
            xdict = dict(enumerate(skip))
            # axis_1 = [(i, list(skip)[i]) for i in range(0, len(skip), 5)]
            axis_1 = [(i, list(skip)[i]) for i in range(0, len(skip))]

            stringaxis = self.graphicsView.getAxis('bottom')
            stringaxis.setTicks([axis_1, xdict.items()])

            x = [i for i in range(0, len(skip))]
            columns_count = 3

            # plot.plot(x=list(xdict.keys()), y=data['open'].values, pen='r', name='开盘指数', symbolBrush=(255, 0, 0), )
            # plot.plot(x=list(xdict.keys()), y=data['close'].values, pen='g', name='收盘指数', symbolBrush=(0, 255, 0))
            self.curves = []
            # self.curves = [self.graphicsView.plot(x=list(xdict.keys()), y=data[i], pen=pg.intColor(i), name=self.curves_names[i]) for i in range(columns_count - 1)]
            self.curves.append( self.graphicsView.plot(x=list(xdict.keys()), y=data1, pen=pg.intColor(0), name=self.curves_names[0]))
            self.curves.append( self.graphicsView.plot(x=list(xdict.keys()), y=data2, pen=pg.intColor(1), name=self.curves_names[1]))
            self.curves.append( self.graphicsView.plot(x=list(xdict.keys()), y=data3, pen=pg.intColor(2), name=self.curves_names[2]))

    def update_checkbox(self, curve_number, value):
        self.communicator.update_filtered_curves.emit(curve_number, not value)  # значения чекбокса с флагами инверировано

    @pyqtSlot(int, bool)
    def update_filtered_curves(self, curve_number: int, set_invisible: bool):
        """ Слот на изменении чекбокса в ViewSettingsWindow """
        self.invisible_curves[curve_number] = set_invisible
        if set_invisible:
            self.curves[curve_number].clear()
        else :
            global xdict, stringaxis, data1
            self.curves[curve_number] = self.graphicsView.plot(x=list(xdict.keys()), y=data1, pen=pg.intColor(0), name="A")

if __name__ == "__main__":
    # 每一pyqt5应用程序必须创建一个应用程序对象。sys.argv参数是一个列表，从命令行输入参数。
    app = QApplication(sys.argv)
    myWin = MyMainWindow()
    # 显示在屏幕上
    myWin.show()
    # 系统exit()方法确保应用程序干净的退出
    # 的exec_()方法有下划线。因为执行是一个Python关键词。因此，exec_()代替
    sys.exit(app.exec_())
