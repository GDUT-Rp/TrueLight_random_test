# -*- coding: utf-8 -*-
# -*- coding: utf-8 -*-

# Form implementation generated from reading ui file 'zuoweiV1.0.ui'
#
# Created by: PyQt5 UI code generator 5.9.1
#
# WARNING! All changes made in this file will be lost!

import sys
import wx
import os
from PyQt5 import QtCore, QtGui, QtWidgets
from PyQt5.QtWidgets import QFileDialog, QAction, QTextEdit
import random
import time
import pandas as pd
import numpy as np
from xlrd import open_workbook
from xlutils.copy import copy
from openpyxl import Workbook
import xlwt

import Control.make_excel


class Ui_index_frame(object):
    def setupUi(self, index_frame):
        index_frame.setObjectName("index_frame")
        index_frame.resize(1050, 650)
        font = QtGui.QFont()
        font.setPointSize(12)
        index_frame.setFont(font)
        self.title_label = QtWidgets.QLabel(index_frame)
        self.title_label.setGeometry(QtCore.QRect(400, 50, 300, 30))
        font = QtGui.QFont()
        font.setFamily("微软雅黑")
        font.setPointSize(16)
        font.setBold(True)
        font.setWeight(75)
        self.title_label.setFont(font)
        self.title_label.setObjectName("title_label")
        self.choose_namexls_Button = QtWidgets.QPushButton(index_frame)
        self.choose_namexls_Button.setGeometry(QtCore.QRect(110, 120, 270, 35))
        self.choose_namexls_Button.setObjectName("choose_namexls_Button")
        self.choose_namexls_Button.clicked.connect(self.open_choose_file)
        self.name_address_lineEdit = QtWidgets.QLineEdit(index_frame)
        self.name_address_lineEdit.setGeometry(QtCore.QRect(390, 120, 620, 35))
        self.name_address_lineEdit.setObjectName("name_address_lineEdit")
        self.one_label = QtWidgets.QLabel(index_frame)
        self.one_label.setGeometry(QtCore.QRect(60, 120, 50, 35))
        font = QtGui.QFont()
        font.setPointSize(14)
        self.one_label.setFont(font)
        self.one_label.setObjectName("one_label")
        self.two_label = QtWidgets.QLabel(index_frame)
        self.two_label.setGeometry(QtCore.QRect(60, 190, 50, 35))
        font = QtGui.QFont()
        font.setPointSize(14)
        self.two_label.setFont(font)
        self.two_label.setObjectName("two_label")
        self.three_label = QtWidgets.QLabel(index_frame)
        self.three_label.setGeometry(QtCore.QRect(60, 260, 50, 35))
        font = QtGui.QFont()
        font.setPointSize(14)
        self.three_label.setFont(font)
        self.three_label.setObjectName("three_label")
        self.number_class_label = QtWidgets.QLabel(index_frame)
        self.number_class_label.setGeometry(QtCore.QRect(310, 190, 150, 35))
        self.number_class_label.setObjectName("number_class_label")
        self.lineEdit = QtWidgets.QLineEdit(index_frame)
        self.lineEdit.setGeometry(QtCore.QRect(450, 190, 100, 35))
        self.lineEdit.setObjectName("lineEdit")
        self.lineEdit.setText('8')
        self.grade_comboBox = QtWidgets.QComboBox(index_frame)
        self.grade_comboBox.setGeometry(QtCore.QRect(640, 190, 100, 35))
        self.grade_comboBox.setObjectName("grade_comboBox")
        self.grade_comboBox.addItem("")
        self.grade_comboBox.addItem("")
        self.grade_comboBox.addItem("")
        self.grade_comboBox.addItem("")
        self.grade_comboBox.addItem("")
        self.grade_comboBox.addItem("")
        self.grade_comboBox.addItem("")
        self.grade_label = QtWidgets.QLabel(index_frame)
        self.grade_label.setGeometry(QtCore.QRect(580, 190, 50, 35))
        self.grade_label.setObjectName("grade_label")
        self.test_name_comboBox = QtWidgets.QComboBox(index_frame)
        self.test_name_comboBox.setGeometry(QtCore.QRect(880, 190, 100, 35))
        self.test_name_comboBox.setObjectName("test_name_comboBox")
        self.test_name_comboBox.addItem("")
        self.test_name_comboBox.addItem("")
        self.test_name_comboBox.addItem("")
        self.test_name_comboBox.addItem("")
        self.test_name_comboBox.addItem("")
        self.test_name_comboBox.addItem("")
        self.test_name_comboBox.addItem("")
        self.test_name_comboBox.addItem("")
        self.test_name_comboBox.addItem("")
        self.test_name_comboBox.addItem("")
        self.test_name_comboBox.addItem("")
        self.test_name_comboBox.addItem("")
        self.test_name_comboBox.addItem("")
        self.test_name_comboBox.addItem("")
        self.test_name_comboBox.addItem("")
        self.test_name_comboBox.addItem("")
        self.test_name_comboBox.addItem("")
        self.test_name_comboBox.addItem("")
        self.test_name_comboBox.addItem("")
        self.test_name_comboBox.addItem("")
        self.test_name_comboBox.addItem("")
        self.test_name_comboBox.addItem("")
        self.test_name_comboBox.addItem("")
        self.test_name_comboBox.addItem("")
        self.test_name_comboBox.addItem("")
        self.test_name_comboBox.addItem("")
        self.test_name_comboBox.addItem("")
        self.test_name_comboBox.addItem("")
        self.test_name_label = QtWidgets.QLabel(index_frame)
        self.test_name_label.setGeometry(QtCore.QRect(780, 190, 120, 35))
        self.test_name_label.setObjectName("test_name_label")
        self.lineEdit_2 = QtWidgets.QLineEdit(index_frame)
        self.lineEdit_2.setGeometry(QtCore.QRect(180, 190, 100, 35))        # 年份
        self.lineEdit_2.setText("2019")
        self.lineEdit_2.setObjectName("lineEdit_2")
        self.year_label = QtWidgets.QLabel(index_frame)
        self.year_label.setGeometry(QtCore.QRect(120, 190, 75, 35))
        self.year_label.setObjectName("year_label")
        self.line = QtWidgets.QFrame(index_frame)
        self.line.setGeometry(QtCore.QRect(300, 180, 3, 61))
        self.line.setFrameShape(QtWidgets.QFrame.VLine)
        self.line.setFrameShadow(QtWidgets.QFrame.Sunken)
        self.line.setObjectName("line")
        self.line_2 = QtWidgets.QFrame(index_frame)
        self.line_2.setGeometry(QtCore.QRect(570, 180, 3, 61))
        self.line_2.setFrameShape(QtWidgets.QFrame.VLine)
        self.line_2.setFrameShadow(QtWidgets.QFrame.Sunken)
        self.line_2.setObjectName("line_2")
        self.line_3 = QtWidgets.QFrame(index_frame)
        self.line_3.setGeometry(QtCore.QRect(760, 180, 3, 61))
        self.line_3.setFrameShape(QtWidgets.QFrame.VLine)
        self.line_3.setFrameShadow(QtWidgets.QFrame.Sunken)
        self.line_3.setObjectName("line_3")
        self.line_4 = QtWidgets.QFrame(index_frame)
        self.line_4.setGeometry(QtCore.QRect(120, 170, 881, 16))
        self.line_4.setFrameShape(QtWidgets.QFrame.HLine)
        self.line_4.setFrameShadow(QtWidgets.QFrame.Sunken)
        self.line_4.setObjectName("line_4")
        self.line_5 = QtWidgets.QFrame(index_frame)
        self.line_5.setGeometry(QtCore.QRect(120, 240, 881, 16))
        self.line_5.setFrameShape(QtWidgets.QFrame.HLine)
        self.line_5.setFrameShadow(QtWidgets.QFrame.Sunken)
        self.line_5.setObjectName("line_5")
        self.class1_label = QtWidgets.QLabel(index_frame)
        self.class1_label.setGeometry(QtCore.QRect(120, 310, 75, 35))
        self.class1_label.setObjectName("class1_label")
        self.lineEdit_3 = QtWidgets.QLineEdit(index_frame)          # 1班人数
        self.lineEdit_3.setGeometry(QtCore.QRect(170, 310, 61, 35))
        self.lineEdit_3.setText("50")        # 一班
        self.lineEdit_3.setObjectName("lineEdit_3")
        self.bignumber_label = QtWidgets.QLabel(index_frame)
        self.bignumber_label.setGeometry(QtCore.QRect(120, 260, 300, 35))
        font = QtGui.QFont()
        font.setPointSize(14)
        font.setBold(False)
        font.setWeight(50)
        self.bignumber_label.setFont(font)
        self.bignumber_label.setObjectName("bignumber_label")
        self.lineEdit_4 = QtWidgets.QLineEdit(index_frame)
        self.lineEdit_4.setGeometry(QtCore.QRect(170, 360, 61, 35))
        self.lineEdit_4.setText("51")       # 二班
        self.lineEdit_4.setObjectName("lineEdit_4")
        self.class2_label = QtWidgets.QLabel(index_frame)
        self.class2_label.setGeometry(QtCore.QRect(120, 360, 75, 35))
        self.class2_label.setObjectName("class2_label")
        self.lineEdit_5 = QtWidgets.QLineEdit(index_frame)
        self.lineEdit_5.setGeometry(QtCore.QRect(170, 410, 61, 35))
        self.lineEdit_5.setText("49")
        self.lineEdit_5.setObjectName("lineEdit_5")
        self.class3_label = QtWidgets.QLabel(index_frame)
        self.class3_label.setGeometry(QtCore.QRect(120, 410, 75, 35))
        self.class3_label.setObjectName("class3_label")
        self.lineEdit_6 = QtWidgets.QLineEdit(index_frame)
        self.lineEdit_6.setGeometry(QtCore.QRect(170, 460, 61, 35))
        self.lineEdit_6.setText("40")
        self.lineEdit_6.setObjectName("lineEdit_6")
        self.class4_label = QtWidgets.QLabel(index_frame)
        self.class4_label.setGeometry(QtCore.QRect(120, 460, 75, 35))
        self.class4_label.setObjectName("class4_label")
        self.lineEdit_7 = QtWidgets.QLineEdit(index_frame)
        self.lineEdit_7.setGeometry(QtCore.QRect(170, 510, 61, 35))
        self.lineEdit_7.setText("49")
        self.lineEdit_7.setObjectName("lineEdit_7")
        self.class5_label = QtWidgets.QLabel(index_frame)
        self.class5_label.setGeometry(QtCore.QRect(120, 510, 75, 35))
        self.class5_label.setObjectName("class5_label")
        self.class6_label = QtWidgets.QLabel(index_frame)
        self.class6_label.setGeometry(QtCore.QRect(120, 560, 75, 35))
        self.class6_label.setObjectName("class6_label")
        self.lineEdit_8 = QtWidgets.QLineEdit(index_frame)
        self.lineEdit_8.setGeometry(QtCore.QRect(170, 560, 61, 35))
        self.lineEdit_8.setText("51")
        self.lineEdit_8.setObjectName("lineEdit_8")
        self.class7_label = QtWidgets.QLabel(index_frame)
        self.class7_label.setGeometry(QtCore.QRect(270, 310, 75, 35))
        self.class7_label.setObjectName("class7_label")
        self.lineEdit_9 = QtWidgets.QLineEdit(index_frame)
        self.lineEdit_9.setGeometry(QtCore.QRect(320, 310, 61, 35))
        self.lineEdit_9.setText("50")
        self.lineEdit_9.setObjectName("lineEdit_9")
        self.class8_label = QtWidgets.QLabel(index_frame)
        self.class8_label.setGeometry(QtCore.QRect(270, 360, 75, 35))
        self.class8_label.setObjectName("class8_label")
        self.lineEdit_10 = QtWidgets.QLineEdit(index_frame)
        self.lineEdit_10.setGeometry(QtCore.QRect(320, 360, 61, 35))
        self.lineEdit_10.setText("40")
        self.lineEdit_10.setObjectName("lineEdit_10")
        self.lineEdit_11 = QtWidgets.QLineEdit(index_frame)
        self.lineEdit_11.setGeometry(QtCore.QRect(320, 410, 61, 35))
        self.lineEdit_11.setText("")
        self.lineEdit_11.setObjectName("lineEdit_11")
        self.class9_label = QtWidgets.QLabel(index_frame)
        self.class9_label.setGeometry(QtCore.QRect(270, 410, 75, 35))
        self.class9_label.setObjectName("class9_label")
        self.class10_label = QtWidgets.QLabel(index_frame)
        self.class10_label.setGeometry(QtCore.QRect(270, 460, 75, 35))
        self.class10_label.setObjectName("class10_label")
        self.lineEdit_12 = QtWidgets.QLineEdit(index_frame)
        self.lineEdit_12.setGeometry(QtCore.QRect(320, 460, 61, 35))
        self.lineEdit_12.setText("")
        self.lineEdit_12.setObjectName("lineEdit_12")
        self.lineEdit_13 = QtWidgets.QLineEdit(index_frame)
        self.lineEdit_13.setGeometry(QtCore.QRect(320, 510, 61, 35))
        self.lineEdit_13.setText("")
        self.lineEdit_13.setObjectName("lineEdit_13")
        self.class11_label = QtWidgets.QLabel(index_frame)
        self.class11_label.setGeometry(QtCore.QRect(270, 510, 75, 35))
        self.class11_label.setObjectName("class11_label")
        self.class12_label = QtWidgets.QLabel(index_frame)
        self.class12_label.setGeometry(QtCore.QRect(270, 560, 75, 35))
        self.class12_label.setObjectName("class12_label")
        self.lineEdit_14 = QtWidgets.QLineEdit(index_frame)
        self.lineEdit_14.setGeometry(QtCore.QRect(320, 560, 61, 35))
        self.lineEdit_14.setText("")
        self.lineEdit_14.setObjectName("lineEdit_14")
        self.class13_label = QtWidgets.QLabel(index_frame)
        self.class13_label.setGeometry(QtCore.QRect(420, 310, 75, 35))
        self.class13_label.setObjectName("class13_label")
        self.lineEdit_15 = QtWidgets.QLineEdit(index_frame)
        self.lineEdit_15.setGeometry(QtCore.QRect(470, 310, 61, 35))
        self.lineEdit_15.setText("")
        self.lineEdit_15.setObjectName("lineEdit_15")
        self.class14_label = QtWidgets.QLabel(index_frame)
        self.class14_label.setGeometry(QtCore.QRect(420, 360, 75, 35))
        self.class14_label.setObjectName("class14_label")
        self.lineEdit_16 = QtWidgets.QLineEdit(index_frame)
        self.lineEdit_16.setGeometry(QtCore.QRect(470, 360, 61, 35))
        self.lineEdit_16.setText("")
        self.lineEdit_16.setObjectName("lineEdit_16")
        self.class16_label = QtWidgets.QLabel(index_frame)
        self.class16_label.setGeometry(QtCore.QRect(420, 460, 75, 35))
        self.class16_label.setObjectName("class16_label")
        self.lineEdit_17 = QtWidgets.QLineEdit(index_frame)
        self.lineEdit_17.setGeometry(QtCore.QRect(470, 460, 61, 35))
        self.lineEdit_17.setText("")
        self.lineEdit_17.setObjectName("lineEdit_17")
        self.class15_label = QtWidgets.QLabel(index_frame)
        self.class15_label.setGeometry(QtCore.QRect(420, 410, 75, 35))
        self.class15_label.setObjectName("class15_label")
        self.lineEdit_18 = QtWidgets.QLineEdit(index_frame)
        self.lineEdit_18.setGeometry(QtCore.QRect(470, 410, 61, 35))
        self.lineEdit_18.setText("")
        self.lineEdit_18.setObjectName("lineEdit_18")
        self.class17_label = QtWidgets.QLabel(index_frame)
        self.class17_label.setGeometry(QtCore.QRect(420, 510, 75, 35))
        self.class17_label.setObjectName("class17_label")
        self.lineEdit_19 = QtWidgets.QLineEdit(index_frame)
        self.lineEdit_19.setGeometry(QtCore.QRect(470, 510, 61, 35))
        self.lineEdit_19.setText("")      # 第十七个班
        self.lineEdit_19.setObjectName("lineEdit_19")
        self.class18_label = QtWidgets.QLabel(index_frame)
        self.class18_label.setGeometry(QtCore.QRect(420, 560, 75, 35))
        self.class18_label.setObjectName("class18_label")
        self.lineEdit_20 = QtWidgets.QLineEdit(index_frame)
        self.lineEdit_20.setGeometry(QtCore.QRect(470, 560, 61, 35))
        self.lineEdit_20.setText("")      # 第十八个班
        self.lineEdit_20.setObjectName("lineEdit_20")
        self.choose_savexls_Button = QtWidgets.QPushButton(index_frame)
        self.choose_savexls_Button.setGeometry(QtCore.QRect(660, 280, 250, 35))
        self.choose_savexls_Button.setObjectName("choose_savexls_Button")
        self.choose_savexls_Button.clicked.connect(self.save_choose_file)
        self.four_label = QtWidgets.QLabel(index_frame)
        self.four_label.setGeometry(QtCore.QRect(570, 270, 50, 35))
        font = QtGui.QFont()
        font.setPointSize(14)
        self.four_label.setFont(font)
        self.four_label.setObjectName("four_label")
        self.line_6 = QtWidgets.QFrame(index_frame)
        self.line_6.setGeometry(QtCore.QRect(540, 260, 20, 331))
        self.line_6.setFrameShape(QtWidgets.QFrame.VLine)
        self.line_6.setFrameShadow(QtWidgets.QFrame.Sunken)
        self.line_6.setObjectName("line_6")
        self.choose_savexls__address_lineEdit = QtWidgets.QLineEdit(index_frame)
        self.choose_savexls__address_lineEdit.setGeometry(QtCore.QRect(570, 340, 461, 35))
        self.choose_savexls__address_lineEdit.setObjectName("choose_savexls__address_lineEdit")
        self.five_label = QtWidgets.QLabel(index_frame)
        self.five_label.setGeometry(QtCore.QRect(570, 430, 50, 35))
        font = QtGui.QFont()
        font.setPointSize(14)
        self.five_label.setFont(font)
        self.five_label.setObjectName("five_label")
        self.make_Button = QtWidgets.QPushButton(index_frame)
        self.make_Button.setGeometry(QtCore.QRect(660, 450, 250, 35))
        self.make_Button.setObjectName("make_Button")
        self.make_Button.clicked.connect(self.make)
        self.warning_lineEdit = QtWidgets.QLineEdit(index_frame)
        self.warning_lineEdit.setGeometry(QtCore.QRect(700, 510, 311, 35))
        self.warning_lineEdit.setObjectName("warning_lineEdit")
        self.warning_label = QtWidgets.QLabel(index_frame)
        self.warning_label.setGeometry(QtCore.QRect(600, 510, 111, 35))
        self.warning_label.setObjectName("warning_label")
        self.ps_label = QtWidgets.QLabel(index_frame)
        self.ps_label.setGeometry(QtCore.QRect(120, 600, 271, 35))
        self.ps_label.setObjectName("ps_label")
        self.line_7 = QtWidgets.QFrame(index_frame)
        self.line_7.setGeometry(QtCore.QRect(550, 400, 481, 16))
        self.line_7.setFrameShape(QtWidgets.QFrame.HLine)
        self.line_7.setFrameShadow(QtWidgets.QFrame.Sunken)
        self.line_7.setObjectName("line_7")

        self.retranslateUi(index_frame)
        QtCore.QMetaObject.connectSlotsByName(index_frame)


    def retranslateUi(self, index_frame):
        _translate = QtCore.QCoreApplication.translate
        index_frame.setWindowTitle(_translate("index_frame", "真光考场随机座位表生成"))
        self.title_label.setText(_translate("index_frame", "考场随机座位生成功能"))
        self.choose_namexls_Button.setText(_translate("index_frame", "点击选择考生名单表格文件："))
        self.one_label.setText(_translate("index_frame", "1、"))
        self.two_label.setText(_translate("index_frame", "2、"))
        self.three_label.setText(_translate("index_frame", "3、"))
        self.number_class_label.setText(_translate("index_frame", "考场班别数量："))
        self.grade_comboBox.setItemText(0, _translate("index_frame", "初一"))
        self.grade_comboBox.setItemText(1, _translate("index_frame", "初二"))
        self.grade_comboBox.setItemText(2, _translate("index_frame", "初三"))
        self.grade_comboBox.setItemText(3, _translate("index_frame", "高一"))
        self.grade_comboBox.setItemText(4, _translate("index_frame", "高二"))
        self.grade_comboBox.setItemText(5, _translate("index_frame", "高三"))
        self.grade_comboBox.setItemText(6, _translate("index_frame", "其他"))
        self.grade_label.setText(_translate("index_frame", "年级："))
        self.test_name_comboBox.setItemText(0, _translate("index_frame", "入学考"))
        self.test_name_comboBox.setItemText(1, _translate("index_frame", "摸底考"))
        self.test_name_comboBox.setItemText(2, _translate("index_frame", "期中考"))
        self.test_name_comboBox.setItemText(3, _translate("index_frame", "模拟考"))
        self.test_name_comboBox.setItemText(4, _translate("index_frame", "期末考"))
        self.test_name_comboBox.setItemText(5, _translate("index_frame", "1月月考"))
        self.test_name_comboBox.setItemText(6, _translate("index_frame", "2月月考"))
        self.test_name_comboBox.setItemText(7, _translate("index_frame", "3月月考"))
        self.test_name_comboBox.setItemText(8, _translate("index_frame", "4月月考"))
        self.test_name_comboBox.setItemText(9, _translate("index_frame", "5月月考"))
        self.test_name_comboBox.setItemText(10, _translate("index_frame", "6月月考"))
        self.test_name_comboBox.setItemText(11, _translate("index_frame", "7月月考"))
        self.test_name_comboBox.setItemText(12, _translate("index_frame", "8月月考"))
        self.test_name_comboBox.setItemText(13, _translate("index_frame", "9月月考"))
        self.test_name_comboBox.setItemText(14, _translate("index_frame", "10月月考"))
        self.test_name_comboBox.setItemText(15, _translate("index_frame", "11月月考"))
        self.test_name_comboBox.setItemText(16, _translate("index_frame", "12月月考"))
        self.test_name_comboBox.setItemText(17, _translate("index_frame", "小测"))
        self.test_name_comboBox.setItemText(18, _translate("index_frame", "周测"))
        self.test_name_comboBox.setItemText(19, _translate("index_frame", "水平考"))
        self.test_name_comboBox.setItemText(20, _translate("index_frame", "一模考试"))
        self.test_name_comboBox.setItemText(21, _translate("index_frame", "区统考"))
        self.test_name_comboBox.setItemText(22, _translate("index_frame", "二模考试"))
        self.test_name_comboBox.setItemText(23, _translate("index_frame", "三模考试"))
        self.test_name_comboBox.setItemText(24, _translate("index_frame", "四模考试"))
        self.test_name_comboBox.setItemText(25, _translate("index_frame", "文科考试"))
        self.test_name_comboBox.setItemText(26, _translate("index_frame", "理科考试"))
        self.test_name_comboBox.setItemText(27, _translate("index_frame", "其他考试"))
        self.test_name_label.setText(_translate("index_frame", "考试名称："))
        self.year_label.setText(_translate("index_frame", "年份："))
        self.class1_label.setText(_translate("index_frame", "1班："))
        self.bignumber_label.setText(_translate("index_frame", "以下请输入各班容纳最大人数"))
        self.class2_label.setText(_translate("index_frame", "2班："))
        self.class3_label.setText(_translate("index_frame", "3班："))
        self.class4_label.setText(_translate("index_frame", "4班："))
        self.class5_label.setText(_translate("index_frame", "5班："))
        self.class6_label.setText(_translate("index_frame", "6班："))
        self.class7_label.setText(_translate("index_frame", "7班："))
        self.class8_label.setText(_translate("index_frame", "8班："))
        self.class9_label.setText(_translate("index_frame", "9班："))
        self.class10_label.setText(_translate("index_frame", "10班："))
        self.class11_label.setText(_translate("index_frame", "11班："))
        self.class12_label.setText(_translate("index_frame", "12班："))
        self.class13_label.setText(_translate("index_frame", "13班："))
        self.class14_label.setText(_translate("index_frame", "14班："))
        self.class16_label.setText(_translate("index_frame", "16班："))
        self.class15_label.setText(_translate("index_frame", "15班："))
        self.class17_label.setText(_translate("index_frame", "17班："))
        self.class18_label.setText(_translate("index_frame", "18班："))
        self.choose_savexls_Button.setText(_translate("index_frame", "请选择生成座位表保存位置"))
        self.four_label.setText(_translate("index_frame", "4、"))
        self.five_label.setText(_translate("index_frame", "5、"))
        self.make_Button.setText(_translate("index_frame", "点击生成考场座位表"))
        self.warning_label.setText(_translate("index_frame", "提示信息："))
        self.ps_label.setText(_translate("index_frame", "提示信息：只填前多少个班别"))

    def open_choose_file(self):
        # app = wx.App(False)
        # # frame = wx.Frame(None, wx.ID_ANY, "真光成绩管理系统")
        # frame = my_frame(None, '真光成绩管理系统')
        # frame.Show(False)
        # frame.on_open
        filename = QFileDialog.getOpenFileName(None, 'Open file', './')
        address = filename[0]
        print(address)
        self.name_address_lineEdit.setText(address)

    def save_choose_file(self):
        # app = wx.App(False)
        # # frame = wx.Frame(None, wx.ID_ANY, "真光成绩管理系统")
        # frame = my_frame(None, '真光成绩管理系统')
        # frame.Show(False)
        # frame.on_open
        filename = QFileDialog.getExistingDirectory(None, "选择文件夹", "./")
        print(filename)
        address = filename
        print(address)
        self.choose_savexls__address_lineEdit.setText(address)

    def make(self):
        # 记得判断文本是否为空，若为空，则默认失败
        all_dict = {}
        warning = '填写一切正常'
        filename_address = self.name_address_lineEdit.text()
        year = 2019
        year = self.lineEdit_2.text()
        year = int(year)
        number_class = self.lineEdit.text()
        filename_save = self.choose_savexls__address_lineEdit.text()
        test_name = self.test_name_comboBox.currentText()
        grade = self.grade_comboBox.currentText()
        number_class = int(number_class)
        number_inclass1_max = self.lineEdit_3.text()
        number_inclass2_max = self.lineEdit_4.text()
        number_inclass3_max = self.lineEdit_5.text()
        number_inclass4_max = self.lineEdit_6.text()
        number_inclass5_max = self.lineEdit_7.text()
        number_inclass6_max = self.lineEdit_8.text()
        number_inclass7_max = self.lineEdit_9.text()
        number_inclass8_max = self.lineEdit_10.text()
        number_inclass9_max = self.lineEdit_11.text()
        number_inclass10_max = self.lineEdit_12.text()
        number_inclass11_max = self.lineEdit_13.text()
        number_inclass12_max = self.lineEdit_14.text()
        number_inclass13_max = self.lineEdit_15.text()
        number_inclass14_max = self.lineEdit_16.text()
        number_inclass15_max = self.lineEdit_17.text()
        number_inclass16_max = self.lineEdit_18.text()
        number_inclass17_max = self.lineEdit_19.text()
        number_inclass18_max = self.lineEdit_20.text()
        number_classMax_writed = 0          # 已填各班容纳数量的班数量
        number_inclass_dict = {}
        number_inclass_key = []
        if(number_inclass1_max.__eq__('') == False):
            number_classMax_writed = number_classMax_writed + 1
            number_inclass_dict['1'] = number_inclass1_max
            number_inclass_key.append(1)
        if(number_inclass2_max.__eq__('') == False):
            number_classMax_writed = number_classMax_writed + 1
            number_inclass_dict['2'] = number_inclass2_max
            number_inclass_key.append(2)
        if (number_inclass3_max.__eq__('') == False):
            number_classMax_writed = number_classMax_writed + 1
            number_inclass_dict['3'] = number_inclass3_max
            number_inclass_key.append(3)
        if (number_inclass4_max.__eq__('') == False):
            number_classMax_writed = number_classMax_writed + 1
            number_inclass_dict['4'] = number_inclass4_max
            number_inclass_key.append(4)
        if (number_inclass5_max.__eq__('') == False):
            number_classMax_writed = number_classMax_writed + 1
            number_inclass_dict['5'] = number_inclass5_max
            number_inclass_key.append(5)
        if (number_inclass6_max.__eq__('') == False):
            number_classMax_writed = number_classMax_writed + 1
            number_inclass_dict['6'] = number_inclass6_max
            number_inclass_key.append(6)
        if (number_inclass7_max.__eq__('') == False):
            number_classMax_writed = number_classMax_writed + 1
            number_inclass_dict['7'] = number_inclass7_max
            number_inclass_key.append(7)
        if (number_inclass8_max.__eq__('') == False):
            number_classMax_writed = number_classMax_writed + 1
            number_inclass_dict['8'] = number_inclass8_max
            number_inclass_key.append(8)
        if (number_inclass9_max.__eq__('') == False):
            number_classMax_writed = number_classMax_writed + 1
            number_inclass_dict['9'] = number_inclass9_max
            number_inclass_key.append(9)
        if (number_inclass10_max.__eq__('') == False):
            number_classMax_writed = number_classMax_writed + 1
            number_inclass_dict['10'] = number_inclass10_max
            number_inclass_key.append(10)
        if (number_inclass11_max.__eq__('') == False):
            number_classMax_writed = number_classMax_writed + 1
            number_inclass_dict['11'] = number_inclass11_max
            number_inclass_key.append(11)
        if (number_inclass12_max.__eq__('') == False):
            number_classMax_writed = number_classMax_writed + 1
            number_inclass_dict['12'] = number_inclass12_max
            number_inclass_key.append(12)
        if (number_inclass13_max.__eq__('') == False):
            number_classMax_writed = number_classMax_writed + 1
            number_inclass_dict['13'] = number_inclass13_max
            number_inclass_key.append(13)
        if (number_inclass14_max.__eq__('') == False):
            number_classMax_writed = number_classMax_writed + 1
            number_inclass_dict['14'] = number_inclass14_max
            number_inclass_key.append(14)
        if (number_inclass15_max.__eq__('') == False):
            number_classMax_writed = number_classMax_writed + 1
            number_inclass_dict['15'] = number_inclass15_max
            number_inclass_key.append(15)
        if (number_inclass16_max.__eq__('') == False):
            number_classMax_writed = number_classMax_writed + 1
            number_inclass_dict['16'] = number_inclass16_max
            number_inclass_key.append(16)
        if (number_inclass17_max.__eq__('') == False):
            number_classMax_writed = number_classMax_writed + 1
            number_inclass_dict['17'] = number_inclass17_max
            number_inclass_key.append(17)
        if (number_inclass18_max.__eq__('') == False):
            number_classMax_writed = number_classMax_writed + 1
            number_inclass_dict['18'] = number_inclass18_max
            number_inclass_key.append(18)
        print('number_inclass_dict:', end="")
        print(number_inclass_dict)
        for i in number_inclass_dict:
            value = number_inclass_dict[i]
            number_inclass_dict[i] = int(value)


        print(number_inclass_key)
        try:
            print('已填各班容纳最大数量的班级数量为%d' % number_classMax_writed)
            if(number_classMax_writed != number_class):
                warning = '请留意所填写班级容纳最大人数的数量与考场班别数量不符'
                print(warning)
                self.warning_lineEdit.setText(warning)
            if(filename_address.__len__() == 0):
                warning = '请选择考生名字表格的文件'
            elif(number_class.__eq__('') == False):
                warning = '请注意填写班别数量'
            elif(filename_save.__len__() == 0):
                warning = '请注意选择保存路径'
            self.warning_lineEdit.setText(warning)
        except Exception as error:
            print('error' + str(error))
        self.warning_lineEdit.setText(warning)
        all_dict['filename_address'] = filename_address
        all_dict['year'] = year
        all_dict['number_class'] = number_class
        all_dict['filename_save'] = filename_save
        all_dict['test_name'] = test_name
        all_dict['grade'] = grade
        all_dict['number_inclass_dict'] = number_inclass_dict
        all_dict['number_inclass_key'] = number_inclass_key
        print(all_dict)
        if warning == '填写一切正常':
            warning = main(all_dict)
            self.warning_lineEdit.setText(warning)


def get_names(address):
    """
    该函数的作用是根据地址取出表格，返回数据表格及打乱后考号列表
    未完成：判断是否考生名单文件
    :param address:
    :return:
    """
    time1 = time.time()
    name_xls = pd.read_excel(address)   # 根据传入地址进行读取文件
    # print(type(name_xls))
    # print(name_xls)
    # print(name_xls.shape)
    dataset = np.array(name_xls)        # 转换成ndarray
    # print(dataset)
    # print(type(dataset))
    primary_key = dataset[:, 1]         # 把考号这列取出
    primary_key_list = list(primary_key)
    # print(primary_key_list)
    # print(type(primary_key_list))
    random.shuffle(primary_key_list)    # 进行随机打乱考号
    primary_key_random_list = primary_key_list
    print('Reshuffled list : ')
    print(primary_key_random_list)
    time2 = time.time()
    print('def : get_names time == ' + str(time2 - time1))
    return dataset, primary_key_random_list


def excel_done(dataset, primary_key_random_list, number_class, number_inclass_max):
    """
    该函数的作用是根据数据表格+打乱考号列表+个性化需求，返回多个班考生名单的大列表
    :param dataset: 考生名单的数据表格，ndarray结构
    :param primary_key_random_list: 主键，打乱后的随机考生号列表
    :param number_class:班别数量
    :param number_inclass_max: 班级坐得下的最大人数
    :return: 多个班考生名单的大列表，小列表为一个班的考生名单
    """
    time1 = time.time()
    number_class = 8            # 调试所用，待删
    number_inclass_max = 50      #
    len_primary_key = len(primary_key_random_list)      # 考生名单数量
    big_name_list = []
    for index_class in range(number_class):
        name_list = []
        print('第%s份打乱名单已出' % str(index_class + 1))
        for i in range(number_inclass_max):
            if((index_class * number_inclass_max + i) >= len_primary_key):  # 遍历到最后一个考生
                break
            kaohao = primary_key_random_list[i + index_class * number_inclass_max]  # 取出考号
            name = ''
            for j in range(dataset.shape[0]):
                if dataset[j][1] != kaohao:     # 没找到考号的学生则跳过
                    continue
                else:
                    name = dataset[j][4]        # 找到则取出该学生的姓名
            name_list.append(name)              # 这个考场的名单加上该名字
        print('name_list:')
        big_name_list.append(name_list)         # 全级的考场名单加上该考场的名单
        print(name_list)
    print('已完成所有名单的表格')
    time2 = time.time()
    # print('big_name_list:')
    # print(big_name_list)
    print('def : excel_done time == ' + str(time2 - time1))
    return big_name_list


def big_name_list_make(dataset, primary_key_random_list, number_class, number_inclass_key, number_inclass_dict):
    """
    该函数的作用是根据数据表格+打乱考号列表+班别数量+每个班容纳人数，返回多个班考生名单的大列表
    :param dataset: 数据表格
    :param primary_key_random_list:打乱考号列表
    :param number_class: 班别数量
    :param number_inclass_max_list:每个班容纳人数的列表
    :return:
    """
    error = 0
    error_information = '程序执行结束，过程正常'
    time1 = time.time()
    number_inclass_max_list = []
    number_inclass_key.sort()       # 有哪些班，如1，2，4，5
    for index in number_inclass_key:
        value = number_inclass_dict[str(index)]
        number_inclass_max_list.append(value)
    len_primary_key = len(primary_key_random_list)  # 考生名单数量
    total_number_max = 0    # 所有班级加起来容纳的最大人数
    for i in number_inclass_max_list:
        total_number_max = total_number_max + int(i)
    if len(primary_key_random_list) > total_number_max:
        error = 1
        error_information = '所有班级最大容纳人数小于考生人数，请检查'
    big_name_list = []
    yipai = 0
    for index_class in range(number_class):     # 一个班一个班这样弄表格
        name_list = []
        print('第%s班打乱名单已出' % str(number_inclass_key[index_class]))
        for i in range(number_inclass_max_list[index_class]):
            if(yipai >= len_primary_key):  # 遍历到最后一个考生
                break
            kaohao = primary_key_random_list[yipai]  # 取出考号
            name = ''
            for j in range(dataset.shape[0]):
                if dataset[j][1] != kaohao:     # 没找到考号的学生则跳过
                    continue
                else:
                    name = dataset[j][4]        # 找到则取出该学生的姓名
                    break
            name_list.append(name)  # 这个考场的名单加上该名字
            yipai = yipai + 1
        print('name_list:')
        print(name_list)
        big_name_list.append(name_list)         # 全级的考场名单加上该考场的名单
    time2 = time.time()
    print('big_name_list已完成各班的所有名单时间：' + str(time2 - time1))
    return big_name_list, number_inclass_key, number_inclass_max_list, error_information


def write_excel(big_name_list, number_inclass_key, number_inclass_max_list, test_name, year, grade, dataset, filename):
    """
    该函数是根据大的名字列表，写入Excel文件中
    :param big_name_list:
    """

    """格式信息"""
    style = xlwt.XFStyle()
    style2 = xlwt.XFStyle()
    style3 = xlwt.XFStyle()
    style4 = xlwt.XFStyle()
    font = xlwt.Font()  # 字体基本设置
    font.name = u'宋体'
    font.color = 'black'
    font.height = 200  # 字体大小，220就是11号字体，大概就是11*20得来的吧
    style.font = font

    font2 = xlwt.Font()  # 字体基本设置
    font2.name = u'宋体'
    font2.color = 'black'
    font2.height = 240  # 字体大小，220就是11号字体，大概就是11*20得来的吧
    style2.font = font2

    font3 = xlwt.Font()  # 字体基本设置
    font3.name = u'宋体'
    font3.color = 'black'
    font3.bold
    font3.height = 240  # 字体大小，220就是11号字体，大概就是11*20得来的吧
    style3.font = font3

    font4 = xlwt.Font()  # 字体基本设置
    font4.name = u'宋体'
    font4.color = 'black'
    font4.bold
    font4.height = 240  # 字体大小，220就是11号字体，大概就是11*20得来的吧
    style4.font = font4

    alignment = xlwt.Alignment()  # 设置字体在单元格的位置
    alignment.horz = xlwt.Alignment.HORZ_CENTER  # 水平方向
    alignment.vert = xlwt.Alignment.VERT_CENTER  # 竖直方向
    style.alignment = alignment
    style2.alignment = alignment
    style3.alignment = alignment
    style4.alignment = alignment

    border = xlwt.Borders()  # 给单元格加框线
    border.left = xlwt.Borders.THIN  # 左
    border.top = xlwt.Borders.THIN  # 上
    border.right = xlwt.Borders.THIN  # 右
    border.bottom = xlwt.Borders.THIN  # 下
    border.left_colour = 0x40  # 设置框线颜色，0x40是黑色，颜色真的巨多，都晕了
    border.right_colour = 0x40
    border.top_colour = 0x40
    border.bottom_colour = 0x40
    style.borders = border
    style2.borders = border
    style3.borders = border
    # 要做到用户可以选择考试名字及年级
    rb = open_workbook(r'D:\StudentSystem\test.xls')

    # 通过sheet_by_index()获取的sheet没有write()方法
    rs = rb.sheet_by_index(0)

    wb = copy(rb)
    # wb = Workbook()

    dengfen_name_list = []
    dengfen_dict = {}
    for i in range(len(number_inclass_key)):
        banbie = number_inclass_key[i]
        number_inclass_maxx = number_inclass_max_list[i]
        # 通过get_sheet()获取的sheet有write()方法
        ws = wb.get_sheet(i)
        # ws = wb.worksheets[0]
        columns = 6
        return_name_list, return_str_list = class_names(big_name_list, i, number_inclass_maxx, grade, banbie)
        print('return_str_list', end="")
        print(return_str_list)
        print('return_name_list', end="")
        print(return_name_list)
        ren = 0
        for i in range(columns):
            for j in range(len(return_name_list[i])):
                ren = ren + 1
                ws.write(10 - j, i * 3, int(return_str_list[i][j]), style=style2)
                ws.write(10 - j, i * 3 + 1, return_name_list[i][j], style=style)
                dengfen_dict[return_name_list[i][j]] = int(return_str_list[i][j])
                dengfen_name_list.append(return_name_list[i][j])
        # if(len(return_name_list) > 6):
        #     for i in range(len(return_name_list[6])):
        #         ws.write(10 - j, i * 3, int(return_str_list[i][j]))
        chuji_utf8 = grade
        ws.write(0, 7, str(year) + chuji_utf8 + test_name + str(banbie) + '班座位表' + str(ren) + '人', style=style4)
        ws.write(12, 0, '前门', style=style3)
        ws.write(1, 0, '后门', style=style3)
        ws.write(12, 8, '讲台', style=style3)

    # 这里是登分表
    ws = wb.get_sheet(len(number_inclass_key))
    ws.write(0, 0, '班别', style=style3)
    ws.write(0, 1, '学号', style=style3)
    ws.write(0, 2, '考号', style=style3)
    ws.write(0, 3, '座位号', style=style3)
    ws.write(0, 4, '姓名', style=style3)
    for i in range(len(dengfen_name_list)):
        ws.write(i + 1, 4, dengfen_name_list[i], style=style)
        ws.write(i + 1, 3, int(dengfen_dict[dengfen_name_list[i]]), style=style)
        for j in range(dataset.shape[0]):
            if str(dataset[j][4]).__eq__(dengfen_name_list[i]):
                ws.write(i + 1, 2, int(dataset[j][1]), style=style)
                ws.write(i + 1, 1, str(dataset[j][3]), style=style)
                ws.write(i + 1, 0, int(dataset[j][2]), style=style)

    String = str(year) + chuji_utf8 + test_name
    wb.save(r'%s\%s考试座位表.xls' % (filename, String) )


    # 独立的登分表
    wb = copy(rb)
    ws = wb.get_sheet(0)
    ws.write(0, 0, '班别', style=style3)
    ws.write(0, 1, '学号', style=style3)
    ws.write(0, 2, '考号', style=style3)
    ws.write(0, 3, '座位号', style=style3)
    ws.write(0, 4, '姓名', style=style3)
    for i in range(len(dengfen_name_list)):
        ws.write(i + 1, 4, dengfen_name_list[i], style=style)
        ws.write(i + 1, 3, int(dengfen_dict[dengfen_name_list[i]]), style=style)
        for j in range(dataset.shape[0]):
            if str(dataset[j][4]).__eq__(dengfen_name_list[i]):
                ws.write(i + 1, 2, int(dataset[j][1]), style=style)
                ws.write(i + 1, 1, str(dataset[j][3]), style=style)
                ws.write(i + 1, 0, int(dataset[j][2]), style=style)
    wb.save(r'%s\%s考试登分表.xls' % (filename, String))


    #以下是完整表
    wb = copy(rb)
    # wb = Workbook()
    dengfen_name_list = []
    dengfen_dict = {}
    ws = wb.get_sheet(0)
    for i in range(len(number_inclass_key)):
        index = i
        banbie = number_inclass_key[i]
        number_inclass_maxx = number_inclass_max_list[i]
        # 通过get_sheet()获取的sheet有write()方法
        # ws = wb.worksheets[0]
        columns = 6
        return_name_list, return_str_list = class_names(big_name_list, i, number_inclass_maxx, grade, banbie)
        print('return_str_list', end="")
        print(return_str_list)
        print('return_name_list', end="")
        print(return_name_list)
        ren = 0
        for i in range(columns):
            for j in range(len(return_name_list[i])):
                ren = ren + 1
                ws.write(10 - j + index * 14, i * 3, int(return_str_list[i][j]), style=style2)
                ws.write(10 - j + index * 14, i * 3 + 1, return_name_list[i][j], style=style)
                dengfen_dict[return_name_list[i][j]] = int(return_str_list[i][j])
                dengfen_name_list.append(return_name_list[i][j])
        # if(len(return_name_list) > 6):
        #     for i in range(len(return_name_list[6])):
        #         ws.write(10 - j, i * 3, int(return_str_list[i][j]))
        chuji_utf8 = grade
        ws.write(0 + index * 14, 7, str(year) + chuji_utf8 + test_name + str(banbie) + '班座位表' + str(ren) + '人', style=style4)
        ws.write(12 + index * 14, 0, '前门', style=style3)
        ws.write(1 + index * 14, 0, '后门', style=style3)
        ws.write(12 + index * 14, 8, '讲台', style=style3)

    # 完整表第二页的登分表
    ws = wb.get_sheet(1)
    ws.write(0, 0, '班别', style=style3)
    ws.write(0, 1, '学号', style=style3)
    ws.write(0, 2, '考号', style=style3)
    ws.write(0, 3, '座位号', style=style3)
    ws.write(0, 4, '姓名', style=style3)
    for i in range(len(dengfen_name_list)):
        ws.write(i + 1, 4, dengfen_name_list[i], style=style)
        ws.write(i + 1, 3, int(dengfen_dict[dengfen_name_list[i]]), style=style)
        for j in range(dataset.shape[0]):
            if str(dataset[j][4]).__eq__(dengfen_name_list[i]):
                ws.write(i + 1, 2, int(dataset[j][1]), style=style)
                ws.write(i + 1, 1, str(dataset[j][3]), style=style)
                ws.write(i + 1, 0, int(dataset[j][2]), style=style)
    wb.save(r'%s\%s考试完整表.xls' % (filename, String))
    print('excel写入完毕！')


def class_names(big_name_list, index, number_inclass_max, grade, banbie):
    if grade == '初一':
        chuji = 1
    elif grade == '初二':
        chuji = 2
    elif grade == '初三':
        chuji = 3
    elif grade == '高一':
        chuji = 4
    elif grade == '高二':
        chuji = 5
    elif grade == '高三':
        chuji = 6
    else:
        chuji = 7
    columns = 6
    name_list = big_name_list[index]
    print('name_list:')
    print(name_list)
    return_name_list = []
    return_str_list = []
    if len(name_list) == number_inclass_max:    # 整个班达到了最大容纳量
        j = 0                                   # 序数，方便看第几个学生
        rows = int(len(name_list) / columns)    # 计算能填满多少行
        number_zhengchang = rows * columns      # 行和列是多少，以便看还剩下多少人
        number_others = number_inclass_max - number_zhengchang  # 最后一行补充学生数量
        index_list = []                         # 名字列表
        str_list = []                           # 编号列表
        print('class_names——number_zhengchang:', end="")
        print(number_zhengchang)
        print('number_others:', end="")
        print(number_others)
        flag_rows = 0                           # 标志，是否要在某一列学生加一个位后再减一进入下一列
        for i in range(len(name_list)):
            # if(flag_rows == 1):
            #     rows = rows - 1                     # 进入下一行用正常的行数
            #     number_others = number_others - 1   # 剪掉一个多余的学生
            #     flag_rows = 0                       # 重新初始化
            # if(number_others > 0 and i > 0 and flag_rows == 0 and len(return_name_list) > 0):     # 有多余学生且不是第一列
            #     flag_rows = 1
            #     rows = rows + 1
            print('rows:', end="")
            print(rows)
            j = j + 1
            index_list.append(name_list[i])
            if j <= 9:
                xushu = '0' + str(j)
            else:
                xushu = str(j)
            string = str(chuji) + str(banbie) + str(xushu)
            str_list.append(string)
            if(j > number_zhengchang):
                continue
            if((len(index_list)) == rows ):
            # if((i + 1) % rows == 0):
                # 一列满了
                return_name_list.append(index_list)
                index_list = []
                return_str_list.append(str_list)
                str_list = []
                if(len(return_name_list) == 1 and number_others > 0):
                    rows = rows + 1
                    flag_rows = 1
                elif((number_others > 0) and (len(return_name_list) > 1) and flag_rows == 1):
                    number_others = number_others - 1
                if(flag_rows == 1 and number_others == 0):
                    flag_rows = 0
                    rows = rows - 1
        if(len(index_list) != 0):
            return_name_list.append(index_list)     # 剩下的人数
            return_str_list.append(str_list)
    else:
        j = 0  # 序数，方便看第几个学生
        rows = int(len(name_list) / columns)  # 计算能填满多少行
        number_zhengchang = rows * columns  # 行和列是多少，以便看还剩下多少人
        number_others = len(name_list) - number_zhengchang  # 最后一行补充学生数量
        index_list = []  # 名字列表
        str_list = []  # 编号列表
        print('class_names——number_zhengchang:', end="")
        print(number_zhengchang)
        print('number_others:', end="")
        print(number_others)
        flag_rows = 0  # 标志，是否要在某一列学生加一个位后再减一进入下一列
        for i in range(len(name_list)):
            # if(flag_rows == 1):
            #     rows = rows - 1                     # 进入下一行用正常的行数
            #     number_others = number_others - 1   # 剪掉一个多余的学生
            #     flag_rows = 0                       # 重新初始化
            # if(number_others > 0 and i > 0 and flag_rows == 0 and len(return_name_list) > 0):     # 有多余学生且不是第一列
            #     flag_rows = 1
            #     rows = rows + 1
            print('rows:', end="")
            print(rows)
            j = j + 1
            index_list.append(name_list[i])
            if j <= 9:
                xushu = '0' + str(j)
            else:
                xushu = str(j)
            string = str(chuji) + str(banbie) + str(xushu)
            str_list.append(string)
            if (j > number_zhengchang):
                continue
            if ((len(index_list)) == rows):
                # if((i + 1) % rows == 0):
                # 一列满了
                return_name_list.append(index_list)
                index_list = []
                return_str_list.append(str_list)
                str_list = []
                if (len(return_name_list) == 1 and number_others > 0):
                    rows = rows + 1
                    flag_rows = 1
                elif ((number_others > 0) and (len(return_name_list) > 1) and flag_rows == 1):
                    number_others = number_others - 1
                if (flag_rows == 1 and number_others == 0):
                    flag_rows = 0
                    rows = rows - 1
        if (len(index_list) != 0):
            return_name_list.append(index_list)  # 剩下的人数
            return_str_list.append(str_list)
    return return_name_list, return_str_list


# 设置主函数，便于图形界面直接调用完整功能
def main(adict):
    time1 = time.time()
    try:
        dataset, primary_key_random_list = get_names(adict['filename_address'])
    except Exception as e:
        print('get_names报错信息')
        error_information = e
        print(e)
    try:
        big_name_list, number_inclass_key, number_inclass_max_list, error_information = big_name_list_make(dataset, primary_key_random_list, adict['number_class'], adict['number_inclass_key'], adict['number_inclass_dict'])
    except Exception as e:
        print('big_name_list_make报错信息')
        error_information = e
        print(e)
    try:
        write_excel(big_name_list, number_inclass_key, number_inclass_max_list, adict['test_name'], adict['year'], adict['grade'], dataset, adict['filename_save'])
    except Exception as e:
        print('write_excel报错信息')
        error_information = e
        print(e)
    time2 = time.time()
    print('完整所需时间:' + str(time2 - time1))
    return error_information


if __name__=="__main__":
    app=QtWidgets.QApplication(sys.argv)
    widget=QtWidgets.QWidget()
    ui =Ui_index_frame()
    ui.setupUi(widget)
    widget.show()
    sys.exit(app.exec_())