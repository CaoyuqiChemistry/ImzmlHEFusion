from PyQt5 import QtCore, QtGui, QtWidgets
from PyQt5.QtWidgets import QComboBox,QLineEdit,QListWidget,QCheckBox,QListWidgetItem
from pyimzml.ImzMLParser import ImzMLParser
from matplotlib import cm
from progressbar import *
import numpy as np
import os
import matplotlib.colors as col
from scipy.interpolate import griddata
from matplotlib import pyplot as plt
import xlwt,xlrd
from matplotlib.widgets import Button,TextBox,Slider,RadioButtons
from matplotlib.patches import Polygon
from sklearn.cross_decomposition import PLSRegression
from sklearn.tree import DecisionTreeRegressor
from sklearn.neural_network import MLPRegressor
from sklearn.preprocessing import StandardScaler
from sklearn.pipeline import make_pipeline
import PIL
import matplotlib
matplotlib.use('Qt5Agg')
import warnings
import me_rc
warnings.filterwarnings('ignore')

class ComboCheckBox(QComboBox):
    def __init__(self, items):  # items==[str,str...]
        super(ComboCheckBox, self).__init__()
        self.items = items
        self.items.insert(0, 'Total')
        self.clear()
        self.addItems(self.items)
        self.row_num = len(self.items)
        self.Selectedrow_num = 0
        self.qCheckBox = []
        self.qLineEdit = QLineEdit()
        self.qLineEdit.setReadOnly(True)
        self.qListWidget = QListWidget()
        self.addQCheckBox(0)
        self.qCheckBox[0].stateChanged.connect(self.All)
        for i in range(1, self.row_num):
            self.addQCheckBox(i)
            self.qCheckBox[i].stateChanged.connect(self.show)
        self.setModel(self.qListWidget.model())
        self.setView(self.qListWidget)
        self.setLineEdit(self.qLineEdit)

    def Add_Items(self,p):
        self.items = p
        self.items.insert(0, 'Total')
        self.row_num = len(self.items)
        self.Selectedrow_num = 0
        self.qCheckBox = []
        self.qLineEdit = QLineEdit()
        self.qLineEdit.setReadOnly(True)
        self.qListWidget = QListWidget()
        self.addQCheckBox(0)
        self.qCheckBox[0].stateChanged.connect(self.All)
        for i in range(1, self.row_num):
            self.addQCheckBox(i)
            self.qCheckBox[i].stateChanged.connect(self.Checkbox_Show)
        self.setModel(self.qListWidget.model())
        self.setView(self.qListWidget)
        self.setLineEdit(self.qLineEdit)

    def leaveEvent(self,event):
        m = self.SelectIndex()
        txt = self.qLineEdit.text()
        if self.qLineEdit.text() != '':
            self.row_num = len(self.items)
            self.qCheckBox = []
            self.qLineEdit = QLineEdit()
            self.qLineEdit.setReadOnly(True)
            self.qListWidget = QListWidget()
            self.addQCheckBox(0)
            self.qCheckBox[0].stateChanged.connect(self.All)
            for i in range(1, self.row_num):
                self.addQCheckBox(i)
                self.qCheckBox[i].stateChanged.connect(self.Checkbox_Show)
            for j in m:
                self.qCheckBox[j+1].setCheckState(2)
            self.setModel(self.qListWidget.model())
            self.setView(self.qListWidget)
            self.setLineEdit(self.qLineEdit)
            self.qLineEdit.setReadOnly(False)
            self.qLineEdit.setText(txt)
            self.qLineEdit.setReadOnly(True)

    def Clear_Items(self):
        self.items = []
        self.row_num = len(self.items)
        self.Selectedrow_num = 0
        self.qCheckBox = []

    def addQCheckBox(self, i):
        self.qCheckBox.append(QCheckBox())
        qItem = QListWidgetItem(self.qListWidget)
        self.qCheckBox[i].setText(self.items[i])
        self.qListWidget.setItemWidget(qItem, self.qCheckBox[i])

    def Selectlist(self):
        Outputlist = []
        for i in range(1, self.row_num):
            if self.qCheckBox[i].isChecked() == True:
                Outputlist.append(self.qCheckBox[i].text())
        self.Selectedrow_num = len(Outputlist)
        return Outputlist

    def SelectIndex(self):
        Outputlist = []
        for i in range(1, self.row_num):
            if self.qCheckBox[i].isChecked() == True:
                Outputlist.append(i-1)
        self.Selectedrow_num = len(Outputlist)
        return Outputlist

    def Checkbox_Show(self):
        show = ''
        Outputlist = self.Selectlist()
        self.qLineEdit.setReadOnly(False)
        self.qLineEdit.clear()
        for i in Outputlist:
            show += i + ';'
        if self.Selectedrow_num == 0:
            self.qCheckBox[0].setCheckState(0)
        elif self.Selectedrow_num == self.row_num - 1:
            self.qCheckBox[0].setCheckState(2)
        else:
            self.qCheckBox[0].setCheckState(1)
        self.qLineEdit.setText(show)
        self.qLineEdit.setReadOnly(True)

    def All(self, zhuangtai):
        if zhuangtai == 2:
            for i in range(1, self.row_num):
                self.qCheckBox[i].setChecked(True)
        elif zhuangtai == 1:
            if self.Selectedrow_num == 0:
                self.qCheckBox[0].setCheckState(2)
        elif zhuangtai == 0:
            self.Checkbox_Clear()

    def Checkbox_Clear(self):
        for i in range(self.row_num):
            self.qCheckBox[i].setChecked(False)

class Progress_Form(object):
    def setupUi(self, Form):
        Form.setObjectName("Processing...")
        Form.resize(497, 147)
        icon = QtGui.QIcon()
        icon.addPixmap(QtGui.QPixmap(":/newPrefix/ui/icon.png"), QtGui.QIcon.Normal, QtGui.QIcon.Off)
        Form.setWindowIcon(icon)
        self.setWindowModality(QtCore.Qt.ApplicationModal)
        self.setWindowFlags(QtCore.Qt.FramelessWindowHint)
        self.horizontalLayout_2 = QtWidgets.QHBoxLayout(Form)
        self.horizontalLayout_2.setObjectName("horizontalLayout_2")
        self.verticalLayout = QtWidgets.QVBoxLayout()
        self.verticalLayout.setContentsMargins(5, 5, 5, 5)
        self.verticalLayout.setObjectName("verticalLayout")
        self.progressBar = QtWidgets.QProgressBar(Form)
        sizePolicy = QtWidgets.QSizePolicy(QtWidgets.QSizePolicy.Expanding, QtWidgets.QSizePolicy.Expanding)
        sizePolicy.setHorizontalStretch(0)
        sizePolicy.setVerticalStretch(1)
        sizePolicy.setHeightForWidth(self.progressBar.sizePolicy().hasHeightForWidth())
        self.progressBar.setSizePolicy(sizePolicy)
        font = QtGui.QFont()
        font.setFamily("Arial")
        font.setBold(True)
        font.setPointSize(12)
        self.progressBar.setFont(font)
        self.progressBar.setProperty("value", 0)
        self.progressBar.setOrientation(QtCore.Qt.Horizontal)
        self.progressBar.setObjectName("progressBar")
        self.verticalLayout.addWidget(self.progressBar)
        spacerItem = QtWidgets.QSpacerItem(20, 10, QtWidgets.QSizePolicy.Minimum, QtWidgets.QSizePolicy.Expanding)
        self.verticalLayout.addItem(spacerItem)
        self.horizontalLayout = QtWidgets.QHBoxLayout()
        self.horizontalLayout.setContentsMargins(10, 10, 10, 10)
        self.horizontalLayout.setSpacing(10)
        self.horizontalLayout.setObjectName("horizontalLayout")
        self.label = QtWidgets.QLabel(Form)
        sizePolicy = QtWidgets.QSizePolicy(QtWidgets.QSizePolicy.Preferred, QtWidgets.QSizePolicy.Preferred)
        sizePolicy.setHorizontalStretch(1)
        sizePolicy.setVerticalStretch(0)
        sizePolicy.setHeightForWidth(self.label.sizePolicy().hasHeightForWidth())
        self.label.setSizePolicy(sizePolicy)
        font = QtGui.QFont()
        font.setFamily("Arial")
        font.setBold(True)
        font.setPointSize(12)
        self.label.setFont(font)
        self.label.setObjectName("label")
        self.horizontalLayout.addWidget(self.label)
        spacerItem1 = QtWidgets.QSpacerItem(60, 20, QtWidgets.QSizePolicy.Expanding, QtWidgets.QSizePolicy.Minimum)
        self.horizontalLayout.addItem(spacerItem1)
        self.pushButton = QtWidgets.QPushButton(Form)
        sizePolicy = QtWidgets.QSizePolicy(QtWidgets.QSizePolicy.Minimum, QtWidgets.QSizePolicy.Fixed)
        sizePolicy.setHorizontalStretch(1)
        sizePolicy.setVerticalStretch(0)
        sizePolicy.setHeightForWidth(self.pushButton.sizePolicy().hasHeightForWidth())
        self.pushButton.setSizePolicy(sizePolicy)
        font = QtGui.QFont()
        font.setFamily("Arial")
        font.setBold(True)
        font.setPointSize(12)
        self.pushButton.setFont(font)
        self.pushButton.setObjectName("pushButton")
        self.horizontalLayout.addWidget(self.pushButton)
        self.verticalLayout.addLayout(self.horizontalLayout)
        self.verticalLayout.setStretch(0, 2)
        self.verticalLayout.setStretch(1, 2)
        self.verticalLayout.setStretch(2, 3)
        self.horizontalLayout_2.addLayout(self.verticalLayout)

        self.retranslateUi(Form)
        self.pushButton.clicked.connect(self.close)
        QtCore.QMetaObject.connectSlotsByName(Form)

    def retranslateUi(self, Form):
        _translate = QtCore.QCoreApplication.translate
        Form.setWindowTitle(_translate("Form", "Processing..."))
        self.label.setText(_translate("Form", "Processing..."))
        self.pushButton.setText(_translate("Form", "Confirm"))
class My_Progress_Form(QtWidgets.QWidget,Progress_Form):
    def __init__(self):
        super().__init__()
        self.setupUi(self)
        col = QtGui.QColor(255,194,189)
        self.setStyleSheet('QWidget{background-color:%s}' % col.name())
class Message_Form(object):
    def setupUi(self, Form):
        Form.setObjectName("Form")
        Form.resize(250, 100)
        icon = QtGui.QIcon()
        icon.addPixmap(QtGui.QPixmap(":/newPrefix/ui/icon.png"), QtGui.QIcon.Normal, QtGui.QIcon.Off)
        Form.setWindowIcon(icon)
        self.horizontalLayout = QtWidgets.QHBoxLayout(Form)
        self.horizontalLayout.setObjectName("horizontalLayout")
        self.label = QtWidgets.QLabel(Form)
        font = QtGui.QFont()
        font.setFamily("Arial")
        font.setBold(True)
        font.setPointSize(14)
        self.label.setFont(font)
        self.label.setAlignment(QtCore.Qt.AlignCenter)
        self.label.setObjectName("label")
        self.horizontalLayout.addWidget(self.label)

        self.retranslateUi(Form)
        QtCore.QMetaObject.connectSlotsByName(Form)

    def retranslateUi(self, Form):
        _translate = QtCore.QCoreApplication.translate
        Form.setWindowTitle(_translate("Form", "Form"))
        self.label.setText(_translate("Form", "Successfully!"))
class My_Message_Form(QtWidgets.QWidget,Message_Form):
    def __init__(self):
        super().__init__()
        self.setupUi(self)
        col = QtGui.QColor(255, 194, 189)
        self.setStyleSheet('QWidget{background-color:%s}' % col.name())
class Error_Form(object):
    def setupUi(self, Form):
        Form.setObjectName("Form")
        Form.resize(400, 200)
        icon = QtGui.QIcon()
        icon.addPixmap(QtGui.QPixmap(":/newPrefix/ui/warning.png"), QtGui.QIcon.Normal, QtGui.QIcon.Off)
        Form.setWindowIcon(icon)
        self.setWindowModality(QtCore.Qt.ApplicationModal)
        self.label = QtWidgets.QLabel(Form)
        self.label.setGeometry(QtCore.QRect(50, 30, 341, 51))
        self.label.setWordWrap(True)
        font = QtGui.QFont()
        font.setFamily("Arial")
        font.setBold(True)
        font.setPointSize(10)
        self.label.setFont(font)
        self.label.setObjectName("label")

        self.retranslateUi(Form)
        QtCore.QMetaObject.connectSlotsByName(Form)

    def retranslateUi(self, Form):
        _translate = QtCore.QCoreApplication.translate
        Form.setWindowTitle(_translate("Form", "运行错误"))
        self.label.setText(_translate("Form", "您的输入有误，请重新输入！"))
class My_Error_Form(QtWidgets.QWidget,Error_Form):
    def __init__(self):
        super().__init__()
        self.setupUi(self)
        col = QtGui.QColor(255, 194, 189)
        self.setStyleSheet('QWidget{background-color:%s}' % col.name())
        self.setMinimumSize(400, 200)
        self.setMaximumSize(400, 200)
class Combo_choose_Form(object):
    def setupUi(self, Form):
        Form.setObjectName("Form")
        Form.resize(490, 286)
        sizePolicy = QtWidgets.QSizePolicy(QtWidgets.QSizePolicy.Preferred, QtWidgets.QSizePolicy.Preferred)
        sizePolicy.setHorizontalStretch(0)
        sizePolicy.setVerticalStretch(0)
        sizePolicy.setHeightForWidth(Form.sizePolicy().hasHeightForWidth())
        self.setWindowModality(QtCore.Qt.ApplicationModal)
        Form.setSizePolicy(sizePolicy)
        icon = QtGui.QIcon()
        icon.addPixmap(QtGui.QPixmap(":/newPrefix/ui/icon.png"), QtGui.QIcon.Normal, QtGui.QIcon.Off)
        Form.setWindowIcon(icon)
        self.horizontalLayout_3 = QtWidgets.QHBoxLayout(Form)
        self.horizontalLayout_3.setObjectName("horizontalLayout_3")
        self.verticalLayout_2 = QtWidgets.QVBoxLayout()
        self.verticalLayout_2.setContentsMargins(10, 20, 10, 10)
        self.verticalLayout_2.setObjectName("verticalLayout_2")
        self.horizontalLayout_2 = QtWidgets.QHBoxLayout()
        self.horizontalLayout_2.setObjectName("horizontalLayout_2")
        self.label_2 = QtWidgets.QLabel(Form)
        sizePolicy = QtWidgets.QSizePolicy(QtWidgets.QSizePolicy.Maximum, QtWidgets.QSizePolicy.Preferred)
        sizePolicy.setHorizontalStretch(0)
        sizePolicy.setVerticalStretch(0)
        sizePolicy.setHeightForWidth(self.label_2.sizePolicy().hasHeightForWidth())
        self.label_2.setSizePolicy(sizePolicy)
        font = QtGui.QFont()
        font.setFamily("Arial")
        font.setPointSize(12)
        font.setBold(True)
        font.setWeight(75)
        self.label_2.setFont(font)
        self.label_2.setObjectName("label_2")
        self.horizontalLayout_2.addWidget(self.label_2)
        self.lineEdit = QtWidgets.QLineEdit(Form)
        font = QtGui.QFont()
        font.setFamily("Arial")
        font.setPointSize(12)
        font.setBold(True)
        font.setWeight(75)
        self.lineEdit.setFont(font)
        self.lineEdit.setObjectName("lineEdit")
        self.horizontalLayout_2.addWidget(self.lineEdit)
        self.pushButton_2 = QtWidgets.QPushButton(Form)
        font = QtGui.QFont()
        font.setFamily("Arial")
        font.setPointSize(12)
        font.setBold(True)
        font.setWeight(75)
        self.pushButton_2.setFont(font)
        self.pushButton_2.setObjectName("pushButton_2")
        self.horizontalLayout_2.addWidget(self.pushButton_2)
        self.horizontalLayout_2.setStretch(0, 1)
        self.verticalLayout_2.addLayout(self.horizontalLayout_2)
        spacerItem = QtWidgets.QSpacerItem(20, 40, QtWidgets.QSizePolicy.Minimum, QtWidgets.QSizePolicy.Expanding)
        self.verticalLayout_2.addItem(spacerItem)
        self.horizontalLayout = QtWidgets.QHBoxLayout()
        self.horizontalLayout.setObjectName("horizontalLayout")
        self.label = QtWidgets.QLabel(Form)
        font = QtGui.QFont()
        font.setFamily("Arial")
        font.setPointSize(12)
        font.setBold(True)
        font.setWeight(75)
        self.label.setFont(font)
        self.label.setObjectName("label")
        self.horizontalLayout.addWidget(self.label)
        self.comboBox = ComboCheckBox([''])
        font = QtGui.QFont()
        font.setFamily("Arial")
        font.setPointSize(12)
        font.setBold(True)
        font.setWeight(75)
        self.comboBox.setFont(font)
        self.comboBox.setObjectName("comboBox")
        self.horizontalLayout.addWidget(self.comboBox)
        self.horizontalLayout.setStretch(0, 1)
        self.horizontalLayout.setStretch(1, 2)
        self.verticalLayout_2.addLayout(self.horizontalLayout)
        spacerItem1 = QtWidgets.QSpacerItem(20, 40, QtWidgets.QSizePolicy.Minimum, QtWidgets.QSizePolicy.Expanding)
        self.verticalLayout_2.addItem(spacerItem1)
        self.pushButton = QtWidgets.QPushButton(Form)
        font = QtGui.QFont()
        font.setFamily("Arial")
        font.setPointSize(12)
        font.setBold(True)
        font.setWeight(75)
        self.pushButton.setFont(font)
        self.pushButton.setObjectName("pushButton")
        self.verticalLayout_2.addWidget(self.pushButton)
        self.horizontalLayout_3.addLayout(self.verticalLayout_2)

        self.retranslateUi(Form)
        QtCore.QMetaObject.connectSlotsByName(Form)

    def retranslateUi(self, Form):
        _translate = QtCore.QCoreApplication.translate
        Form.setWindowTitle(_translate("Form", "HE-MSI Data Export"))
        self.label_2.setText(_translate("Form", "Data File Save Path:"))
        self.pushButton_2.setText(_translate("Form", "Choose"))
        self.label.setText(_translate("Form", "Target Metabolites:"))
        self.pushButton.setText(_translate("Form", "Start Exporting"))
class My_Combo_choose_Form(QtWidgets.QWidget,Combo_choose_Form):
    def __init__(self):
        super().__init__()
        self.setupUi(self)
        self.pushButton_2.clicked.connect(self.btn_save)

    def btn_save(self):
        try:
            path = os.getcwd()
            file_name, _ = QtWidgets.QFileDialog.getSaveFileName(self, u'Save file', path,'Excel files (*.xls)')
            self.lineEdit.setText(file_name)
        except Exception as e:
            m = 'Running error, info: ' + str(e)
            self.error(m)
class Fusion_Combo_choose_Form(object):
    def setupUi(self, Form):
        Form.setObjectName("Form")
        Form.resize(490, 315)
        sizePolicy = QtWidgets.QSizePolicy(QtWidgets.QSizePolicy.Preferred, QtWidgets.QSizePolicy.Preferred)
        sizePolicy.setHorizontalStretch(0)
        self.setWindowModality(QtCore.Qt.ApplicationModal)
        sizePolicy.setVerticalStretch(0)
        sizePolicy.setHeightForWidth(Form.sizePolicy().hasHeightForWidth())
        Form.setSizePolicy(sizePolicy)
        icon = QtGui.QIcon()
        icon.addPixmap(QtGui.QPixmap(":/newPrefix/ui/icon.png"), QtGui.QIcon.Normal, QtGui.QIcon.Off)
        Form.setWindowIcon(icon)
        self.horizontalLayout_5 = QtWidgets.QHBoxLayout(Form)
        self.horizontalLayout_5.setObjectName("horizontalLayout_5")
        self.verticalLayout = QtWidgets.QVBoxLayout()
        self.verticalLayout.setObjectName("verticalLayout")
        self.horizontalLayout_2 = QtWidgets.QHBoxLayout()
        self.horizontalLayout_2.setObjectName("horizontalLayout_2")
        self.label_2 = QtWidgets.QLabel(Form)
        sizePolicy = QtWidgets.QSizePolicy(QtWidgets.QSizePolicy.Maximum, QtWidgets.QSizePolicy.Preferred)
        sizePolicy.setHorizontalStretch(0)
        sizePolicy.setVerticalStretch(0)
        sizePolicy.setHeightForWidth(self.label_2.sizePolicy().hasHeightForWidth())
        self.label_2.setSizePolicy(sizePolicy)
        font = QtGui.QFont()
        font.setFamily("Arial")
        font.setPointSize(12)
        font.setBold(True)
        font.setWeight(75)
        self.label_2.setFont(font)
        self.label_2.setObjectName("label_2")
        self.horizontalLayout_2.addWidget(self.label_2)
        self.lineEdit = QtWidgets.QLineEdit(Form)
        font = QtGui.QFont()
        font.setFamily("Arial")
        font.setPointSize(12)
        font.setBold(True)
        font.setWeight(75)
        self.lineEdit.setFont(font)
        self.lineEdit.setObjectName("lineEdit")
        self.horizontalLayout_2.addWidget(self.lineEdit)
        self.pushButton_2 = QtWidgets.QPushButton(Form)
        font = QtGui.QFont()
        font.setFamily("Arial")
        font.setPointSize(12)
        font.setBold(True)
        font.setWeight(75)
        self.pushButton_2.setFont(font)
        self.pushButton_2.setObjectName("pushButton_2")
        self.horizontalLayout_2.addWidget(self.pushButton_2)
        self.horizontalLayout_2.setStretch(0, 1)
        self.verticalLayout.addLayout(self.horizontalLayout_2)
        spacerItem = QtWidgets.QSpacerItem(20, 40, QtWidgets.QSizePolicy.Minimum, QtWidgets.QSizePolicy.Expanding)
        self.verticalLayout.addItem(spacerItem)
        self.horizontalLayout = QtWidgets.QHBoxLayout()
        self.horizontalLayout.setObjectName("horizontalLayout")
        self.label = QtWidgets.QLabel(Form)
        font = QtGui.QFont()
        font.setFamily("Arial")
        font.setPointSize(12)
        font.setBold(True)
        font.setWeight(75)
        self.label.setFont(font)
        self.label.setObjectName("label")
        self.horizontalLayout.addWidget(self.label)
        self.comboBox = ComboCheckBox([''])
        font = QtGui.QFont()
        font.setFamily("Arial")
        font.setPointSize(12)
        font.setBold(True)
        font.setWeight(75)
        self.comboBox.setFont(font)
        self.comboBox.setObjectName("comboBox")
        self.horizontalLayout.addWidget(self.comboBox)
        self.horizontalLayout.setStretch(0, 1)
        self.horizontalLayout.setStretch(1, 2)
        self.verticalLayout.addLayout(self.horizontalLayout)
        spacerItem1 = QtWidgets.QSpacerItem(20, 40, QtWidgets.QSizePolicy.Minimum, QtWidgets.QSizePolicy.Expanding)
        self.verticalLayout.addItem(spacerItem1)
        self.horizontalLayout_4 = QtWidgets.QHBoxLayout()
        self.horizontalLayout_4.setObjectName("horizontalLayout_4")
        self.label_3 = QtWidgets.QLabel(Form)
        font = QtGui.QFont()
        font.setFamily("Arial")
        font.setPointSize(12)
        font.setBold(True)
        font.setWeight(75)
        self.label_3.setFont(font)
        self.label_3.setObjectName("label_3")
        self.horizontalLayout_4.addWidget(self.label_3)
        self.horizontalLayout_3 = QtWidgets.QHBoxLayout()
        self.horizontalLayout_3.setObjectName("horizontalLayout_3")
        self.radioButton = QtWidgets.QRadioButton(Form)
        font = QtGui.QFont()
        font.setFamily("Arial")
        font.setPointSize(12)
        font.setBold(True)
        font.setWeight(75)
        self.radioButton.setFont(font)
        self.radioButton.setObjectName("radioButton")
        self.horizontalLayout_3.addWidget(self.radioButton)
        self.radioButton_2 = QtWidgets.QRadioButton(Form)
        font = QtGui.QFont()
        font.setFamily("Arial")
        font.setPointSize(12)
        font.setBold(True)
        font.setWeight(75)
        self.radioButton_2.setFont(font)
        self.radioButton_2.setObjectName("radioButton_2")
        self.horizontalLayout_3.addWidget(self.radioButton_2)
        self.horizontalLayout_4.addLayout(self.horizontalLayout_3)
        self.verticalLayout.addLayout(self.horizontalLayout_4)
        spacerItem2 = QtWidgets.QSpacerItem(20, 40, QtWidgets.QSizePolicy.Minimum, QtWidgets.QSizePolicy.Expanding)
        self.verticalLayout.addItem(spacerItem2)
        self.gridLayout = QtWidgets.QGridLayout()
        self.gridLayout.setContentsMargins(40, -1, 80, -1)
        self.gridLayout.setObjectName("gridLayout")
        self.label_6 = QtWidgets.QLabel(Form)
        font = QtGui.QFont()
        font.setFamily("Arial")
        font.setPointSize(12)
        font.setBold(False)
        font.setWeight(50)
        self.label_6.setFont(font)
        self.label_6.setObjectName("label_6")
        self.gridLayout.addWidget(self.label_6, 2, 1, 1, 1)
        self.lineEdit_3 = QtWidgets.QLineEdit(Form)
        font = QtGui.QFont()
        font.setFamily("Arial")
        font.setPointSize(12)
        font.setBold(True)
        font.setWeight(75)
        self.lineEdit_3.setFont(font)
        self.lineEdit_3.setObjectName("lineEdit_3")
        self.gridLayout.addWidget(self.lineEdit_3, 4, 2, 1, 1)
        self.comboBox_2 = QtWidgets.QComboBox(Form)
        sizePolicy = QtWidgets.QSizePolicy(QtWidgets.QSizePolicy.Preferred, QtWidgets.QSizePolicy.Minimum)
        sizePolicy.setHorizontalStretch(0)
        sizePolicy.setVerticalStretch(0)
        sizePolicy.setHeightForWidth(self.comboBox_2.sizePolicy().hasHeightForWidth())
        self.comboBox_2.setSizePolicy(sizePolicy)
        self.comboBox_2.setMinimumSize(QtCore.QSize(0, 25))
        font = QtGui.QFont()
        font.setFamily("Arial")
        font.setPointSize(12)
        font.setBold(True)
        font.setWeight(75)
        self.comboBox_2.setFont(font)
        self.comboBox_2.setObjectName("comboBox_2")
        self.gridLayout.addWidget(self.comboBox_2, 2, 2, 1, 1)
        self.label_5 = QtWidgets.QLabel(Form)
        font = QtGui.QFont()
        font.setFamily("Arial")
        font.setPointSize(12)
        font.setBold(False)
        font.setWeight(50)
        self.label_5.setFont(font)
        self.label_5.setObjectName("label_5")
        self.gridLayout.addWidget(self.label_5, 3, 1, 1, 1)
        self.lineEdit_2 = QtWidgets.QLineEdit(Form)
        font = QtGui.QFont()
        font.setFamily("Arial")
        font.setPointSize(12)
        font.setBold(True)
        font.setWeight(75)
        self.lineEdit_2.setFont(font)
        self.lineEdit_2.setObjectName("lineEdit_2")
        self.gridLayout.addWidget(self.lineEdit_2, 3, 2, 1, 1)
        self.label_7 = QtWidgets.QLabel(Form)
        font = QtGui.QFont()
        font.setFamily("Arial")
        font.setPointSize(12)
        font.setBold(False)
        font.setWeight(50)
        self.label_7.setFont(font)
        self.label_7.setObjectName("label_7")
        self.gridLayout.addWidget(self.label_7, 4, 1, 1, 1)
        self.label_4 = QtWidgets.QLabel(Form)
        font = QtGui.QFont()
        font.setFamily("Arial")
        font.setPointSize(12)
        font.setBold(True)
        font.setWeight(75)
        self.label_4.setFont(font)
        self.label_4.setObjectName("label_4")
        self.gridLayout.addWidget(self.label_4, 0, 1, 1, 1)
        self.verticalLayout.addLayout(self.gridLayout)
        spacerItem3 = QtWidgets.QSpacerItem(20, 40, QtWidgets.QSizePolicy.Minimum, QtWidgets.QSizePolicy.Expanding)
        self.verticalLayout.addItem(spacerItem3)
        self.pushButton = QtWidgets.QPushButton(Form)
        font = QtGui.QFont()
        font.setFamily("Arial")
        font.setPointSize(12)
        font.setBold(True)
        font.setWeight(75)
        self.pushButton.setFont(font)
        self.pushButton.setObjectName("pushButton")
        self.verticalLayout.addWidget(self.pushButton)
        self.horizontalLayout_5.addLayout(self.verticalLayout)

        self.retranslateUi(Form)
        QtCore.QMetaObject.connectSlotsByName(Form)

    def retranslateUi(self, Form):
        _translate = QtCore.QCoreApplication.translate
        Form.setWindowTitle(_translate("Form", "HE-MSI Data Export"))
        self.label_2.setText(_translate("Form", "MSI Fusion File Path:"))
        self.pushButton_2.setText(_translate("Form", "Choose"))
        self.label.setText(_translate("Form", "Target Metabolites:"))
        self.label_3.setText(_translate("Form", "Choose Algorithm:"))
        self.radioButton.setText(_translate("Form", "PLSR"))
        self.radioButton_2.setText(_translate("Form", "ANN"))
        self.label_6.setText(_translate("Form", "      Hidden Layer:"))
        self.lineEdit_3.setText(_translate("Form", "150"))
        self.label_5.setText(_translate("Form", "      Neuron Num:"))
        self.lineEdit_2.setText(_translate("Form", "100"))
        self.label_7.setText(_translate("Form", "      Iteration Num:"))
        self.label_4.setText(_translate("Form", "Parameter: "))
        self.pushButton.setText(_translate("Form", "Start Exporting"))
class My_Fusion_Combo_choose_Form(QtWidgets.QWidget,Fusion_Combo_choose_Form):
    def __init__(self):
        super().__init__()
        self.setupUi(self)
        self.setWindowTitle('Fusion Image Generate')
        self.label_2.setText('MSI Fusion File Path: ')
        self.pushButton.setText('Start Drawing')
        self.pushButton_2.clicked.connect(self.btn_save)
        self.radioButton_2.setChecked(True)
        self.radioButton.setChecked(False)
        self.radioButton.toggled.connect(lambda: self.rbtn_state(self.radioButton))
        self.radioButton_2.toggled.connect(lambda: self.rbtn_state(self.radioButton_2))
        self.comboBox_2.addItems(['1','2','3','4'])

    def rbtn_state(self,rbtn):
        if rbtn.text() == 'PLSR':
            self.label_5.setText('      n_component:')
            self.label_6.setVisible(False)
            self.label_7.setVisible(False)
            self.comboBox_2.setVisible(False)
            self.lineEdit_2.setText('')
            self.lineEdit_3.setVisible(False)
        else:
            self.label_5.setText('      Neuron Num:')
            self.label_6.setVisible(True)
            self.label_7.setVisible(True)
            self.comboBox_2.setVisible(True)
            self.lineEdit_2.setText('100')
            self.lineEdit_3.setVisible(True)

    def btn_save(self):
        try:
            path = os.getcwd()
            file_name, _ = QtWidgets.QFileDialog.getOpenFileName(self, u'Open file', path,'Npz files (*.npz)')
            self.lineEdit.setText(file_name)
            self.comboBox.Clear_Items()
            meta_name_list_path = file_name.split('.')[0]+'_meta_names.npy'
            meta_name_list = list(np.load(meta_name_list_path))
            self.comboBox.Add_Items(meta_name_list)
        except Exception as e:
            m = 'Running error, info: ' + str(e)
            self.error(m)

def update(number,name):
    global Meta,norm,Predict_Image,cb,cmap3,som_val,score_l
    met_len = len(cb)
    maxval ,minval = float(Predict_Image[number-1].max()),float(Predict_Image[number-1].min())
    s1 = som1.val
    s2 = som2.val
    if name == 's1':
        som_val['min_'+str(number)] = s1
    else:
        som_val['max_'+str(number)] = s2
    vmin = minval+(maxval-minval)*s1/100
    vmax = maxval - (maxval - minval) * (100-s2) / 100
    norm = col.Normalize(vmin = vmin,vmax = vmax)
    if met_len ==1:
        ax_n = figure2.add_subplot(1,1,number)
    elif met_len ==2:
        ax_n = figure2.add_subplot(1,2,number)
    elif met_len ==3 or met_len ==4:
        ax_n = figure2.add_subplot(2,2,number)
    elif met_len ==5 or met_len ==6:
        ax_n = figure2.add_subplot(2,3,number)
    else:
        ax_n = figure2.add_subplot(3,3,number)

    ax_n.clear()
    cb[number-1].remove()
    ax_n.set_title(str(number) + ')' + str(Meta[number - 1])+' Score:' + str(score_l[number - 1]))
    gg = ax_n.imshow(Predict_Image[number-1], cmap=cmap3,norm = norm)
    cb[number - 1] = plt.colorbar(gg, ax = ax_n)
    figure2.canvas.draw()

def PeakIntensitySum(spec,diffl,diffr):
    s=spec[0][np.where(spec[0]<=diffr)]
    ss=spec[1][np.where(s>=diffl)]
    return(ss.sum())

def cmp_change(label):
    global token,Meta_len,cmap3,Predict_Image,cb,som_val
    if token == 'Subtract':
        if label == 're':
            cmap3 = col.LinearSegmentedColormap.from_list('own2', ['#ffffff','#00ebff','#00bfff','#007cd9','#000000','#e95b5b','#fc8b8b','#feabab','#ffffff'])
        elif label == 'ue':
            cmap3 = 'hot'
        elif label == 'een':
            cmap3 = col.LinearSegmentedColormap.from_list('own2', ['#000000','#520068','#1500ff','#00e1ff','#00ff50','#77ff00','#fcff00','#ffc300','#ff0000'])
    else:
        if label == 're':
            cmap3 = 'viridis'
        elif label == 'ue':
            cmap3 = 'hot'
        elif label == 'een':
            cmap3 = col.LinearSegmentedColormap.from_list('own2', ['#ffffff','#00ebff','#00bfff','#007cd9','#000000','#e95b5b','#fc8b8b','#feabab','#ffffff'])
    norm_l = []
    for i in range(0,Meta_len):
        maxval, minval = float(Predict_Image[i].max()), float(Predict_Image[i].min())
        s1=som_val['min_' + str(i+1)]
        s2 = som_val['max_' + str(i + 1)]
        vmin = minval + (maxval - minval) * s1 / 100
        vmax = maxval - (maxval - minval) * (100 - s2) / 100
        norm_l.append(col.Normalize(vmin=vmin, vmax=vmax))
        cb[i].remove()

    cb, p, ax_l = [], [], []
    if Meta_len ==1: sub_inde = [1,1]
    elif Meta_len == 2: sub_inde = [1,2]
    elif Meta_len == 3 or Meta_len == 4: sub_inde = [2,2]
    elif Meta_len == 5 or Meta_len == 6: sub_inde = [2,3]
    elif Meta_len == 7 or Meta_len == 8 or Meta_len == 9: sub_inde = [3,3]

    try:
        for k in range(0, Meta_len):
            ax_l.append(figure2.add_subplot(sub_inde[0], sub_inde[1], k + 1))
            ax_l[k].clear()
            ax_l[k].set_title(str(k+1) + ')' + str(Meta[k]) + ' Score:' + str(score_l[k]))
            p.append(ax_l[k].imshow(Predict_Image[k], cmap=cmap3,norm = norm_l[k]))
            cb.append(plt.colorbar(p[k], ax=ax_l[k]))
    except Exception as e:
        m = 'Running Error, info: ' + str(e)
        print(m)
    figure2.canvas.draw()

def pngfile_path(HE_dir):
    m = os.walk(HE_dir)
    for d, d1, file in m: files = file
    png_files = []
    for file in files:
        if file.split('.')[-1] == 'png':
            png_files.append(HE_dir + '/' + file)
    return png_files

class Shape_Operation():
    def move(self,x,y,tx,ty):
        origin = np.array([x, y, 1])
        mov = np.array([[1,0,tx],
                        [0,1,ty],
                        [0,0,1]])
        return (np.dot(mov,origin)[0],np.dot(mov,origin)[1])

    def reverse(self,x,y,x0,y0):
        origin = np.array([x, y, 1])
        rever = np.array([[1,0,0],
                          [0,-1,2*y0],
                          [0,0,1]])
        return (np.dot(rever,origin)[0],np.dot(rever,origin)[1])

    def zuoyoureverse(self,x,y,x0,y0):
        origin = np.array([x, y, 1])
        rever = np.array([[-1,0,2*x0],
                          [0,1,0],
                          [0,0,1]])
        return (np.dot(rever,origin)[0],np.dot(rever,origin)[1])

    def rotate(self,x,y,x0,y0,angle):
        origin = np.array([x,y,1])
        angle = angle*np.pi/180
        rot = np.array([[np.cos(angle),-np.sin(angle),(1-np.cos(angle))*x0+np.sin(angle)*y0],
                        [np.sin(angle),np.cos(angle),(1-np.cos(angle))*y0-np.sin(angle)*x0],
                        [0,0,1]])
        return (np.dot(rot,origin)[0],np.dot(rot,origin)[1])

    def lmove(self,xlist,ylist,tx,ty):
        xlist1,ylist1 = [], []
        for i in range(0,len(xlist)):
            x1 , y1 = self.move(xlist[i],ylist[i],tx,ty)
            xlist1.append(x1)
            ylist1.append(y1)
        print('Move Operation: ','dx: ',tx,' dy: ',ty)
        return xlist1,ylist1

    def lrotate(self,xlist,ylist,x0,y0,angle):
        xlist1,ylist1 = [], []
        for i in range(0,len(xlist)):
            x1 , y1 = self.rotate(xlist[i],ylist[i],x0,y0,angle)
            xlist1.append(x1)
            ylist1.append(y1)
        print('Rotation Operation: ', 'Center x0: ', x0, ' Center y0: ', y0,' Angle(clockwise positive): ',angle)
        return xlist1,ylist1

    def lreverse(self,xlist,ylist,x0,y0):
        xlist1,ylist1 = [], []
        for i in range(0,len(xlist)):
            x1 , y1 = self.reverse(xlist[i],ylist[i],x0,y0)
            xlist1.append(x1)
            ylist1.append(y1)
        print('Up-Down Flip: ', 'x0: ', x0, '  y0: ', y0)
        return xlist1,ylist1

    def lzuoyoureverse(self,xlist,ylist,x0,y0):
        xlist1,ylist1 = [], []
        for i in range(0,len(xlist)):
            x1 , y1 = self.zuoyoureverse(xlist[i],ylist[i],x0,y0)
            xlist1.append(x1)
            ylist1.append(y1)
        print('Left-Right Flip: ', 'x0: ', x0, '  y0: ', y0)
        return xlist1,ylist1

class MyButton(Button):
    def __init__(self, ax, label='', image=None, image_pressed = None,color='0.85', hovercolor='0.95'):
        super(MyButton, self).__init__(ax, label, image, color, hovercolor)
        self.ax ,self.image,self.image_pressed = ax , image,image_pressed
        self.connect_event('axes_enter_event', self.change1)
        self.connect_event('axes_leave_event', self.change2)

    def change1(self,event):
        if self.image is not None:
            if event.inaxes == self.ax:
                self.ax.imshow(self.image_pressed)
    def change2(self,event):
        if self.image_pressed is not None:
            if event.inaxes == self.ax:
                self.ax.imshow(self.image)

def points_confirm(x,y,point_data):
    c=np.where(point_data[:,0]==x)[0]
    m=np.where(point_data[c,1]==y)[0]
    if len(m)==0:
        return -1
    else:
        return c[m[0]]

def list_points_confirm(point_data,origin_data):
    l= []
    for i in range(0,len(point_data)):
        x,y = point_data[i,0],point_data[i,1]
        tmp = points_confirm(x,y,origin_data)
        if tmp!=-1:
            l.append(tmp+1)
    l = np.array(l)
    return l

def pngfile_path(HE_dir):
    m = os.walk(HE_dir)
    for d, d1, file in m: files = file
    png_files = []
    for file in files:
        if file.split('.')[-1] == 'png':
            png_files.append(HE_dir + '/' + file)
    return png_files

def origin_export_list(point_list,xlspath):
    data = xlrd.open_workbook(xlspath)
    table = data.sheets()[0]
    nrows = table.nrows

    result = []
    for i in range(0,len(point_list)):
        result.append(table.row_values(point_list[i]))

def MSI_xls_data(biaoji,MSI_dir,index):
    global subtract_points,xmin,ymin
    data = xlrd.open_workbook(MSI_dir)
    table = data.sheets()[0]
    nrows = table.nrows

    if biaoji == 'Original':
        add_n = 5
    elif biaoji == 'Reflect':
        add_n = 6
    elif biaoji == 'Subtract':
        add_n = 7
    subtract_points = np.array([[table.cell_value(1, 1), table.cell_value(1, 2)]])
    subtract_intensity = np.array([table.cell_value(1, 3*index+add_n)])

    for i in range(2, nrows):
        subtract_points = np.append(subtract_points, [[table.cell_value(i, 1), table.cell_value(i, 2)]], axis=0)
        subtract_intensity = np.append(subtract_intensity, [table.cell_value(i, 3*index+add_n)], axis=0)

    total_points = []
    total_intensity = []

    if biaoji == 'Original' or biaoji == 'Reflect':
        blank_intensity = 0
    else:
        if abs(subtract_intensity).max() > subtract_intensity.max():
            blank_intensity = abs(subtract_intensity).max()
        else:
            blank_intensity = -subtract_intensity.max()

    xmin, xmax = int(subtract_points[:, 0].min()), int(subtract_points[:, 0].max())

    ymin, ymax = int(subtract_points[:, 1].min()), int(subtract_points[:, 1].max())

    for x in range(xmin - 5, xmax + 5):
        for y in range(ymin - 5, ymax + 5):
            index = points_confirm(x, y, subtract_points)
            total_points.append([x, y])
            if index == -1:
                total_intensity.append(blank_intensity)
            else:
                total_intensity.append(subtract_intensity[index])
    total_points = np.array(total_points)
    total_intensity = np.array(total_intensity)
    grid_x, grid_y = np.mgrid[(xmin - 5):(xmax + 4):(xmax - xmin + 10) * 10j,
                     (ymin - 5):(ymax + 4):(ymax - ymin + 10) * 10j]
    grid_z0 = griddata(total_points, total_intensity, (grid_x, grid_y), method='linear')
    MSI_points, MSI_intensity = [], []
    for i in range(0, (xmax - xmin + 10) * 10):
        for j in range(0, (ymax - ymin + 10) * 10):
            MSI_points.append([i, j])
            MSI_intensity.append(grid_z0[i, j])
    #MSI_points, MSI_intensity = np.array(MSI_points), np.array(MSI_intensity)
    #MSI_shape = np.shape(grid_z0)
    MSI_points, MSI_intensity = total_points, total_intensity
    MSI_shape = (xmax - xmin + 10, ymax - ymin + 10)
    return xmin,xmax,ymin,ymax,MSI_points, MSI_intensity,grid_z0,MSI_shape

def MSI_imzml_data(MSI_dir,left,right):
    p = ImzMLParser(MSI_dir)
    Coor = p.coordinates
    total = len(Coor)

    m = p.getspectrum(0)
    points = np.array([list(Coor[0][:-1])])
    Tempinte = PeakIntensitySum(m, left, right)
    intensity = np.array([Tempinte])
    for indecount in range(1, total):
        m = p.getspectrum(indecount)
        points = np.append(points, [list(Coor[indecount][:-1])], axis=0)
        intensity = np.append(intensity, [PeakIntensitySum(m, left, right)], axis=0)

    xmax = np.array(Coor)[:, 0].max()
    ymax = np.array(Coor)[:, 1].max()
    xmin = np.array(Coor)[:, 0].min()
    ymin = np.array(Coor)[:, 1].min()

    grid_x, grid_y = np.mgrid[1:xmax:(xmax * 10j), 1:ymax:(ymax * 10j)]

    grid_z0 = griddata(points, intensity, (grid_x, grid_y), method='linear')
    MSI_points, MSI_intensity = [], []
    for i in range(0, xmax  * 10):
        for j in range(0, ymax * 10):
            MSI_points.append([i, j])
            MSI_intensity.append(grid_z0[i, j])
    MSI_points, MSI_intensity = np.array(MSI_points), np.array(MSI_intensity)
    MSI_shape = np.shape(grid_z0)

    return xmin,xmax,ymin,ymax,MSI_points, MSI_intensity,grid_z0,MSI_shape

class ButtonHandler():
    def __init__(self,fig,MSimg,HEimg,MSI_points,HE_points,MSI_intensity,HE_intensity,MSI_shape,HE_shape,xls_path):
        self.fig = fig
        self.loc = np.array([[0, 0]])
        self.x, self.y = [], []
        self.xHE, self.yHE = [], []
        self.xls_path = xls_path
        self.kuang_x, self.kuang_y = [], []
        self.operation = []
        self.shape_operation = Shape_Operation()
        self.MSimg = MSimg
        self.HEimg = HEimg
        self.MSI_points = MSI_points
        self.HE_points = HE_points
        self.MSI_intensity = MSI_intensity
        self.HE_intensity = HE_intensity
        self.MSI_shape = MSI_shape
        self.HE_shape = HE_shape

    def connect(self):
        self.cidclick = self.fig.canvas.mpl_connect('button_press_event', self.on_click)
        self.cidpress = self.fig.canvas.mpl_connect('key_press_event', self.on_press)
        self.cidmotion = self.fig.canvas.mpl_connect('motion_notify_event',self.on_motion)
        self.cidrelease = self.fig.canvas.mpl_connect('key_release_event', self.on_release)

    def error(self,m):
        self.eW=My_Error_Form()
        self.eW.label.setText(m)
        self.eW.show()

    def on_entering(self,event):
        if event.inaxes == MSimg.axes:
            print('sssd')
            if self.x!=[] and self.y!=[]:
                if max(self.x) >= float(event.xdata) >= float(min(self.x)):
                    if abs(float(event.ydata)-max(self.y))<=25:
                        self.fig.canvas.setCursor(QtCore.Qt.CrossCursor)

    def reelect(self,event):
        global cmap2,token
        ax = self.fig.add_subplot(1, 2, 1)
        ax.clear()
        if token == 'Original' or token == 'Reflect':
            MSimg = plt.imshow(grid_z0.T, extent=(xmin - 5, xmax + 4, ymax + 4, ymin - 5))
        else:
            MSimg = plt.imshow(grid_z0.T, cmap=cmap2, extent=(xmin - 5, xmax + 4, ymax + 4, ymin - 5))
        ax2 = self.fig.add_subplot(1, 2, 2)
        ax2.clear()
        plt.imshow(self.HEimg)
        self.x ,self.xHE= [],[]
        self.y,self.yHE= [],[]
        self.loc = np.array([[0, 0]])

    def close_window(self,val):
        if val==100:
            self.msg.close()

    def on_click(self,event):
        global token
        if event.button == 1 and event.inaxes==MSimg.axes and event.key is None:
            if event.xdata != None and event.ydata != None :
                x = float(event.xdata)
                y = float(event.ydata)
                if x>1 and y>1:
                    self.x.append(x)
                    self.y.append(y)
                    self.loc=np.append(self.loc,[[x,y]],axis=0)

                ax = self.fig.add_subplot(1,2,1)
                ax.lines.clear()
                self.x.append(self.x[0])
                self.y.append(self.y[0])
                ax.plot(self.x[0],self.y[0], 'r+', ms = 20)
                ax.plot(self.x[1:], self.y[1:], 'r*')
                ax.plot(self.x,self.y,'g',linestyle='-.')
                self.x = self.x[:-1]
                self.y = self.y[:-1]
                self.fig.canvas.draw()
        if event.button == 3 and event.inaxes==MSimg.axes and event.key is None:
            id = -1
            for i in range(0, len(self.x)):
                distx = event.xdata - self.x[i]
                disty = event.ydata - self.y[i]
                if distx ** 2 < 10 and disty ** 2 < 10:
                    id = i
                    break
            if id != -1:
                self.x = self.x[0:id] + self.x[id+1:]
                self.y = self.y[0:id] + self.y[id+1:]
                ax = self.fig.add_subplot(1, 2, 1)
                ax.lines.clear()
                self.x.append(self.x[0])
                self.y.append(self.y[0])
                ax.plot(self.x[0], self.y[0], 'r+', ms=20)
                ax.plot(self.x[1:], self.y[1:], 'r*')
                ax.plot(self.x, self.y, 'g', linestyle='-.')
                self.x = self.x[:-1]
                self.y = self.y[:-1]
                self.fig.canvas.draw()
        if event.button == 1 and event.inaxes==HEimg.axes and event.key is None:
            if event.xdata != None and event.ydata != None :
                x = float(event.xdata)
                y = float(event.ydata)
                if x>1 and y>1:
                    self.xHE.append(x)
                    self.yHE.append(y)
                    self.loc=np.append(self.loc,[[x,y]],axis=0)

                ax2 = self.fig.add_subplot(1,2,2)
                ax2.lines.clear()
                self.xHE.append(self.xHE[0])
                self.yHE.append(self.yHE[0])
                ax2.plot(self.xHE[0],self.yHE[0], 'r+', ms = 20)
                ax2.plot(self.xHE[1:], self.yHE[1:], 'r*')
                ax2.plot(self.xHE,self.yHE,'g',linestyle='-.')
                self.xHE = self.xHE[:-1]
                self.yHE = self.yHE[:-1]
                self.fig.canvas.draw()
        if event.button == 3 and event.inaxes==HEimg.axes and event.key is None:
            id = -1
            for i in range(0, len(self.xHE)):
                distx = event.xdata - self.xHE[i]
                disty = event.ydata - self.yHE[i]
                if distx ** 2 < 10 and disty ** 2 < 10:
                    id = i
                    break
            if id != -1:
                self.xHE = self.xHE[0:id] + self.xHE[id+1:]
                self.yHE = self.yHE[0:id] + self.yHE[id+1:]
                ax2 = self.fig.add_subplot(1, 2, 2)
                ax2.lines.clear()
                self.xHE.append(self.xHE[0])
                self.yHE.append(self.yHE[0])
                ax2.plot(self.xHE[0], self.yHE[0], 'r+', ms=20)
                ax2.plot(self.xHE[1:], self.yHE[1:], 'r*')
                ax2.plot(self.xHE, self.yHE, 'g', linestyle='-.')
                self.xHE = self.xHE[:-1]
                self.yHE = self.yHE[:-1]
                self.fig.canvas.draw()

    def RemoveLearning(self,event):
        global cmap2
        ax = self.fig.add_subplot(1, 2, 1)
        ax.clear()
        if token == 'Original' or token == 'Reflect':
            MSimg = plt.imshow(grid_z0.T, extent=(xmin - 5, xmax + 4, ymax + 4, ymin - 5))
        else:
            MSimg = plt.imshow(grid_z0.T, cmap=cmap2, extent=(xmin - 5, xmax + 4, ymax + 4, ymin - 5))
        self.x = []
        self.y = []
        self.loc = np.array([[0, 0]])
        self.operation = []
        self.msg = My_Message_Form()
        self.msg.setWindowFlags(QtCore.Qt.FramelessWindowHint)
        self.msg.label.setText(u'Fusion Recording Data Removed Successfully!')
        self.msg.show()
        self.clo = close_widget_thread(1)
        self.clo.start()
        self.clo.trigger.connect(self.close_window)

    def Operation_import(self,event):
        try:
            data = xlrd.open_workbook('Fusion Recording Data.xls')
            table = data.sheets()[0]
            nrows = table.nrows
            temp = []
            for i in range(0,nrows):
                m=table.row_values(i)
                rtemp = []
                for j in m:
                    if j!='':
                        if type(j) != str:
                            rtemp.append(float(j))
                        else:
                            rtemp.append(j)
                print(rtemp)
                temp.append(rtemp)
            self.operation = temp
            print(u'Fusion Recording Data Imported Successfully!')
            self.msg = My_Message_Form()
            self.msg.setWindowFlags(QtCore.Qt.FramelessWindowHint)
            self.msg.label.setText(u'Fusion Recording Data Imported Successfully!')
            self.msg.show()
            self.clo = close_widget_thread(1)
            self.clo.start()
            self.clo.trigger.connect(self.close_window)
        except Exception as e :
            m='运行错误，错误信息：'+str(e)
            self.error(m)

    def Operation_export(self,event):
        f = xlwt.Workbook()
        sheet1 = f.add_sheet(u'sheet1', cell_overwrite_ok=True)
        if self.operation is not None:
            for i in range(0,len(self.operation)):
                for j in range(0,len(self.operation[i])):
                    sheet1.write(i,j,self.operation[i][j])
        f.save('Fusion Recording Data.xls')
        print('Fusion Recording Data Exported Successfully!')
        self.msg = My_Message_Form()
        self.msg.setWindowFlags(QtCore.Qt.FramelessWindowHint)
        self.msg.label.setText(u'Fusion Recording Data Exported Successfully!')
        self.msg.show()
        self.clo = close_widget_thread(1)
        self.clo.start()
        self.clo.trigger.connect(self.close_window)

    def on_press(self,event):
        global token,MSI_shape,MSimg
        if event.key == 'control':
            self.cidkuang = self.fig.canvas.mpl_connect('button_press_event', self.kuang)
        if event.key in ['up','down','left','right','alt+up','alt+down','alt+left','alt+right']:
            dist = MSI_shape[0]/20
            if event.key == 'up':
                xmove ,ymove= 0,-dist
            elif event.key == 'down':
                xmove, ymove = 0, dist
            elif event.key == 'left':
                xmove, ymove = -dist, 0
            elif event.key == 'right':
                xmove, ymove = dist, 0
            elif event.key == 'alt+up':
                xmove, ymove = 0, -dist*0.1
            elif event.key == 'alt+down':
                xmove, ymove = 0, dist*0.1
            elif event.key == 'alt+left':
                xmove, ymove = -dist*0.1, 0
            elif event.key == 'alt+right':
                xmove, ymove = dist*0.1, 0
            if event.inaxes == MSimg.axes:
                self.x, self.y = self.shape_operation.lmove(self.x, self.y, xmove, ymove)
                self.operation.append([xmove, ymove])
                x_max, x_min, y_max, y_min = max(self.x), min(self.x), max(self.y), min(self.y)
                self.kuang_x = [x_min, x_min, x_min, (x_min + x_max) / 2, x_max, x_max, x_max, (x_min + x_max) / 2,
                                x_min]
                self.kuang_y = [y_min, (y_min + y_max) / 2, y_max, y_max, y_max, (y_min + y_max) / 2, y_min, y_min,
                                y_min]
                ax = self.fig.add_subplot(1, 2, 1)
                ax.clear()
                if token == 'Original' or token == 'Reflect':
                    MSimg = plt.imshow(grid_z0.T, extent=(xmin - 5, xmax + 4, ymax + 4, ymin - 5))
                else:
                    MSimg = plt.imshow(grid_z0.T, cmap=cmap2, extent=(xmin - 5, xmax + 4, ymax + 4, ymin - 5))
                self.x.append(self.x[0])
                self.y.append(self.y[0])
                ax.plot(self.x[0], self.y[0], 'r+', ms=20)
                ax.plot(self.x[1:], self.y[1:], 'r*')
                ax.plot(self.x, self.y, 'g', linestyle='-.')
                ax.plot(self.kuang_x, self.kuang_y, color='gold', marker='o', ms=10)
                ax.plot(self.kuang_x, self.kuang_y, 'gold', linestyle='-.')
                ax.plot((x_min + x_max) / 2, (y_min + y_max) / 2, color = 'gold',marker = '*', ms=10)
                self.x = self.x[:-1]
                self.y = self.y[:-1]
                self.fig.canvas.draw()

    def on_motion(self,event):
        global xmin, xmax, ymin, ymax
        subtract_value_x = (xmax - xmin)/10
        subtract_value_y = (ymax - ymin)/10
        if event.key=='shift' and event.inaxes == MSimg.axes:
            if self.x!=[] and self.y!=[]:
                if max(self.x) >= float(event.xdata) >= float(min(self.x)):
                    if abs(float(event.ydata)-max(self.y))<=subtract_value_y:
                        self.fig.canvas.setCursor(QtCore.Qt.SizeVerCursor)
                    if abs(float(event.ydata) - min(self.y)) <= subtract_value_y:
                        self.fig.canvas.setCursor(QtCore.Qt.SizeVerCursor)
                if max(self.y) >= float(event.ydata) >= float(min(self.y)):
                    if abs(float(event.xdata) - max(self.x)) <= subtract_value_x:
                        self.fig.canvas.setCursor(QtCore.Qt.SizeHorCursor)
                    if abs(float(event.xdata) - min(self.x)) <= subtract_value_x:
                        self.fig.canvas.setCursor(QtCore.Qt.SizeHorCursor)
        else:
            self.fig.canvas.setCursor(QtCore.Qt.ArrowCursor)
        if event.key=='shift' and event.button == 1 and event.inaxes == MSimg.axes:
            if max(self.x) >= float(event.xdata) >= float(min(self.x)):
                if abs(float(event.ydata)-max(self.y))<=subtract_value_y:
                    ratio = (float(event.ydata)-max(self.y))/(max(self.y)-min(self.y))+1
                    self.operation.append(['ymax',ratio])
                    print('Scale operation: ','Target Image：MSI  Scaling control axis: ymax  Ratio: ',ratio)
                    self.y = list((np.array(self.y)-min(self.y))*ratio+min(self.y))
                    x_max, x_min, y_max, y_min = max(self.x), min(self.x), max(self.y), min(self.y)
                    self.kuang_x = [x_min, x_min, x_min, (x_min + x_max) / 2, x_max, x_max, x_max, (x_min + x_max) / 2,
                                    x_min]
                    self.kuang_y = [y_min, (y_min + y_max) / 2, y_max, y_max, y_max, (y_min + y_max) / 2, y_min, y_min,
                                    y_min]
                    ax = self.fig.add_subplot(1, 2, 1)
                    ax.lines.clear()
                    self.x.append(self.x[0])
                    self.y.append(self.y[0])
                    ax.plot(self.x[0], self.y[0], 'r+', ms=20)
                    ax.plot(self.x[1:], self.y[1:], 'r*')
                    ax.plot(self.x, self.y, 'g', linestyle='-.')
                    ax.plot(self.kuang_x, self.kuang_y, color='gold', marker='o', ms=10)
                    ax.plot(self.kuang_x, self.kuang_y, 'gold', linestyle='-.')
                    ax.plot((x_min + x_max) / 2, (y_min + y_max) / 2, color='gold', marker='*', ms=10)
                    self.x = self.x[:-1]
                    self.y = self.y[:-1]
                    self.fig.canvas.draw()
                if abs(float(event.ydata)-min(self.y))<=subtract_value_y:
                    ratio = (-float(event.ydata) + min(self.y)) / (max(self.y) - min(self.y)) + 1
                    self.operation.append(['ymin', ratio])
                    print('Scale operation: ', 'Target Image：MSI  Scaling control axis: ymin  Ratio: ', ratio)
                    self.y = list((np.array(self.y)-max(self.y))*ratio+max(self.y))
                    x_max, x_min, y_max, y_min = max(self.x), min(self.x), max(self.y), min(self.y)
                    self.kuang_x = [x_min, x_min, x_min, (x_min + x_max) / 2, x_max, x_max, x_max, (x_min + x_max) / 2,
                                    x_min]
                    self.kuang_y = [y_min, (y_min + y_max) / 2, y_max, y_max, y_max, (y_min + y_max) / 2, y_min, y_min,
                                    y_min]
                    ax = self.fig.add_subplot(1, 2, 1)
                    ax.lines.clear()
                    self.x.append(self.x[0])
                    self.y.append(self.y[0])
                    ax.plot(self.x[0], self.y[0], 'r+', ms=20)
                    ax.plot(self.x[1:], self.y[1:], 'r*')
                    ax.plot(self.x, self.y, 'g', linestyle='-.')
                    ax.plot(self.kuang_x, self.kuang_y, color='gold', marker='o', ms=10)
                    ax.plot(self.kuang_x, self.kuang_y, 'gold', linestyle='-.')
                    ax.plot((x_min + x_max) / 2, (y_min + y_max) / 2, color='gold', marker='*', ms=10)
                    self.x = self.x[:-1]
                    self.y = self.y[:-1]
                    self.fig.canvas.draw()
            if max(self.y) >= float(event.ydata) >= float(min(self.y)):
                if abs(float(event.xdata)-max(self.x))<=subtract_value_x:
                    ratio = (float(event.xdata)-max(self.x))/(max(self.x)-min(self.x))+1
                    self.operation.append(['xmax', ratio])
                    print('Scale operation: ', 'Target Image：MSI  Scaling control axis: xmax  Ratio: ', ratio)
                    self.x = list((np.array(self.x)-min(self.x))*ratio+min(self.x))
                    x_max, x_min, y_max, y_min = max(self.x), min(self.x), max(self.y), min(self.y)
                    self.kuang_x = [x_min, x_min, x_min, (x_min + x_max) / 2, x_max, x_max, x_max, (x_min + x_max) / 2,
                                    x_min]
                    self.kuang_y = [y_min, (y_min + y_max) / 2, y_max, y_max, y_max, (y_min + y_max) / 2, y_min, y_min,
                                    y_min]
                    ax = self.fig.add_subplot(1, 2, 1)
                    ax.lines.clear()
                    self.x.append(self.x[0])
                    self.y.append(self.y[0])
                    ax.plot(self.x[0], self.y[0], 'r+', ms=20)
                    ax.plot(self.x[1:], self.y[1:], 'r*')
                    ax.plot(self.x, self.y, 'g', linestyle='-.')
                    ax.plot(self.kuang_x, self.kuang_y, color='gold', marker='o', ms=10)
                    ax.plot(self.kuang_x, self.kuang_y, 'gold', linestyle='-.')
                    ax.plot((x_min + x_max) / 2, (y_min + y_max) / 2, color='gold', marker='*', ms=10)
                    self.x = self.x[:-1]
                    self.y = self.y[:-1]
                    self.fig.canvas.draw()
                if abs(float(event.xdata)-min(self.x))<=subtract_value_x:
                    ratio = (-float(event.xdata) + min(self.x)) / (max(self.x) - min(self.x)) + 1
                    self.operation.append(['xmin', ratio])
                    print('Scale operation: ', 'Target Image：MSI  Scaling control axis: xmin  Ratio: ', ratio)
                    self.x = list((np.array(self.x)-max(self.x))*ratio+max(self.x))
                    x_max, x_min, y_max, y_min = max(self.x), min(self.x), max(self.y), min(self.y)
                    self.kuang_x = [x_min, x_min, x_min, (x_min + x_max) / 2, x_max, x_max, x_max, (x_min + x_max) / 2,
                                    x_min]
                    self.kuang_y = [y_min, (y_min + y_max) / 2, y_max, y_max, y_max, (y_min + y_max) / 2, y_min, y_min,
                                    y_min]
                    ax = self.fig.add_subplot(1, 2, 1)
                    ax.lines.clear()
                    self.x.append(self.x[0])
                    self.y.append(self.y[0])
                    ax.plot(self.x[0], self.y[0], 'r+', ms=20)
                    ax.plot(self.x[1:], self.y[1:], 'r*')
                    ax.plot(self.x, self.y, 'g', linestyle='-.')
                    ax.plot(self.kuang_x, self.kuang_y, color='gold', marker='o', ms=10)
                    ax.plot(self.kuang_x, self.kuang_y, 'gold', linestyle='-.')
                    ax.plot((x_min + x_max) / 2, (y_min + y_max) / 2, color='gold', marker='*', ms=10)
                    self.x = self.x[:-1]
                    self.y = self.y[:-1]
                    self.fig.canvas.draw()
        if event.key=='shift' and event.button == 1 and event.inaxes == HEimg.axes:
            if max(self.xHE) >= float(event.xdata) >= float(min(self.xHE)):
                if abs(float(event.ydata)-max(self.yHE))<=20:
                    ratio = (float(event.ydata)-max(self.yHE))/(max(self.yHE)-min(self.yHE))+1
                    self.operation.append(['yHEmax',ratio])
                    print('Scale operation: ', 'Target Image：HE  Scaling control axis: ymax  Ratio: ', ratio)
                    self.yHE = list((np.array(self.yHE)-min(self.yHE))*ratio+min(self.yHE))
                    x_max, x_min, y_max, y_min = max(self.xHE), min(self.xHE), max(self.yHE), min(self.yHE)
                    self.kuang_x = [x_min, x_min, x_min, (x_min + x_max) / 2, x_max, x_max, x_max, (x_min + x_max) / 2,
                                    x_min]
                    self.kuang_y = [y_min, (y_min + y_max) / 2, y_max, y_max, y_max, (y_min + y_max) / 2, y_min, y_min,
                                    y_min]
                    ax2 = self.fig.add_subplot(1, 2, 2)
                    ax2.lines.clear()
                    self.xHE.append(self.xHE[0])
                    self.yHE.append(self.yHE[0])
                    ax2.plot(self.xHE[0], self.yHE[0], 'r+', ms=20)
                    ax2.plot(self.xHE[1:], self.yHE[1:], 'r*')
                    ax2.plot(self.xHE, self.yHE, 'g', linestyle='-.')
                    ax2.plot(self.kuang_x, self.kuang_y, color='gold', marker='o', ms=10)
                    ax2.plot(self.kuang_x, self.kuang_y, 'gold', linestyle='-.')
                    ax2.plot((x_min + x_max) / 2, (y_min + y_max) / 2, color='gold', marker='*', ms=10)
                    self.xHE = self.xHE[:-1]
                    self.yHE = self.yHE[:-1]
                    self.fig.canvas.draw()
                if abs(float(event.ydata)-min(self.yHE))<=20:
                    ratio = (-float(event.ydata) + min(self.yHE)) / (max(self.yHE) - min(self.yHE)) + 1
                    self.operation.append(['yHEmin', ratio])
                    print('Scale operation: ', 'Target Image：HE  Scaling control axis: ymin  Ratio: ', ratio)
                    self.yHE = list((np.array(self.yHE)-max(self.yHE))*ratio+max(self.yHE))
                    x_max, x_min, y_max, y_min = max(self.xHE), min(self.xHE), max(self.yHE), min(self.yHE)
                    self.kuang_x = [x_min, x_min, x_min, (x_min + x_max) / 2, x_max, x_max, x_max, (x_min + x_max) / 2,
                                    x_min]
                    self.kuang_y = [y_min, (y_min + y_max) / 2, y_max, y_max, y_max, (y_min + y_max) / 2, y_min, y_min,
                                    y_min]
                    ax2 = self.fig.add_subplot(1, 2, 2)
                    ax2.lines.clear()
                    self.xHE.append(self.xHE[0])
                    self.yHE.append(self.yHE[0])
                    ax2.plot(self.xHE[0], self.yHE[0], 'r+', ms=20)
                    ax2.plot(self.xHE[1:], self.yHE[1:], 'r*')
                    ax2.plot(self.xHE, self.yHE, 'g', linestyle='-.')
                    ax2.plot(self.kuang_x, self.kuang_y, color='gold', marker='o', ms=10)
                    ax2.plot(self.kuang_x, self.kuang_y, 'gold', linestyle='-.')
                    ax2.plot((x_min + x_max) / 2, (y_min + y_max) / 2, color='gold', marker='*', ms=10)
                    self.xHE = self.xHE[:-1]
                    self.yHE = self.yHE[:-1]
                    self.fig.canvas.draw()
            if max(self.yHE) >= float(event.ydata) >= float(min(self.yHE)):
                if abs(float(event.xdata)-max(self.xHE))<=20:
                    ratio = (float(event.xdata)-max(self.xHE))/(max(self.xHE)-min(self.xHE))+1
                    self.operation.append(['xHEmax', ratio])
                    print('Scale operation: ', 'Target Image：HE  Scaling control axis: xmax  Ratio: ', ratio)
                    self.xHE = list((np.array(self.xHE)-min(self.xHE))*ratio+min(self.xHE))
                    x_max, x_min, y_max, y_min = max(self.xHE), min(self.xHE), max(self.yHE), min(self.yHE)
                    self.kuang_x = [x_min, x_min, x_min, (x_min + x_max) / 2, x_max, x_max, x_max, (x_min + x_max) / 2,
                                    x_min]
                    self.kuang_y = [y_min, (y_min + y_max) / 2, y_max, y_max, y_max, (y_min + y_max) / 2, y_min, y_min,
                                    y_min]
                    ax2 = self.fig.add_subplot(1, 2, 2)
                    ax2.lines.clear()
                    self.xHE.append(self.xHE[0])
                    self.yHE.append(self.yHE[0])
                    ax2.plot(self.xHE[0], self.yHE[0], 'r+', ms=20)
                    ax2.plot(self.xHE[1:], self.yHE[1:], 'r*')
                    ax2.plot(self.xHE, self.yHE, 'g', linestyle='-.')
                    ax2.plot(self.kuang_x, self.kuang_y, color='gold', marker='o', ms=10)
                    ax2.plot(self.kuang_x, self.kuang_y, 'gold', linestyle='-.')
                    ax2.plot((x_min + x_max) / 2, (y_min + y_max) / 2, color='gold', marker='*', ms=10)
                    self.xHE = self.xHE[:-1]
                    self.yHE = self.yHE[:-1]
                    self.fig.canvas.draw()
                if abs(float(event.xdata)-min(self.xHE))<=20:
                    ratio = (-float(event.xdata) + min(self.xHE)) / (max(self.xHE) - min(self.xHE)) + 1
                    self.operation.append(['xHEmin', ratio])
                    print('Scale operation: ', 'Target Image：HE  Scaling control axis: xmin  Ratio: ', ratio)
                    self.xHE = list((np.array(self.xHE)-max(self.xHE))*ratio+max(self.xHE))
                    x_max, x_min, y_max, y_min = max(self.xHE), min(self.xHE), max(self.yHE), min(self.yHE)
                    self.kuang_x = [x_min, x_min, x_min, (x_min + x_max) / 2, x_max, x_max, x_max, (x_min + x_max) / 2,
                                    x_min]
                    self.kuang_y = [y_min, (y_min + y_max) / 2, y_max, y_max, y_max, (y_min + y_max) / 2, y_min, y_min,
                                    y_min]
                    ax2 = self.fig.add_subplot(1, 2, 2)
                    ax2.lines.clear()
                    self.xHE.append(self.xHE[0])
                    self.yHE.append(self.yHE[0])
                    ax2.plot(self.xHE[0], self.yHE[0], 'r+', ms=20)
                    ax2.plot(self.xHE[1:], self.yHE[1:], 'r*')
                    ax2.plot(self.xHE, self.yHE, 'g', linestyle='-.')
                    ax2.plot(self.kuang_x, self.kuang_y, color='gold', marker='o', ms=10)
                    ax2.plot(self.kuang_x, self.kuang_y, 'gold', linestyle='-.')
                    ax2.plot((x_min + x_max) / 2, (y_min + y_max) / 2, color='gold', marker='*', ms=10)
                    self.xHE = self.xHE[:-1]
                    self.yHE = self.yHE[:-1]
                    self.fig.canvas.draw()

    def on_release(self,event):
        pass

    def kuang(self,event):
        if event.button == 1 and event.inaxes==MSimg.axes and event.key == 'control':
            x_max , x_min ,y_max ,y_min = max(self.x), min(self.x), max(self.y), min(self.y)
            self.kuang_x = [x_min,x_min,x_min,(x_min+x_max)/2,x_max,x_max,x_max,(x_min+x_max)/2,x_min]
            self.kuang_y = [y_min,(y_min+y_max)/2,y_max,y_max,y_max,(y_min+y_max)/2,y_min,y_min,y_min]
            ax = self.fig.add_subplot(1, 2, 1)
            ax.plot(self.kuang_x,self.kuang_y,color = 'gold',marker = 'o',ms = 10)
            ax.plot(self.kuang_x,self.kuang_y,'gold',linestyle = '-.')
            ax.plot((x_min+x_max)/2,(y_min+y_max)/2,color = 'gold',marker = '*',ms = 10)
            self.fig.canvas.draw()
        if event.button == 1 and event.inaxes==HEimg.axes and event.key == 'control':
            x_max , x_min ,y_max ,y_min = max(self.xHE), min(self.xHE), max(self.yHE), min(self.yHE)
            self.kuang_x = [x_min,x_min,x_min,(x_min+x_max)/2,x_max,x_max,x_max,(x_min+x_max)/2,x_min]
            self.kuang_y = [y_min,(y_min+y_max)/2,y_max,y_max,y_max,(y_min+y_max)/2,y_min,y_min,y_min]
            ax2 = self.fig.add_subplot(1, 2, 2)
            ax2.plot(self.kuang_x,self.kuang_y,color = 'gold',marker = 'o',ms = 10)
            ax2.plot(self.kuang_x,self.kuang_y,'gold',linestyle = '-.')
            ax2.plot((x_min+x_max)/2,(y_min+y_max)/2,color = 'gold',marker = '*',ms = 10)
            self.fig.canvas.draw()

    def reflectxuanqu(self,event):
        global token
        self.x = list(np.array(self.xHE)*self.MSI_shape[0]/self.HE_shape[1])
        self.y = list(np.array(self.yHE) * self.MSI_shape[1] / self.HE_shape[0])
        self.operation.append([self.MSI_shape[0],self.HE_shape[1],self.MSI_shape[1],self.HE_shape[0],2])

        ax = self.fig.add_subplot(1, 2, 1)
        ax.clear()
        if token == 'Original' or token == 'Reflect':
            MSimg = plt.imshow(grid_z0.T, extent=(xmin - 5, xmax + 4, ymax + 4, ymin - 5))
        else:
            MSimg = plt.imshow(grid_z0.T, cmap=cmap2, extent=(xmin - 5, xmax + 4, ymax + 4, ymin - 5))
        self.x.append(self.x[0])
        self.y.append(self.y[0])
        ax.plot(self.x[0], self.y[0], 'r+', ms=20)
        ax.plot(self.x[1:], self.y[1:], 'r*')
        ax.plot(self.x, self.y, 'g', linestyle='-.')
        self.x = self.x[:-1]
        self.y = self.y[:-1]
        self.fig.canvas.draw()

    def reflect(self,x, y, operation):
        m = Shape_Operation()
        x1, y1 = x, y
        for i in range(0, len(operation)):
            if len(operation[i]) == 2:
                if type(operation[i][0]) is str:
                    if operation[i][0] == 'xmin':
                        x1 ,y1= list((np.array(x1) - max(x1)) * operation[i][1] + max(x1)) ,y1
                        print('Scale operation: ', 'Scaling control axis：xmin  Ratio: ', operation[i][1])
                    if operation[i][0] == 'xmax':
                        x1 ,y1= list((np.array(x1)-min(x1))* operation[i][1] +min(x1)) ,y1
                        print('Scale operation: ', 'Scaling control axis：xmax  Ratio: ', operation[i][1])
                    if operation[i][0] == 'ymin':
                        x1 ,y1= x1, list((np.array(y1) - max(y1)) * operation[i][1] + max(y1))
                        print('Scale operation: ', 'Scaling control axis：ymin  Ratio: ', operation[i][1])
                    if operation[i][0] == 'ymax':
                        x1 ,y1= x1, list((np.array(y1)-min(y1))* operation[i][1] +min(y1))
                        print('Scale operation: ', 'Scaling control axis：ymax  Ratio: ', operation[i][1])
                else:
                    x1, y1 = m.lmove(x1, y1, operation[i][0], operation[i][1])
            elif len(operation[i]) == 3:
                if operation[i][-1]==1:
                    x1, y1 = m.lreverse(x1, y1, operation[i][0], operation[i][1])
                else:
                    x1, y1 = m.lzuoyoureverse(x1, y1, operation[i][0], operation[i][1])
            elif len(operation[i]) == 4:
                x1, y1 = m.lrotate(x1, y1, operation[i][0], operation[i][1], operation[i][2])
            elif len(operation[i]) == 5:
                x1 = list(np.array(x1) * operation[i][0] / operation[i][1])
                y1 = list(np.array(y1) * operation[i][2] / operation[i][3])
        return x1, y1

    def do_operation(self,event):
        print('Operation: Reflect Selection')
        self.x ,self.y = self.reflect(self.xHE,self.yHE,self.operation)
        ax = self.fig.add_subplot(1, 2, 1)
        ax.lines.clear()
        self.x.append(self.x[0])
        self.y.append(self.y[0])
        ax.plot(self.x[0], self.y[0], 'r+', ms=20)
        ax.plot(self.x[1:], self.y[1:], 'r*')
        ax.plot(self.x, self.y, 'g', linestyle='-.')
        self.x = self.x[:-1]
        self.y = self.y[:-1]
        self.fig.canvas.draw()

    def rotateshape(self,event):
        global token
        angle = eval(Rotate_box.text)
        self.x, self.y = self.shape_operation.lrotate(self.x, self.y, self.x[0], self.y[0],angle)
        self.operation.append([self.x[0],self.y[0],angle,1])
        ax = self.fig.add_subplot(1, 2, 1)
        ax.clear()
        if token == 'Original' or token == 'Reflect':
            MSimg = plt.imshow(grid_z0.T, extent=(xmin - 5, xmax + 4, ymax + 4, ymin - 5))
        else:
            MSimg = plt.imshow(grid_z0.T, cmap=cmap2, extent=(xmin - 5, xmax + 4, ymax + 4, ymin - 5))
        self.x.append(self.x[0])
        self.y.append(self.y[0])
        ax.plot(self.x[0], self.y[0], 'r+', ms=20)
        ax.plot(self.x[1:], self.y[1:], 'r*')
        ax.plot(self.x, self.y, 'g', linestyle='-.')
        self.x = self.x[:-1]
        self.y = self.y[:-1]
        self.fig.canvas.draw()

    def reflect_progress_update(self,val,stry):
        global info
        if val!=-1:
            self.progressBar.progressBar.setValue(val)
            if val == 95:
                self.progressBar.label.setText('Saving Data...'+info)
            if val==100:
                self.progressBar.label.setText('Finished!')
                self.clo = close_widget_thread(1)
                self.clo.start()
                self.clo.trigger.connect(self.close_progressbar)
        else:
            self.progressBar.label.setText(u'Running Error: '+stry)
            self.progressBar.progressBar.setValue(0)
            self.progressBar.pushButton.setVisible(True)

    def Reflect_export(self,event):
        global HE_export_meta_list, HE_export_meta_names
        HE_export_meta_list = self.cho2.comboBox.SelectIndex()
        HE_export_meta_names = self.cho2.comboBox.Selectlist()
        if len(HE_export_meta_list)==0:
            m = 'No metabolite chosen! Please choose fusion metabolites.'
            self.error(m)
        else:
            file_name = self.cho2.lineEdit.text()
            if file_name.split('.')[-1] == 'xls':
                self.cho2.close()
                self.reflect_export_thread = HE_MSI_export_thread(file_name,self.xHE, self.yHE, self.MSI_points, self.HE_points,self.MSI_intensity, self.HE_intensity,self.operation,HE_export_meta_list,self.MSI_shape)
                self.progressBar = My_Progress_Form()
                self.progressBar.progressBar.setValue(0)
                self.progressBar.pushButton.setVisible(True)
                self.progressBar.pushButton.setText('Cancel')
                self.progressBar.pushButton.clicked.connect(lambda : self.thread_terminate(self.reflect_export_thread))
                self.progressBar.show()
                self.reflect_export_thread.start()
                self.reflect_export_thread.trigger.connect(self.reflect_progress_update)
            else:
                e = 'Please input the file save path!'
                m = 'Running error, info: ' + str(e)
                self.error(m)

    def Reflect_export_choose(self,event):
        global Meta_Names
        p = list(Meta_Names)
        self.cho2 = My_Combo_choose_Form()
        self.cho2.comboBox.Clear_Items()
        self.cho2.comboBox.Add_Items(p)
        self.cho2.pushButton.clicked.connect(self.Reflect_export)
        self.cho2.show()

    def thread_terminate(self,thread_name):
        thread_name.terminate()

    def progress_update(self,val,stry):
        if val!=-1:
            self.progressBar.progressBar.setValue(val)
            if val==100:
                self.progressBar.label.setText('Finished!')
                self.clo = close_widget_thread(1)
                self.clo.start()
                self.clo.trigger.connect(self.close_progressbar)
        else:
            self.progressBar.label.setText('Running Error: '+stry)
            self.progressBar.progressBar.setValue(0)
            self.progressBar.pushButton.setVisible(True)

    def origin_progress_update(self,val,stry):
        if val!=-1:
            self.progressBar.progressBar.setValue(val)
            if val==100:
                self.progressBar.label.setText('Finished!')
                self.clo = close_widget_thread(1)
                self.clo.start()
                self.clo.trigger.connect(self.close_progressbar)
                data = xlrd.open_workbook('Selection metabolites data.xls')
                table = data.sheets()[0]
                global Meta_Names
                x = range(len(Meta_Names))
                origin_data=[]
                reflect_data =[]
                for i in x:
                    origin_data.append(np.array(table.col_values(5+3*i,start_rowx=1)).mean())
                    reflect_data.append(np.array(table.col_values(6+3*i,start_rowx=1)).mean())
                fig = plt.figure()
                plt.rcParams['font.sans-serif'] = ['SimHei']  # 指定默认字体
                plt.rcParams['axes.unicode_minus'] = False
                plt.bar(x, origin_data, width=0.15)
                plt.bar([i + 0.2 for i in x], reflect_data, width=0.15)
                plt.xticks([i + 0.1 for i in x], Meta_Names)
                plt.show()

        else:
            self.progressBar.label.setText('Running Error: '+stry)
            self.progressBar.progressBar.setValue(0)
            self.progressBar.pushButton.setVisible(True)

    def close_progressbar(self,val):
        if val==100:
            self.progressBar.close()

    def Combine_Pic_Generate(self,event):
        global Fusion_meta_list,Fusion_meta_names
        Fusion_meta_list = self.cho_fusion.comboBox.SelectIndex()
        Fusion_meta_names = self.cho_fusion.comboBox.Selectlist()
        if len(Fusion_meta_list) >= 10:
            m = 'To many metabolites chosen! Choose <=9 metabolites.'
            self.error(m)
        elif len(Fusion_meta_list) == 0:
            m = 'No metabolite chosen! Please choose fusion metabolites.'
            self.error(m)
        else:
            file_name = self.cho_fusion.lineEdit.text()
            if file_name.split('.')[-1] == 'npz':
                self.cho_fusion.close()
                global png_files
                #self.cho_fusion.close()
                self.combine_thread = Combine_picture_thread(png_files)
                self.progressBar = My_Progress_Form()
                self.progressBar.progressBar.setValue(0)
                self.progressBar.pushButton.setVisible(True)
                self.progressBar.pushButton.setText('Cancel')
                self.progressBar.pushButton.clicked.connect(lambda:self.thread_terminate(self.combine_thread))
                self.progressBar.show()
                self.combine_thread.start()
                self.combine_thread.trigger.connect(self.progress_update)
                self.combine_thread.trigger2.connect(self.combine_pic_plot)
            else:
                e = 'Please input the file save path!'
                m = 'Running error, info: ' + str(e)
                self.error(m)

    def Combine_Pic_Choose(self,event):
        self.cho_fusion = My_Fusion_Combo_choose_Form()
        self.cho_fusion.pushButton.clicked.connect(self.Combine_Pic_Generate)
        self.cho_fusion.show()

    def combine_pic_plot(self,total_x,total_y,Final_HE_data):
        try:
            global token,Meta_len,Index_button,som_val,Meta,figure2, cmap3, norm, som1, som2, startcolor, midcolor, endcolor, radio, cb
            global Fusion_meta_list, Fusion_meta_names
            self.progressBar.label.setText(u'Preparing data for fusion......')
            self.progressBar.progressBar.setValue(75)
            print(u'Preparing data for fusion......')
            print(Fusion_meta_list)
            data = np.load(self.cho_fusion.lineEdit.text())

            MSI_Temp_data = data['MSI_meta_data']
            HE_data = data['HE_meta_data']

            MSI_data = []
            for i in range(0,len(Fusion_meta_list)):
                Temp = MSI_Temp_data[:,Fusion_meta_list[i]]
                MSI_data.append(Temp)
            MSI_data = np.array(MSI_data).T

            print(HE_data.shape,MSI_data.shape)
            Meta_len = len(Fusion_meta_list)
            Meta = list(Fusion_meta_names)
            self.progressBar.label.setText(u'Start fusion algorithm, pay attention to console!')
            self.progressBar.progressBar.setValue(80)

            if self.cho_fusion.radioButton_2.isChecked() == True:
                try:
                    neuron_num = eval(self.cho_fusion.lineEdit_2.text())
                    iteration_num = eval(self.cho_fusion.lineEdit_3.text())
                    layer_num = self.cho_fusion.comboBox_2.currentIndex()+1
                    print('neuron num: '+ str(neuron_num))
                    print('iteration num: '+str(iteration_num))
                    print('layer num: '+str(layer_num))
                except Exception as e:
                    m = 'Running Error, info: '+e
                    self.error(m)

                self.mlp_thread = MLP_thread(neuron_num,layer_num,iteration_num,HE_data,MSI_data,Final_HE_data)
                self.mlp_thread.start()
                self.mlp_thread.trigger.connect(lambda data_l:self.fusion_plot(total_x,total_y,data_l))
            else:
                try:
                    n_component = eval(self.cho_fusion.lineEdit_2.text())
                    print('n_component: '+str(n_component))
                except Exception as e:
                    m = 'Running Error, info: '+e
                    self.error(m)
                pls2 = PLSRegression(n_components=HE_data.shape[1])
                pls2.fit(HE_data, MSI_data)
                predict_l_temp = pls2.predict(HE_data)
                score_l = []
                for i in range(0, MSI_data.shape[1]):
                    u = np.square(predict_l_temp[:, i] - MSI_data[:, i]).sum()
                    v = np.square(MSI_data[:, i] - MSI_data[:, i].mean()).sum()
                    score_i = 1 - u / v
                    print('Fusion Score:' + str(i + 1) + '  ' + str(round(80 + 20 * score_i, 4)))
                    score_l.append(str(round(80 + 20 * score_i, 4)))
                Predict_Temp = pls2.predict(Final_HE_data)
                self.fusion_plot(total_x,total_y,[score_l,Predict_Temp])
        except Exception as e:
            m = 'Running Error, info: ' + str(e)
            self.error(m)

    def fusion_plot(self,total_x,total_y,data_l):
        try:
            global token, Meta_len, Index_button, som_val, Meta, figure2, cmap3, norm, som1, som2, startcolor, midcolor, endcolor, radio, cb
            global Fusion_meta_list, Fusion_meta_names
            global Predict_Image,score_l,token
            score_l = data_l[0]
            Predict_Temp = data_l[1]
            self.progressBar.label.setText(u'Finished!')
            self.progressBar.progressBar.setValue(100)
            self.clo = close_widget_thread(1)
            self.clo.start()
            self.clo.trigger.connect(self.close_progressbar)
            print(u'Generating Fusion Image.....')
            Predict_Image = []
            for num in range(0, Meta_len):
                index = -1
                Predict_Image.append([])
                for i in range(0, total_x):
                    Temp = []
                    for j in range(0, total_y):
                        index += 1
                        if Meta_len == 1:
                            Temp.append(Predict_Temp[index])
                        else:
                            Temp.append(Predict_Temp[index][num])
                    Temp = np.array(Temp)
                    if i == 0:
                        Predict_Image[num] = Temp
                    else:
                        Predict_Image[num] = np.vstack((Predict_Image[num], Temp))

                if token == 'Subtract':
                    if abs(Predict_Image[num].min()) > Predict_Image[num].max():
                        Predict_Image[num][0, 0] = abs(Predict_Image[num].min())
                    else:
                        Predict_Image[num][0, 0] = -Predict_Image[num].max()
        except Exception as e:
            m = 'Running Error, info: '+str(e)
            self.error(m)

        figure2 = plt.figure()
        plt.subplots_adjust(bottom=0.2, left=0.25, hspace=0.2)  # 调整子图间距
        plt.rcParams['font.sans-serif'] = ['KaiTi']  # 指定默认字体
        plt.rcParams['axes.unicode_minus'] = False  # 解决保存图像是负号'-'显示为方块的问题

        if token == 'Original' or token == 'Reflect':
            cmap3 = 'viridis'
        else:
            cmap3 = col.LinearSegmentedColormap.from_list('own2',
                                                          ['#ffffff', '#00ebff', '#00bfff', '#007cd9', '#000000', '#e95b5b',
                                                           '#fc8b8b', '#feabab', '#ffffff'])
        cb, p, ax_l = [], [], []
        if Meta_len == 1:
            sub_inde = [1, 1]
        elif Meta_len == 2:
            sub_inde = [1, 2]
        elif Meta_len == 3 or Meta_len == 4:
            sub_inde = [2, 2]
        elif Meta_len == 5 or Meta_len == 6:
            sub_inde = [2, 3]
        elif Meta_len == 7 or Meta_len == 8 or Meta_len == 9:
            sub_inde = [3, 3]

        try:
            for k in range(0, Meta_len):
                ax_l.append(figure2.add_subplot(sub_inde[0], sub_inde[1], k + 1))
                p.append(ax_l[k].imshow(Predict_Image[k], cmap=cmap3))
                cb.append(plt.colorbar(p[k], ax=ax_l[k]))
        except Exception as e:
            m = 'Running Error, info: ' + str(e)
            self.error(m)

        figure2.canvas.mpl_connect('button_press_event', lambda event: self.fig2_on_click(p, event))

        self.cmp_change_i = 1
        om1 = plt.axes([0.25, 0.1, 0.65, 0.03])  # 第一slider的位置
        om2 = plt.axes([0.25, 0.05, 0.65, 0.03])  # 第一slider的位置
        som1 = Slider(om1, u'min:', 0, 100, valstep=1, valinit=35)  # 产生第一slider
        som2 = Slider(om2, u'max:', 0, 100, valstep=1, valinit=65)  # 产生第一slider
        som1.on_changed(lambda ev: update(self.cmp_change_i, 's1'))
        som2.on_changed(lambda ev: update(self.cmp_change_i, 's2'))

        som_val = {'min_1': 35, 'max_1': 65, 'min_2': 35, 'max_2': 65, 'min_3': 35, 'max_3': 65, 'min_4': 35,
                   'max_4': 65, 'min_5': 35, 'max_5': 65, 'min_6': 35, 'max_6': 65, 'min_7': 35, 'max_7': 65,
                   'min_8': 35, 'max_8': 65, 'min_9': 35, 'max_9': 65}
        for inde in range(0, Meta_len):
            update(inde + 1, 's1')

        cc = plt.axes([0.025, 0.5, 0.2, 0.15])
        radio = RadioButtons(cc, ('re', 'ue', 'een'), active=0)
        radio.on_clicked(cmp_change)

        if token == 'Original' or token == 'Reflect':
            path_1, path_2, path_3 = 'colormap/viridis.png', 'colormap/hot.png', 'colormap/RtoB.png'
        else:
            path_1, path_2, path_3 = 'colormap/RtoB.png', 'colormap/YtoB.png', 'colormap/GtoR.png'
        RemoveLearning_button_position = plt.axes([0.07, 0.590, 0.15, 0.05])
        RemoveLearning_button = MyButton(RemoveLearning_button_position,
                                         image=np.array(PIL.Image.open(path_1)),
                                         image_pressed=np.array(PIL.Image.open(path_1)))
        LRemoveLearning_button_position = plt.axes([0.07, 0.550, 0.15, 0.05])
        LRemoveLearning_button = MyButton(LRemoveLearning_button_position,
                                          image=np.array(PIL.Image.open(path_2)),
                                          image_pressed=np.array(PIL.Image.open(path_2)))
        GRemoveLearning_button_position = plt.axes([0.07, 0.513, 0.15, 0.05])
        GRemoveLearning_button = MyButton(GRemoveLearning_button_position,
                                          image=np.array(PIL.Image.open(path_3)),
                                          image_pressed=np.array(PIL.Image.open(path_3)))

        plt.show()

    def fig2_on_click(self,p,event):
        for i in range(0,len(p)):
            if event.button == 1 and event.inaxes == p[i].axes:
                try:
                    global Meta, som_val, som1, som2
                    number = i+1
                    self.cmp_change_i = number
                    self.msg = My_Message_Form()
                    self.msg.setWindowFlags(QtCore.Qt.FramelessWindowHint)
                    self.msg.label.setText(
                        u'Start modifying the colormap of  ' + str(number) + ') ' + str(Meta[number - 1]))
                    self.msg.show()
                    som1.set_val(som_val['min_' + str(number)])
                    som2.set_val(som_val['max_' + str(number)])
                    self.clo_i = close_widget_thread(1)
                    self.clo_i.start()
                    self.clo_i.trigger.connect(self.close_window)
                except Exception as e:
                    m = 'Running Error, info: ' + str(e)
                    self.error(m)
                break

    def origin_export(self,event):
        self.origin_export_thread = Origin_export_thread(self.x, self.y, self.MSI_points,self.xls_path)
        self.progressBar = My_Progress_Form()
        self.progressBar.progressBar.setValue(0)
        self.progressBar.pushButton.setVisible(True)
        self.progressBar.pushButton.setText('Cancel')
        self.progressBar.pushButton.clicked.connect(lambda :self.thread_terminate(self.origin_export_thread))
        self.progressBar.show()
        self.origin_export_thread.start()
        self.origin_export_thread.trigger.connect(self.origin_progress_update)

class s_Form(object):
    def setupUi(self, Form):
        Form.setObjectName("Form")
        Form.resize(576, 386)
        icon = QtGui.QIcon()
        icon.addPixmap(QtGui.QPixmap(":/newPrefix/Mydata/panda.ico"), QtGui.QIcon.Normal, QtGui.QIcon.Off)
        Form.setWindowIcon(icon)
        self.verticalLayout_3 = QtWidgets.QVBoxLayout(Form)
        self.verticalLayout_3.setObjectName("verticalLayout_3")
        self.verticalLayout_2 = QtWidgets.QVBoxLayout()
        self.verticalLayout_2.setObjectName("verticalLayout_2")
        spacerItem = QtWidgets.QSpacerItem(20, 40, QtWidgets.QSizePolicy.Minimum, QtWidgets.QSizePolicy.Expanding)
        self.verticalLayout_2.addItem(spacerItem)
        self.horizontalLayout = QtWidgets.QHBoxLayout()
        self.horizontalLayout.setObjectName("horizontalLayout")
        self.label = QtWidgets.QLabel(Form)
        font = QtGui.QFont()
        font.setFamily("Arial")
        font.setBold(True)
        font.setPointSize(12)
        self.label.setFont(font)
        self.label.setObjectName("label")
        self.horizontalLayout.addWidget(self.label)
        self.lineEdit = QtWidgets.QLineEdit(Form)
        sizePolicy = QtWidgets.QSizePolicy(QtWidgets.QSizePolicy.Expanding, QtWidgets.QSizePolicy.Maximum)
        sizePolicy.setHorizontalStretch(0)
        sizePolicy.setVerticalStretch(0)
        sizePolicy.setHeightForWidth(self.lineEdit.sizePolicy().hasHeightForWidth())
        self.lineEdit.setSizePolicy(sizePolicy)
        self.lineEdit.setMaximumSize(QtCore.QSize(16777215, 40))
        font = QtGui.QFont()
        font.setFamily("Arial")
        font.setBold(True)
        font.setPointSize(12)
        self.lineEdit.setFont(font)
        self.lineEdit.setObjectName("lineEdit")
        self.horizontalLayout.addWidget(self.lineEdit)
        self.pushButton_2 = QtWidgets.QPushButton(Form)
        sizePolicy = QtWidgets.QSizePolicy(QtWidgets.QSizePolicy.Minimum, QtWidgets.QSizePolicy.Minimum)
        sizePolicy.setHorizontalStretch(0)
        sizePolicy.setVerticalStretch(0)
        sizePolicy.setHeightForWidth(self.pushButton_2.sizePolicy().hasHeightForWidth())
        self.pushButton_2.setSizePolicy(sizePolicy)
        self.pushButton_2.setMinimumSize(QtCore.QSize(150, 40))
        self.pushButton_2.setMaximumSize(QtCore.QSize(120, 40))
        font = QtGui.QFont()
        font.setFamily("Arial")
        font.setBold(True)
        font.setPointSize(12)
        self.pushButton_2.setFont(font)
        icon1 = QtGui.QIcon()
        icon1.addPixmap(QtGui.QPixmap(":/newPrefix/Button_Image/excel.png"), QtGui.QIcon.Normal, QtGui.QIcon.Off)
        self.pushButton_2.setIcon(icon1)
        self.pushButton_2.setIconSize(QtCore.QSize(35, 35))
        self.pushButton_2.setObjectName("pushButton_2")
        self.horizontalLayout.addWidget(self.pushButton_2)
        self.verticalLayout_2.addLayout(self.horizontalLayout)
        self.horizontalLayout_3 = QtWidgets.QHBoxLayout()
        self.horizontalLayout_3.setContentsMargins(80, -1, 80, -1)
        self.horizontalLayout_3.setSpacing(15)
        self.horizontalLayout_3.setObjectName("horizontalLayout_3")
        self.label_3 = QtWidgets.QLabel(Form)
        sizePolicy = QtWidgets.QSizePolicy(QtWidgets.QSizePolicy.Preferred, QtWidgets.QSizePolicy.Minimum)
        sizePolicy.setHorizontalStretch(0)
        sizePolicy.setVerticalStretch(0)
        sizePolicy.setHeightForWidth(self.label_3.sizePolicy().hasHeightForWidth())
        self.label_3.setSizePolicy(sizePolicy)
        self.label_3.setMaximumSize(QtCore.QSize(16777215, 40))
        font = QtGui.QFont()
        font.setFamily("Arial")
        font.setBold(True)
        font.setPointSize(12)
        self.label_3.setFont(font)
        self.label_3.setObjectName("label_3")
        self.horizontalLayout_3.addWidget(self.label_3)
        self.comboBox = QtWidgets.QComboBox(Form)
        sizePolicy = QtWidgets.QSizePolicy(QtWidgets.QSizePolicy.Expanding, QtWidgets.QSizePolicy.Preferred)
        sizePolicy.setHorizontalStretch(0)
        sizePolicy.setVerticalStretch(0)
        sizePolicy.setHeightForWidth(self.comboBox.sizePolicy().hasHeightForWidth())
        self.comboBox.setSizePolicy(sizePolicy)
        self.comboBox.setMaximumSize(QtCore.QSize(16777215, 90))
        font = QtGui.QFont()
        font.setFamily("Arial")
        font.setBold(True)
        font.setPointSize(12)
        self.comboBox.setFont(font)
        self.comboBox.setObjectName("comboBox")
        self.horizontalLayout_3.addWidget(self.comboBox)
        self.horizontalLayout_3.setStretch(0, 1)
        self.horizontalLayout_3.setStretch(1, 2)
        self.verticalLayout_2.addLayout(self.horizontalLayout_3)
        self.horizontalLayout_2 = QtWidgets.QHBoxLayout()
        self.horizontalLayout_2.setObjectName("horizontalLayout_2")
        self.label_2 = QtWidgets.QLabel(Form)
        font = QtGui.QFont()
        font.setFamily("Arial")
        font.setBold(True)
        font.setPointSize(12)
        self.label_2.setFont(font)
        self.label_2.setObjectName("label_2")
        self.horizontalLayout_2.addWidget(self.label_2)
        self.lineEdit_2 = QtWidgets.QLineEdit(Form)
        sizePolicy = QtWidgets.QSizePolicy(QtWidgets.QSizePolicy.Expanding, QtWidgets.QSizePolicy.Maximum)
        sizePolicy.setHorizontalStretch(0)
        sizePolicy.setVerticalStretch(0)
        sizePolicy.setHeightForWidth(self.lineEdit_2.sizePolicy().hasHeightForWidth())
        self.lineEdit_2.setSizePolicy(sizePolicy)
        self.lineEdit_2.setMaximumSize(QtCore.QSize(16777215, 40))
        font = QtGui.QFont()
        font.setFamily("Arial")
        font.setBold(True)
        font.setPointSize(12)
        self.lineEdit_2.setFont(font)
        self.lineEdit_2.setObjectName("lineEdit_2")
        self.horizontalLayout_2.addWidget(self.lineEdit_2)
        self.pushButton_3 = QtWidgets.QPushButton(Form)
        sizePolicy = QtWidgets.QSizePolicy(QtWidgets.QSizePolicy.Minimum, QtWidgets.QSizePolicy.Minimum)
        sizePolicy.setHorizontalStretch(0)
        sizePolicy.setVerticalStretch(0)
        sizePolicy.setHeightForWidth(self.pushButton_3.sizePolicy().hasHeightForWidth())
        self.pushButton_3.setSizePolicy(sizePolicy)
        self.pushButton_3.setMinimumSize(QtCore.QSize(160, 40))
        self.pushButton_3.setMaximumSize(QtCore.QSize(120, 40))
        font = QtGui.QFont()
        font.setFamily("Arial")
        font.setBold(True)
        font.setPointSize(12)
        self.pushButton_3.setFont(font)
        icon2 = QtGui.QIcon()
        icon2.addPixmap(QtGui.QPixmap(":/newPrefix/Button_Image/filebrowser.png"), QtGui.QIcon.Normal, QtGui.QIcon.Off)
        self.pushButton_3.setIcon(icon2)
        self.pushButton_3.setIconSize(QtCore.QSize(35, 35))
        self.pushButton_3.setObjectName("pushButton_3")
        self.horizontalLayout_2.addWidget(self.pushButton_3)
        self.verticalLayout_2.addLayout(self.horizontalLayout_2)
        spacerItem1 = QtWidgets.QSpacerItem(20, 40, QtWidgets.QSizePolicy.Minimum, QtWidgets.QSizePolicy.Expanding)
        self.verticalLayout_2.addItem(spacerItem1)
        self.horizontalLayout_5 = QtWidgets.QHBoxLayout()
        self.horizontalLayout_5.setObjectName("horizontalLayout_5")
        spacerItem2 = QtWidgets.QSpacerItem(40, 20, QtWidgets.QSizePolicy.Expanding, QtWidgets.QSizePolicy.Minimum)
        self.horizontalLayout_5.addItem(spacerItem2)
        self.verticalLayout = QtWidgets.QVBoxLayout()
        self.verticalLayout.setObjectName("verticalLayout")
        self.radioButton_2 = QtWidgets.QRadioButton(Form)
        font = QtGui.QFont()
        font.setFamily("Arial")
        font.setPointSize(12)
        font.setBold(True)
        font.setWeight(75)
        self.radioButton_2.setFont(font)
        self.radioButton_2.setIconSize(QtCore.QSize(16, 16))
        self.radioButton_2.setObjectName("radioButton_2")
        self.verticalLayout.addWidget(self.radioButton_2)
        self.radioButton = QtWidgets.QRadioButton(Form)
        font = QtGui.QFont()
        font.setFamily("Arial")
        font.setPointSize(12)
        font.setBold(True)
        font.setWeight(75)
        self.radioButton.setFont(font)
        self.radioButton.setIconSize(QtCore.QSize(16, 16))
        self.radioButton.setObjectName("radioButton")
        self.verticalLayout.addWidget(self.radioButton)
        self.radioButton_3 = QtWidgets.QRadioButton(Form)
        font = QtGui.QFont()
        font.setFamily("Arial")
        font.setPointSize(12)
        font.setBold(True)
        font.setWeight(75)
        self.radioButton_3.setFont(font)
        self.radioButton_3.setIconSize(QtCore.QSize(16, 16))
        self.radioButton_3.setObjectName("radioButton_3")
        self.verticalLayout.addWidget(self.radioButton_3)
        self.horizontalLayout_5.addLayout(self.verticalLayout)
        spacerItem3 = QtWidgets.QSpacerItem(40, 20, QtWidgets.QSizePolicy.Expanding, QtWidgets.QSizePolicy.Minimum)
        self.horizontalLayout_5.addItem(spacerItem3)
        self.verticalLayout_2.addLayout(self.horizontalLayout_5)
        spacerItem4 = QtWidgets.QSpacerItem(513, 13, QtWidgets.QSizePolicy.Minimum, QtWidgets.QSizePolicy.Expanding)
        self.verticalLayout_2.addItem(spacerItem4)
        spacerItem5 = QtWidgets.QSpacerItem(20, 40, QtWidgets.QSizePolicy.Minimum, QtWidgets.QSizePolicy.Expanding)
        self.verticalLayout_2.addItem(spacerItem5)
        self.horizontalLayout_4 = QtWidgets.QHBoxLayout()
        self.horizontalLayout_4.setContentsMargins(-1, 0, -1, -1)
        self.horizontalLayout_4.setObjectName("horizontalLayout_4")
        spacerItem6 = QtWidgets.QSpacerItem(40, 20, QtWidgets.QSizePolicy.Expanding, QtWidgets.QSizePolicy.Minimum)
        self.horizontalLayout_4.addItem(spacerItem6)
        self.pushButton = QtWidgets.QPushButton(Form)
        sizePolicy = QtWidgets.QSizePolicy(QtWidgets.QSizePolicy.Minimum, QtWidgets.QSizePolicy.Minimum)
        sizePolicy.setHorizontalStretch(0)
        sizePolicy.setVerticalStretch(0)
        sizePolicy.setHeightForWidth(self.pushButton.sizePolicy().hasHeightForWidth())
        self.pushButton.setSizePolicy(sizePolicy)
        self.pushButton.setMinimumSize(QtCore.QSize(120, 50))
        font = QtGui.QFont()
        font.setFamily("Arial")
        font.setBold(True)
        font.setPointSize(12)
        self.pushButton.setFont(font)
        self.pushButton.setObjectName("pushButton")
        self.horizontalLayout_4.addWidget(self.pushButton)
        spacerItem7 = QtWidgets.QSpacerItem(40, 20, QtWidgets.QSizePolicy.Expanding, QtWidgets.QSizePolicy.Minimum)
        self.horizontalLayout_4.addItem(spacerItem7)
        self.horizontalLayout_4.setStretch(0, 2)
        self.horizontalLayout_4.setStretch(1, 3)
        self.horizontalLayout_4.setStretch(2, 2)
        self.verticalLayout_2.addLayout(self.horizontalLayout_4)
        self.verticalLayout_3.addLayout(self.verticalLayout_2)

        self.retranslateUi(Form)
        self.pushButton.clicked.connect(Form.imzml_plot)
        self.pushButton_2.clicked.connect(Form.file_open)
        self.pushButton_3.clicked.connect(Form.file_browser_open)
        QtCore.QMetaObject.connectSlotsByName(Form)

    def retranslateUi(self, Form):
        _translate = QtCore.QCoreApplication.translate
        Form.setWindowTitle(_translate("Form", "MSI Fusion"))
        self.label.setText(_translate("Form", "Image Data(.xls) Path: "))
        self.pushButton_2.setText(_translate("Form", "Choose xls "))
        self.label_3.setText(_translate("Form", "Target Metabolite:"))
        self.label_2.setText(_translate("Form", "HE Data Folder Path: "))
        self.pushButton_3.setText(_translate("Form", "Choose Folder"))
        self.radioButton_2.setText(_translate("Form", "Reflect Selection Fusion"))
        self.radioButton.setText(_translate("Form", "Original Selection Fusion"))
        self.radioButton_3.setText(_translate("Form", "Subtract Data Fusion"))
        self.pushButton.setText(_translate("Form", "Start Drawing"))

class My_Form(QtWidgets.QWidget,s_Form):
    def __init__(self):
        super().__init__()
        self.setupUi(self)
        self.radioButton.setChecked(False)
        self.radioButton_2.setChecked(False)
        self.radioButton_3.setChecked(True)

    def close_progressbar(self,val):
        if val==100:
            self.progressBar.close()

    def error(self,m):
        self.eW=My_Error_Form()
        self.eW.label.setText(m)
        self.eW.show()

    def plot(self,biaoji,png_files, xmin,xmax,ymin,ymax,MSI_points, MSI_intensity, grid_z0, MSI_shape):
        try:
            global fig1, MSimg, HEimg, cmap2, HE_points, HE_intensity,cmap2
            global HE_shape, reelect_button, Operation_button, reverse_button, Reflect_export_button, Pic_Generate_button
            global Rotate_button, Rotate_box,Origin_export_button
            global Remove_Learning_button,Import_Learning_button,Export_Learning_button
            fig1 = plt.figure()
            #plt.rcParams['font.sans-serif'] = ['SimHei']  # 指定默认字体
            plt.rcParams['axes.unicode_minus'] = False
            plt.subplot(121)

            cmap2 = col.LinearSegmentedColormap.from_list('own2',['#ffffff','#00ebff','#00bfff','#007cd9','#000000','#e95b5b','#fc8b8b','#feabab','#ffffff'])

            if biaoji == 'Original' or biaoji == 'Reflect':
                MSimg = plt.imshow(grid_z0.T,extent = (xmin-5,xmax+4,ymax+4,ymin-5))
            else:
                MSimg = plt.imshow(grid_z0.T,cmap = cmap2,extent = (xmin-5,xmax+4,ymax+4,ymin-5))

            plt.colorbar()
            plt.subplots_adjust(left=0.05, bottom=0.2, wspace=0.1)

            plt.subplot(122)
            HEImg = np.array(PIL.Image.open(png_files[0]))
            HE_shape = np.shape(HEImg)
            HE_points, HE_intensity = [], []
            for i in range(0, HE_shape[1]):
                for j in range(0, HE_shape[0]):
                    HE_points.append([i, j])
                    HE_intensity.append(HEImg[j, i])
            HE_points, HE_intensity = np.array(HE_points), np.array(HE_intensity)
            HEimg = plt.imshow(HEImg)
            plt.subplots_adjust(right=0.95, bottom=0.2, wspace=0.1)

            callback = ButtonHandler(fig1, grid_z0, HEImg, MSI_points, HE_points, MSI_intensity, HE_intensity,
                                     MSI_shape, HE_shape,self.xls_path)
            callback.connect()

            Remove_Learning_button_position = plt.axes([0.02, 0.15, 0.1, 0.08])
            Remove_Learning_button = MyButton(Remove_Learning_button_position, '',
                                              np.array(PIL.Image.open(u'Button_Image_e/Remove_Learning.png')),
                                              np.array(PIL.Image.open('Button_Image_e/Remove_Learning_pressed.png')))
            Remove_Learning_button.on_clicked(callback.RemoveLearning)

            Import_Learning_button_position = plt.axes([0.12, 0.15, 0.1, 0.08])
            Import_Learning_button = MyButton(Import_Learning_button_position, '',
                                              np.array(PIL.Image.open(u'Button_Image_e/Import_Learning.png')),
                                              np.array(PIL.Image.open('Button_Image_e/Import_Learning_pressed.png')))
            Import_Learning_button.on_clicked(callback.Operation_import)

            Export_Learning_button_position = plt.axes([0.22, 0.15, 0.1, 0.08])
            Export_Learning_button = MyButton(Export_Learning_button_position, '',
                                              np.array(PIL.Image.open(u'Button_Image_e/Export_Learning.png')),
                                              np.array(PIL.Image.open('Button_Image_e/Export_Learning_pressed.png')))
            Export_Learning_button.on_clicked(callback.Operation_export)

            reelect_button_position = plt.axes([0.02, 0.05, 0.1, 0.08])
            reelect_button = MyButton(reelect_button_position, '',
                                      np.array(PIL.Image.open('Button_Image_e/Reselect.png')),
                                      np.array(PIL.Image.open('Button_Image_e/Reselect_pressed.png')))
            reelect_button.on_clicked(callback.reelect)

            Operation_button_position = plt.axes([0.12, 0.05, 0.1, 0.08])
            Operation_button = MyButton(Operation_button_position, '',
                                        np.array(PIL.Image.open('Button_Image_e/Reflect_generate.png')),
                                        np.array(PIL.Image.open('Button_Image_e/Reflect_generate_pressed.png')))
            Operation_button.on_clicked(callback.do_operation)

            Rotate_box_position = plt.axes([0.35, 0.05, 0.1, 0.08])
            Rotate_box = TextBox(Rotate_box_position, 'Angle:', label_pad=0.2)

            Rotate_button_position = plt.axes([0.45, 0.05, 0.1, 0.08])
            Rotate_button = MyButton(Rotate_button_position, '',
                                     np.array(PIL.Image.open('Button_Image_e/Rotation_Transform.png')),
                                     np.array(PIL.Image.open('Button_Image_e/Rotation_Transform_pressed.png')))
            Rotate_button.on_clicked(callback.rotateshape)

            reverse_button_position = plt.axes([0.58, 0.05, 0.1, 0.08])
            reverse_button = MyButton(reverse_button_position, '',
                                      np.array(PIL.Image.open('Button_Image_e/HE映射.png')),
                                      np.array(PIL.Image.open('Button_Image_e/HE映射_pressed.png')))
            reverse_button.on_clicked(callback.reflectxuanqu)

            Reflect_export_button_position = plt.axes([0.68, 0.05, 0.1, 0.08])
            Reflect_export_button = MyButton(Reflect_export_button_position, '',
                                             np.array(PIL.Image.open('Button_Image_e/Selection_export.png')),
                                             np.array(PIL.Image.open('Button_Image_e/Selection_export_pressed.png')))
            Reflect_export_button.on_clicked(callback.Reflect_export_choose)

            Pic_Generate_button_position = plt.axes([0.78, 0.05, 0.1, 0.08])
            Pic_Generate_button = MyButton(Pic_Generate_button_position, '',
                                           np.array(PIL.Image.open('Button_Image_e/Fusion_generate.png')),
                                           np.array(PIL.Image.open('Button_Image_e/Fusion_generate_pressed.png')))
            Pic_Generate_button.on_clicked(callback.Combine_Pic_Choose)

            Origin_export_button_position = plt.axes([0.88, 0.05, 0.1, 0.08])
            Origin_export_button = MyButton(Origin_export_button_position, '',
                                            np.array(PIL.Image.open('Button_Image_e/Meta_intensity_comparison.png')),
                                            np.array(PIL.Image.open('Button_Image_e/Meta_intensity_comparison_pressed.png')))
            Origin_export_button.on_clicked(callback.origin_export)
            plt.show()
        except Exception as e :
            m='Running error, error message:'+str(e)
            self.error(m)

    def imzml_plot(self):
        try:
            global token,png_files, xmin, xmax, ymin, ymax, MSI_points, MSI_intensity, grid_z0, MSI_shape,index
            if self.radioButton.isChecked() == True:
                token = 'Original'
            elif self.radioButton_2.isChecked() == True:
                token = 'Reflect'
            else:
                token = 'Subtract'
            MSI_path = self.lineEdit.text()
            png_files = pngfile_path(self.lineEdit_2.text())
            self.msg = My_Message_Form()
            self.msg.setWindowFlags(QtCore.Qt.FramelessWindowHint)
            self.msg.label.setText('Generating Image, please wait....')
            self.msg.show()
            self.clo = close_widget_thread(3)
            self.clo.start()
            self.clo.trigger.connect(self.close_draw_window)
            index = self.comboBox.currentIndex()
            xmin, xmax, ymin, ymax, MSI_points, MSI_intensity, grid_z0, MSI_shape = MSI_xls_data(token,self.lineEdit.text(),index)
            self.plot(token,png_files, xmin,xmax,ymin,ymax,MSI_points, MSI_intensity, grid_z0, MSI_shape)
        except Exception as e :
            m='Running Error, info: '+str(e)
            self.error(m)

    def file_open(self):
        try:
            global Meta_Names,file_name
            path = os.getcwd()
            self.comboBox.clear()
            file_name,_= QtWidgets.QFileDialog.getOpenFileName(self,u'Choose File',path,'Excel files (*.xls *.xlsx)')
            self.lineEdit.setText(file_name)
            if file_name.split('.')[-1].lower() == 'xls':
                self.xls_path = file_name
            if file_name != '':
                data = xlrd.open_workbook(file_name)
                table = data.sheets()[0]
                nrows = table.nrows
                items = table.col_values(0, 1)
                Meta_Names = []
                for k in items:
                    if k != '': Meta_Names.append(k)
                self.comboBox.addItems(Meta_Names)
        except Exception as e :
            m='Running Error, info: '+str(e)
            self.error(m)

    def file_browser_open(self):
        path = os.getcwd()
        file = QtWidgets.QFileDialog.getExistingDirectory(self, "Choose Folder", path)
        self.lineEdit_2.setText(file)

    def progress_update(self,val,stry):
        if val!=-1:
            self.progressBar.progressBar.setValue(val)
            if val==100:
                self.progressBar.label.setText('Finished!')
                self.clo = close_widget_thread(1)
                self.clo.start()
                self.clo.trigger.connect(self.close_progressbar)
        else:
            self.progressBar.label.setText('Running Error, info: '+stry)
            self.progressBar.progressBar.setValue(0)
            self.progressBar.pushButton.setVisible(True)

    def close_window(self,val):
        if val==100:
            self.progressBar.close()

    def close_draw_window(self,val):
        if val==100:
            self.msg.close()
            fig1.show()

class HE_MSI_export_thread(QtCore.QThread):
    trigger = QtCore.pyqtSignal(int,str)
    def __init__(self,file_name,x, y, MSI_points, HE_points, MSI_intensity,HE_intensity,operation,HE_export_meta,MSI_shape):
        super().__init__()
        self.file_name = file_name
        self.x = x
        self.y = y
        self.MSI_shape = MSI_shape
        self.MSI_points = MSI_points
        self.HE_points = HE_points
        self.MSI_intensity = MSI_intensity
        self.HE_intensity = HE_intensity
        self.operation = operation
        self.HE_export_meta_list = HE_export_meta

    def reflect(self,x, y, operation):
        m = Shape_Operation()
        x1, y1 = x, y
        for i in range(0, len(operation)):
            if len(operation[i]) == 2:
                if type(operation[i][0]) is str:
                    if operation[i][0] == 'xmin':
                        x1 ,y1= list((np.array(x1) - max(x1)) * operation[i][1] + max(x1)) ,y1
                        print('Scale operation: ', 'Scaling control axis：xmin  Ratio: ', operation[i][1])
                    if operation[i][0] == 'xmax':
                        x1 ,y1= list((np.array(x1)-min(x1))* operation[i][1] +min(x1)) ,y1
                        print('Scale operation: ', 'Scaling control axis：xmax  Ratio: ', operation[i][1])
                    if operation[i][0] == 'ymin':
                        x1 ,y1= x1, list((np.array(y1) - max(y1)) * operation[i][1] + max(y1))
                        print('Scale operation: ', 'Scaling control axis：ymin  Ratio: ', operation[i][1])
                    if operation[i][0] == 'ymax':
                        x1 ,y1= x1, list((np.array(y1)-min(y1))* operation[i][1] +min(y1))
                        print('Scale operation: ', 'Scaling control axis：ymax  Ratio: ', operation[i][1])
                else:
                    x1, y1 = m.lmove(x1, y1, operation[i][0], operation[i][1])
            elif len(operation[i]) == 3:
                if operation[i][-1]==1:
                    x1, y1 = m.lreverse(x1, y1, operation[i][0], operation[i][1])
                else:
                    x1, y1 = m.lzuoyoureverse(x1, y1, operation[i][0], operation[i][1])
            elif len(operation[i]) == 4:
                x1, y1 = m.lrotate(x1, y1, operation[i][0], operation[i][1], operation[i][2])
            elif len(operation[i]) == 5:
                x1 = list(np.array(x1) * operation[i][0] / operation[i][1])
                y1 = list(np.array(y1) * operation[i][2] / operation[i][3])
        return x1, y1

    def run(self):
        try:
            global token,file_name,MSI_points,xmin,xmax,ymin,ymax,info,HE_export_meta_names
            print(u'Start exporting selected data, please wait.....')
            #print(self.HE_export_meta_list)
            self.trigger.emit(5,'')
            loc = np.array([[0, 0]])
            for i in range(0, len(self.x)):
                loc = np.append(loc, [[self.x[i], self.y[i]]], axis=0)
            polygon = Polygon(loc[1:], True)
            iden = polygon.contains_points(self.HE_points)
            export_points_x, export_points_y, export_points_intensity = [], [], []

            for i in range(0, len(self.HE_points)):
                if iden[i] == True:
                    export_points_x.append(self.HE_points[i, 0])
                    export_points_y.append(self.HE_points[i, 1])
                    export_points_intensity.append(self.HE_intensity[i])
            export_reflect_x, export_reflect_y = self.reflect(export_points_x, export_points_y, self.operation)

            xx_min = MSI_points[0,0]
            export_reflect_intensity ,meta_i_intensity= [],[]
            key,bb= 5,0
            cc = len(self.HE_export_meta_list)
            for meta_i in self.HE_export_meta_list:
                bb+=1
                key1 = int((bb / cc) * 15)
                if key1 - key >= 1:
                    self.trigger.emit(key1, '')
                    key = key1
                meta_i_total = MSI_xls_data(token,file_name,meta_i)
                meta_i_intensity.append(meta_i_total[5])
                export_reflect_intensity.append([])
            for i in range(0, len(export_points_y)):
                key1 = int((i / len(export_points_y)) * 100)
                if key1-key>=1 and key1<=90:
                    self.trigger.emit(key1, '')
                    key = key1
                flag_point = int(export_reflect_x[i]-xx_min-1) * (self.MSI_shape[1])
                if flag_point<0: flag_point =0
                flag = False
                while not flag:
                    if ((MSI_points[flag_point, 0] - export_reflect_x[i]) ** 2 <= 0.3) and (
                            (MSI_points[flag_point, 1] - export_reflect_y[i]) ** 2 <= 0.3):
                        export_reflect_x[i], export_reflect_y[i] = MSI_points[flag_point, 0], MSI_points[flag_point, 1]
                        for n in range(0,len(meta_i_intensity)):
                            export_reflect_intensity[n].append(meta_i_intensity[n][flag_point])
                        flag = True
                    if flag_point>=len(MSI_points)-1:
                        print(export_reflect_x[i],export_reflect_y[i])
                        flag = True
                    flag_point += 1

            HE_data, rows = [], []
            for i in range(0, len(png_files)):
                Img = PIL.Image.open(png_files[i])
                Img_data = np.array(Img)
                if len(np.shape(Img_data)) == 3:
                    Img_L = np.array(Img.convert('L'))
                    HE_data.append([Img_data, Img_L])
                    rows.append(np.shape(Img_data)[2] + 1)
                else:
                    HE_data.append([Img_data])
                    rows.append(1)
            row0 = ['Metabolite names', 'HE x', 'HE y', 'MSI x', 'MSI y']
            for i in HE_export_meta_names:
                row0.append(str(i) + u'intensity')
            for i in range(0, len(rows)):
                word = 'HE intensity ' + str(i + 1)
                if rows[i] == 1:
                    row0.append(word)
                else:
                    for kk in range(0, rows[i]):
                        row0.append(word + ' f' + str(kk + 1))

            self.trigger.emit(95,'')
            Temp_name_total = []
            for i in range(0,len(HE_export_meta_names)):
                Temp_name_total.append(HE_export_meta_names[i])
            Temp_name_total = np.array(Temp_name_total)
            info = u'Total '+str(len(export_reflect_x))+u' points'
            print(info)

            length = len(export_reflect_x)
            MSI_Temp_total,HE_Temp_total = [],[]
            for i in range(0, length):
                MSI_Temp_l = []
                HE_Temp_l = []
                for z in range(0,len(export_reflect_intensity)):
                    MSI_Temp_l.append(float(export_reflect_intensity[z][i]))
                for j in range(0, len(rows)):
                    if len(HE_data[j]) == 2:
                        num_j = np.shape(HE_data[j][0])[2]
                        for kk in range(0,num_j):
                            HE_Temp_l.append(float(HE_data[j][0][export_points_y[i], export_points_x[i], kk]))
                        HE_Temp_l.append(float(HE_data[j][1][export_points_y[i], export_points_x[i]]))
                    else:
                        HE_Temp_l.append(float(HE_data[j][0][export_points_y[i], export_points_x[i]]))
                MSI_Temp_total.append(np.array(MSI_Temp_l))
                HE_Temp_total.append(np.array(HE_Temp_l))
            MSI_Temp_total = np.array(MSI_Temp_total)
            HE_Temp_total = np.array(HE_Temp_total)
            print(MSI_Temp_total.shape)
            print(HE_Temp_total.shape)
            fs = self.file_name.split('.')[0]
            np.savez(fs, MSI_meta_data = MSI_Temp_total,HE_meta_data = HE_Temp_total)
            np.save(fs+'_meta_names',Temp_name_total)
            print(u'MSI and HE corresponding data export successfully!')
            self.trigger.emit(100,'')
        except Exception as e:
            m = 'Running Error,info: ' + str(e)
            self.trigger.emit(-1,m)

class Origin_export_thread(QtCore.QThread):
    trigger = QtCore.pyqtSignal(int,str)
    def __init__(self,x, y,MSI_points,xlspath):
        super().__init__()
        self.x = x
        self.y = y
        self.MSI_points = MSI_points
        self.xlspath = xlspath

    def run(self):
        try:
            global subtract_points
            print(u'Start exporting selected Metabolites Data, Please wait.....')
            loc = np.array([[0, 0]])
            for i in range(0, len(self.x)):
                loc = np.append(loc, [[self.x[i], self.y[i]]], axis=0)
            polygon = Polygon(loc[1:], True)
            iden = polygon.contains_points(subtract_points)
            export_points = []
            for i in range(0, len(subtract_points)):
                if iden[i] == True:
                    if not [subtract_points[i, 0],subtract_points[i, 1]] in export_points:
                        export_points.append([subtract_points[i, 0],subtract_points[i, 1]])
            export_points = np.array(export_points)
            MSi_rows = np.shape(export_points)[0]
            self.trigger.emit(5, '')
            index_l = list_points_confirm(export_points,subtract_points)
            self.trigger.emit(20, '')
            data = xlrd.open_workbook(self.xlspath)
            table = data.sheets()[0]
            f = xlwt.Workbook()
            sheet1 = f.add_sheet(u'sheet1', cell_overwrite_ok=True)
            row0 = table.row_values(0)
            for i in range(0, len(row0)):
                sheet1.write(0, i, row0[i])
            key = 0
            for i in range(0, len(index_l)):
                key1 = int((i / len(index_l)) * 100)
                if key1 - key >= 1 and key1 <= 98 and key1>=20:
                    self.trigger.emit(key1, '')
                    key = key1
                Tmp = table.row_values(index_l[i])
                for j in range(0,len(Tmp)):
                    if type(Tmp[j])==str:
                        sheet1.write(i + 1, j, Tmp[j])
                    else:
                        sheet1.write(i + 1, j, float(Tmp[j]))
            f.save(u'Selection metabolites data.xls')
            print(u'MSI metabolites data exported successfully!')
            self.trigger.emit(100,'')
        except Exception as e:
            m = 'Running Error, info: ' + str(e)
            self.trigger.emit(-1,m)

class close_widget_thread(QtCore.QThread):
    trigger = QtCore.pyqtSignal(int)

    def __init__(self,seconds):
        super().__init__()
        self.second = seconds

    def run(self):
        self.sleep(self.second)
        self.trigger.emit(100)

class Combine_picture_thread(QtCore.QThread):
    trigger = QtCore.pyqtSignal(int, str)
    trigger2 = QtCore.pyqtSignal(int,int,np.ndarray)
    def __init__(self,png_files):
        super().__init__()
        self.png_files = png_files

    def run(self):
        try:
            New_HE_data = []
            for i in range(0, len(self.png_files)):
                Img = PIL.Image.open(self.png_files[i])
                Img_data = np.array(Img)
                if len(np.shape(Img_data)) == 3:
                    if i == 0:
                        total_x, total_y = np.shape(Img_data)[0], np.shape(Img_data)[1]
                    Img_L = np.array(Img.convert('L'))
                    New_HE_data.append([Img_data, Img_L])
                else:
                    New_HE_data.append([Img_data])
            key = 0
            HE_mat = []
            for x in range(0, total_x):
                key1 = int((x / (total_x - 1)) * 100)
                if key1-key>=1 and key1<=70:
                    self.trigger.emit(key1,'')
                    key = key1
                for y in range(0, total_y):
                    Temp = []
                    for j in range(0, len(New_HE_data)):
                        if len(New_HE_data[j]) == 2:
                            num_j = np.shape(New_HE_data[j][0])[2]
                            for kk in range(0,num_j):
                                Temp.append(float(New_HE_data[j][0][x, y, kk]))
                            Temp.append(float(New_HE_data[j][1][x, y]))
                        else:
                            Temp.append(float(New_HE_data[j][0][x, y]))
                    Temp = np.array(Temp)
                    HE_mat.append(Temp)
            Final_HE_data = np.mat(HE_mat)
            self.trigger2.emit(total_x,total_y,Final_HE_data)
        except Exception as e:
            m = 'Running Error, info: ' + str(e)
            self.trigger.emit(-1, m)

class Imzml_draw_thread(QtCore.QThread):
    trigger = QtCore.pyqtSignal(int,str)
    trigger2 = QtCore.pyqtSignal(str,list,int,int,int,int,np.ndarray,np.ndarray,np.ndarray,tuple)

    def __init__(self,png,a,index):
        super().__init__()
        self.png = png
        self.inputpath = a
        self.index =index

    def run(self):
        try:
            global xmin, xmax, ymin, ymax, MSI_points, MSI_intensity, grid_z0, MSI_shape
            global Meta_Names_M_start, Meta_Names_M_end
            left = Meta_Names_M_start[self.index]
            right = Meta_Names_M_end[self.index]
            p = ImzMLParser(self.inputpath)
            Coor = p.coordinates
            total = len(Coor)

            m = p.getspectrum(0)
            points = np.array([list(Coor[0][:-1])])
            Tempinte = PeakIntensitySum(m, left, right)
            intensity = np.array([Tempinte])
            self.trigger.emit(0, '')
            for indecount in range(1, total):
                self.trigger.emit(int((indecount / (total - 1)) * 100), '')
                m = p.getspectrum(indecount)
                points = np.append(points, [list(Coor[indecount][:-1])], axis=0)
                intensity = np.append(intensity, [PeakIntensitySum(m, left, right)], axis=0)

            xmax = np.array(Coor)[:, 0].max()
            ymax = np.array(Coor)[:, 1].max()
            xmin = np.array(Coor)[:, 0].min()
            ymin = np.array(Coor)[:, 1].min()

            grid_x, grid_y = np.mgrid[1:xmax:(xmax * 10j), 1:ymax:(ymax * 10j)]

            grid_z0 = griddata(points, intensity, (grid_x, grid_y), method='linear')
            MSI_points, MSI_intensity = [], []
            for i in range(0, xmax * 10):
                for j in range(0, ymax * 10):
                    MSI_points.append([i, j])
                    MSI_intensity.append(grid_z0[i, j])
            MSI_points, MSI_intensity = np.array(MSI_points), np.array(MSI_intensity)
            MSI_shape = np.shape(grid_z0)
            self.trigger2.emit('imzml',self.png, xmin,xmax,ymin,ymax, MSI_points, MSI_intensity, grid_z0, MSI_shape)
        except Exception as e:
            m = 'Running Error, info: ' + str(e)
            self.trigger.emit(-1, m)

class MLP_thread(QtCore.QThread):
    trigger = QtCore.pyqtSignal(list)

    def __init__(self,neuron,layer,iter_num,he_data,msi_data,final_he_data):
        super().__init__()
        self.neuron = neuron
        self.layer = layer
        self.iter_num  =iter_num
        self.HE_data = he_data
        self.MSI_data = msi_data
        self.final_HE_data = final_he_data

    def run(self):
        a = []
        for i in range(0,self.layer): a.append(self.neuron)
        a = tuple(a)
        mlp = make_pipeline(StandardScaler(),
                            MLPRegressor(hidden_layer_sizes=a,
                                         tol=1e-2, max_iter=self.iter_num, verbose=True, random_state=5))

        mlp.fit(self.HE_data, self.MSI_data)
        predict_l_temp = mlp.predict(self.HE_data)
        print('fgfgf')
        score_l = []
        mm = len(predict_l_temp.shape)
        for i in range(0, self.MSI_data.shape[1]):
            if mm>1:
                u = np.square(predict_l_temp[:, i] - self.MSI_data[:, i]).sum()
            else:
                u = np.square(predict_l_temp[:] - self.MSI_data[:, i]).sum()
            v = np.square(self.MSI_data[:, i] - self.MSI_data[:, i].mean()).sum()
            score_i = 1 - u / v
            print('Fusion Score:' + str(i + 1) + '  ' + str(round(80 + 20 * score_i, 4)))
            score_l.append(str(round(80 + 20 * score_i, 4)))

        Predict_Temp = mlp.predict(self.final_HE_data)
        self.trigger.emit([score_l,Predict_Temp])
        print('successful!')

if __name__ == "__main__" :
    QtCore.QCoreApplication.setAttribute(QtCore.Qt.AA_EnableHighDpiScaling)
    app = QtWidgets.QApplication(sys.argv)
    my_pyqt_form = My_Form()
    my_pyqt_form.show()
    sys.exit(app.exec_())
