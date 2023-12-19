from PyQt5 import QtCore, QtGui, QtWidgets
from PyQt5.QtWidgets import QComboBox,QLineEdit,QListWidget,QCheckBox,QListWidgetItem
from pyimzml.ImzMLParser import ImzMLParser
from progressbar import *
import numpy as np
import os
import matplotlib.colors as col
from scipy.interpolate import griddata
from matplotlib import pyplot as plt
import xlwt,xlrd
from matplotlib.widgets import Button,TextBox,Slider,RadioButtons
from matplotlib.patches import Polygon
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
        self.horizontalLayout = QtWidgets.QHBoxLayout(Form)
        self.setWindowFlags(QtCore.Qt.FramelessWindowHint)
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
class choose_Form(object):
    def setupUi(self, Form):
        Form.setObjectName("Form")
        Form.resize(540, 341)
        icon = QtGui.QIcon()
        icon.addPixmap(QtGui.QPixmap(":/newPrefix/ui/icon.png"), QtGui.QIcon.Normal, QtGui.QIcon.Off)
        Form.setWindowIcon(icon)
        self.horizontalLayout_3 = QtWidgets.QHBoxLayout(Form)
        self.horizontalLayout_3.setObjectName("horizontalLayout_3")
        self.verticalLayout = QtWidgets.QVBoxLayout()
        self.verticalLayout.setContentsMargins(10, 40, 10, 30)
        self.verticalLayout.setObjectName("verticalLayout")
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
        sizePolicy = QtWidgets.QSizePolicy(QtWidgets.QSizePolicy.Minimum, QtWidgets.QSizePolicy.Maximum)
        sizePolicy.setHorizontalStretch(0)
        sizePolicy.setVerticalStretch(0)
        sizePolicy.setHeightForWidth(self.pushButton_3.sizePolicy().hasHeightForWidth())
        self.pushButton_3.setSizePolicy(sizePolicy)
        self.pushButton_3.setMinimumSize(QtCore.QSize(140, 0))
        self.pushButton_3.setMaximumSize(QtCore.QSize(120, 40))
        font = QtGui.QFont()
        font.setFamily("Arial")
        font.setBold(True)
        font.setPointSize(12)
        self.pushButton_3.setFont(font)
        icon = QtGui.QIcon()
        icon.addPixmap(QtGui.QPixmap(":/newPrefix/Button_Image/excel.png"), QtGui.QIcon.Normal, QtGui.QIcon.Off)
        self.pushButton_3.setIcon(icon)
        self.pushButton_3.setIconSize(QtCore.QSize(35, 35))
        self.pushButton_3.setObjectName("pushButton_3")
        self.horizontalLayout_2.addWidget(self.pushButton_3)
        self.verticalLayout.addLayout(self.horizontalLayout_2)
        spacerItem = QtWidgets.QSpacerItem(20, 40, QtWidgets.QSizePolicy.Minimum, QtWidgets.QSizePolicy.Expanding)
        self.verticalLayout.addItem(spacerItem)
        self.horizontalLayout = QtWidgets.QHBoxLayout()
        self.horizontalLayout.setSpacing(10)
        self.horizontalLayout.setObjectName("horizontalLayout")
        self.label = QtWidgets.QLabel(Form)
        font = QtGui.QFont()
        font.setFamily("Arial")
        font.setBold(True)
        font.setPointSize(12)
        self.label.setFont(font)
        self.label.setObjectName("label")
        self.horizontalLayout.addWidget(self.label)
        self.comboBox = QtWidgets.QComboBox(Form)
        sizePolicy = QtWidgets.QSizePolicy(QtWidgets.QSizePolicy.Preferred, QtWidgets.QSizePolicy.Minimum)
        sizePolicy.setHorizontalStretch(0)
        sizePolicy.setVerticalStretch(0)
        sizePolicy.setHeightForWidth(self.comboBox.sizePolicy().hasHeightForWidth())
        self.comboBox.setSizePolicy(sizePolicy)
        self.comboBox.setMinimumSize(QtCore.QSize(0, 30))
        font = QtGui.QFont()
        font.setFamily("Arial")
        font.setBold(True)
        font.setPointSize(12)
        self.comboBox.setFont(font)
        self.comboBox.setObjectName("comboBox")
        self.horizontalLayout.addWidget(self.comboBox)
        self.horizontalLayout.setStretch(0, 1)
        self.horizontalLayout.setStretch(1, 2)
        self.verticalLayout.addLayout(self.horizontalLayout)
        spacerItem1 = QtWidgets.QSpacerItem(20, 40, QtWidgets.QSizePolicy.Minimum, QtWidgets.QSizePolicy.Expanding)
        self.verticalLayout.addItem(spacerItem1)
        self.pushButton = QtWidgets.QPushButton(Form)
        sizePolicy = QtWidgets.QSizePolicy(QtWidgets.QSizePolicy.Minimum, QtWidgets.QSizePolicy.Minimum)
        sizePolicy.setHorizontalStretch(0)
        sizePolicy.setVerticalStretch(0)
        sizePolicy.setHeightForWidth(self.pushButton.sizePolicy().hasHeightForWidth())
        self.pushButton.setSizePolicy(sizePolicy)
        self.pushButton.setMinimumSize(QtCore.QSize(0, 45))
        font = QtGui.QFont()
        font.setFamily("Arial")
        font.setBold(True)
        font.setPointSize(12)
        self.pushButton.setFont(font)
        self.pushButton.setObjectName("pushButton")
        self.verticalLayout.addWidget(self.pushButton)
        self.horizontalLayout_3.addLayout(self.verticalLayout)

        self.retranslateUi(Form)
        QtCore.QMetaObject.connectSlotsByName(Form)

    def retranslateUi(self, Form):
        _translate = QtCore.QCoreApplication.translate
        Form.setWindowTitle(_translate("Form", "Start Drawing"))
        self.label_2.setText(_translate("Form", "Image File Path:"))
        self.pushButton_3.setText(_translate("Form", "Choose xls"))
        self.label.setText(_translate("Form", "Target Metabolites:"))
        self.pushButton.setText(_translate("Form", "Start Drawing"))
class My_choose_Form(QtWidgets.QWidget,choose_Form):
    def __init__(self):
        super().__init__()
        self.setupUi(self)
class Combo_choose_Form(object):
    def setupUi(self, Form):
        Form.setObjectName("Form")
        Form.resize(470, 306)
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
        font.setBold(True)
        font.setPointSize(12)
        self.label_2.setFont(font)
        self.label_2.setObjectName("label_2")
        self.horizontalLayout_2.addWidget(self.label_2)
        self.lineEdit = QtWidgets.QLineEdit(Form)
        font = QtGui.QFont()
        font.setFamily("Arial")
        font.setBold(True)
        font.setPointSize(12)
        self.lineEdit.setFont(font)
        self.lineEdit.setObjectName("lineEdit")
        self.horizontalLayout_2.addWidget(self.lineEdit)
        self.pushButton_2 = QtWidgets.QPushButton(Form)
        font = QtGui.QFont()
        font.setFamily("Arial")
        font.setBold(True)
        font.setPointSize(12)
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
        font.setBold(True)
        font.setPointSize(12)
        self.label.setFont(font)
        self.label.setObjectName("label")
        self.horizontalLayout.addWidget(self.label)
        self.comboBox = ComboCheckBox([''])
        self.comboBox.setObjectName("comboBox")
        self.horizontalLayout.addWidget(self.comboBox)
        self.horizontalLayout.setStretch(0, 1)
        self.horizontalLayout.setStretch(1, 2)
        self.verticalLayout_2.addLayout(self.horizontalLayout)
        spacerItem1 = QtWidgets.QSpacerItem(20, 40, QtWidgets.QSizePolicy.Minimum, QtWidgets.QSizePolicy.Expanding)
        self.verticalLayout_2.addItem(spacerItem1)
        self.verticalLayout = QtWidgets.QVBoxLayout()
        self.verticalLayout.setContentsMargins(70, -1, -1, -1)
        self.verticalLayout.setObjectName("verticalLayout")
        self.radioButton = QtWidgets.QRadioButton(Form)
        font = QtGui.QFont()
        font.setFamily("Arial")
        font.setBold(True)
        font.setPointSize(12)
        self.radioButton.setFont(font)
        self.radioButton.setObjectName("radioButton")
        self.verticalLayout.addWidget(self.radioButton)
        self.radioButton_2 = QtWidgets.QRadioButton(Form)
        font = QtGui.QFont()
        font.setFamily("Arial")
        font.setBold(True)
        font.setPointSize(12)
        self.radioButton_2.setFont(font)
        self.radioButton_2.setObjectName("radioButton_2")
        self.verticalLayout.addWidget(self.radioButton_2)
        self.verticalLayout_2.addLayout(self.verticalLayout)
        spacerItem2 = QtWidgets.QSpacerItem(20, 40, QtWidgets.QSizePolicy.Minimum, QtWidgets.QSizePolicy.Expanding)
        self.verticalLayout_2.addItem(spacerItem2)
        self.pushButton = QtWidgets.QPushButton(Form)
        font = QtGui.QFont()
        font.setFamily("Arial")
        font.setBold(True)
        font.setPointSize(12)
        self.pushButton.setFont(font)
        self.pushButton.setObjectName("pushButton")
        self.verticalLayout_2.addWidget(self.pushButton)
        self.horizontalLayout_3.addLayout(self.verticalLayout_2)

        self.retranslateUi(Form)
        QtCore.QMetaObject.connectSlotsByName(Form)

    def retranslateUi(self, Form):
        _translate = QtCore.QCoreApplication.translate
        Form.setWindowTitle(_translate("Form", "Data Export"))
        self.label_2.setText(_translate("Form", "Data File Save Path:"))
        self.pushButton_2.setText(_translate("Form", "Choose"))
        self.label.setText(_translate("Form", "Target Metabolites:"))
        self.radioButton.setText(_translate("Form", "Differential imaging data export"))
        self.radioButton_2.setText(_translate("Form", "Ratio imaging data export"))
        self.pushButton.setText(_translate("Form", "Start Exporting"))
class My_Combo_choose_Form(QtWidgets.QWidget,Combo_choose_Form):
    def __init__(self):
        super().__init__()
        self.setupUi(self)
        self.radioButton.setChecked(True)
        self.radioButton_2.setChecked(False)
        self.pushButton_2.clicked.connect(self.btn_save)

    def btn_save(self):
        try:
            path = os.getcwd()
            file_name, _ = QtWidgets.QFileDialog.getSaveFileName(self, u'Save file', path,'Excel files (*.xls)')
            self.lineEdit.setText(file_name)
        except Exception as e:
            m = 'Running error, info: ' + str(e)
            self.error(m)

def imzml_intensity(x,y,startmz,endmz):
    global points
    Coor = np.array(points)
    x_in = np.where(Coor[:,0]==x)[0]
    target_index= x_in[np.where(Coor[x_in][:,1]==y)[0]][0]
    m = p.getspectrum(target_index)
    intensity = PeakIntensitySum(m, startmz, endmz)
    return intensity

def update(val):
    global norm,cb,grid_z1
    maxval ,minval = float(grid_z1.max()),float(grid_z1.min())
    s1 = som1.val
    s2 = som2.val
    vmin = minval+(maxval-minval)*s1/100
    vmax = maxval - (maxval - minval) * (100-s2) / 100
    norm = col.Normalize(vmin = vmin,vmax = vmax)
    ax = figure2.add_subplot(1,1,1)
    cb.remove()
    ax.clear()
    plt.imshow(grid_z1.T, extent=(1, N_xlen, N_ylen, 1), cmap=cmap2,norm = norm)
    cb = plt.colorbar()
    figure2.canvas.draw()

def update2(val):
    global norm,cb1,grid_z0
    maxval ,minval = float(grid_z0.max()),float(grid_z0.min())
    s1 = DDsom1.val
    s2 = DDsom2.val
    vmin = minval+(maxval-minval)*s1/100
    vmax = maxval - (maxval - minval) * (100-s2) / 100
    norm = col.Normalize(vmin = vmin,vmax = vmax)
    ax = figure.add_subplot(1,1,1)
    cb1.remove()
    ax.clear()
    plt.imshow(grid_z0.T, extent=(1, xlen, ylen, 1),norm = norm)
    cb1 = plt.colorbar()
    figure.canvas.draw()

def cmp_change(label):
    global cmap2,cb,grid_z1
    if label == 're':
        cmap2 = col.LinearSegmentedColormap.from_list('own2', ['#ffffff','#00ebff','#00bfff','#007cd9','#000000','#e95b5b','#fc8b8b','#feabab','#ffffff'])
    elif label == 'ue':
        cmap2 = 'hot'
        #cmap2 = col.LinearSegmentedColormap.from_list('own2', [startcolor[1], midcolor[1], endcolor[1]])
    elif label == 'een':
        cmap2 = col.LinearSegmentedColormap.from_list('own2', ['#000000','#520068','#1500ff','#00e1ff','#00ff50','#77ff00','#fcff00','#ffc300','#ff0000'])
    ax = figure2.add_subplot(1, 1, 1)
    ax.clear()
    cb.remove()
    plt.imshow(grid_z1.T, extent=(1, N_xlen, N_ylen, 1), cmap=cmap2, norm=norm)
    cb = plt.colorbar()
    figure2.canvas.draw()

def points_confirm(x,y,point_data):
    c=np.where(point_data[:,0]==x)[0]
    m=np.where(point_data[c,1]==y)[0]
    if len(m)==0:
        return -1
    else:
        return c[m[0]]

class MyButton(Button):
    def __init__(self, ax, label='', image=None, image_pressed = None,color='0.85', hovercolor='0.95'):
        super(MyButton, self).__init__(ax, label, image, color, hovercolor)
        self.ax ,self.image,self.image_pressed = ax , image,image_pressed
        self.connect_event('axes_enter_event',self.change1)
        self.connect_event('axes_leave_event', self.change2)

    def change1(self,event):
        if self.image is not None:
            if event.inaxes == self.ax:
                self.ax.imshow(self.image_pressed)
    def change2(self,event):
        if self.image_pressed is not None:
            if event.inaxes == self.ax:
                self.ax.imshow(self.image)

def PeakIntensitySum(spec,diffl,diffr):
    s=spec[0][np.where(spec[0]<=diffr)]
    ss=spec[1][np.where(s>=diffl)]
    return(ss.sum())

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

class ButtonHandler():
    def __init__(self,fig,points,intensity):
        self.fig = fig
        self.points = points
        self.intensity = intensity
        self.loc=np.array([[0,0]])
        self.x,self.y = [] , []
        self.x1,self.y1 =[] , []
        self.operation = []
        self.shape_operation = Shape_Operation()
        self.cidtotalmotion = None
        self.cidpointmotion = None
        self.pointid = -1
        self.flag = False

    def connect(self):
        self.cidclick= self.fig.canvas.mpl_connect('button_press_event', self.on_click)
        self.cidpress = self.fig.canvas.mpl_connect('key_press_event', self.on_press)
        self.cidrelease = self.fig.canvas.mpl_connect('key_release_event', self.on_release)

    def error(self,m):
        self.eW=My_Error_Form()
        self.eW.label.setText(m)
        self.eW.show()

    def on_click(self,event):
        if event.button == 1 and event.inaxes==img.axes and event.key is None:
            if event.xdata != None and event.ydata != None :
                x = float(event.xdata)
                y = float(event.ydata)
                if x>1 and y>1:
                    self.x.append(x)
                    self.y.append(y)
                    self.loc=np.append(self.loc,[[x,y]],axis=0)

                global ax
                ax = figure.add_subplot(1,1,1)
                ax.lines.clear()
                self.x.append(self.x[0])
                self.y.append(self.y[0])
                ax.plot(self.x[0],self.y[0], 'r+', ms = 20)
                ax.plot(self.x[1:], self.y[1:], 'r*')
                ax.plot(self.x,self.y,'w',linestyle='-.')
                self.x = self.x[:-1]
                self.y = self.y[:-1]
                figure.canvas.draw()
        if event.button == 3 and event.inaxes==img.axes and event.key is None:
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
                ax = figure.add_subplot(1, 1, 1)
                ax.lines.clear()
                self.x.append(self.x[0])
                self.y.append(self.y[0])
                ax.plot(self.x[0], self.y[0], 'r+', ms=20)
                ax.plot(self.x[1:], self.y[1:], 'r*')
                ax.plot(self.x, self.y, 'w', linestyle='-.')
                self.x = self.x[:-1]
                self.y = self.y[:-1]
                figure.canvas.draw()

    def on_press(self,event):
        if event.key == 'shift':
            self.cidtotalmotion = self.fig.canvas.mpl_connect("button_press_event", self.on_total_move)
        if event.key =='control':
            self.cidpointmotion = self.fig.canvas.mpl_connect('motion_notify_event',self.on_point_move)
        if event.key in ['up','down','left','right','alt+up','alt+down','alt+left','alt+right']:
            if event.key == 'up':
                xmove ,ymove= 0,-2
            elif event.key == 'down':
                xmove, ymove = 0, 2
            elif event.key == 'left':
                xmove, ymove = -2, 0
            elif event.key == 'right':
                xmove, ymove = 2, 0
            elif event.key == 'alt+up':
                xmove, ymove = 0, -0.2
            elif event.key == 'alt+down':
                xmove, ymove = 0, 0.2
            elif event.key == 'alt+left':
                xmove, ymove = -0.2, 0
            elif event.key == 'alt+right':
                xmove, ymove = 0.2, 0
            self.x, self.y = self.shape_operation.lmove(self.x, self.y, xmove, ymove)
            if self.flag == True:
                self.operation.append([xmove, ymove])
            ax = figure.add_subplot(1, 1, 1)
            ax.clear()
            ax.imshow(grid_z0.T, extent=(1, xlen, ylen, 1),norm=norm)
            self.x.append(self.x[0])
            self.y.append(self.y[0])
            ax.plot(self.x[0], self.y[0], 'r+', ms=20)
            ax.plot(self.x[1:], self.y[1:], 'r*')
            ax.plot(self.x, self.y, 'w', linestyle='-.')
            self.x = self.x[:-1]
            self.y = self.y[:-1]
            figure.canvas.draw()

    def on_release(self,event):
        pass

    def on_point_move(self,event):
        if event.button ==1 :
            if self.pointid == -1:
                for i in range(0,len(self.x)):
                    distx = event.xdata - self.x[i]
                    disty = event.ydata - self.y[i]
                    if distx**2 < 30 and disty**2<30 :
                        self.pointid = i
                        break
            else:
                self.x[self.pointid] , self.y[self.pointid] = event.xdata , event.ydata
            ax = figure.add_subplot(1, 1, 1)
            ax.lines.clear()
            self.x.append(self.x[0])
            self.y.append(self.y[0])
            ax.plot(self.x[0], self.y[0], 'r+', ms=20)
            ax.plot(self.x[1:], self.y[1:], 'r*')
            ax.plot(self.x, self.y, 'w', linestyle='-.')
            self.x = self.x[:-1]
            self.y = self.y[:-1]
            figure.canvas.draw()

    def on_total_move(self,event):
        if event.button == 1 and event.inaxes==img.axes:
            xmove = event.xdata - self.x[0]
            ymove = event.ydata - self.y[0]
            self.x, self.y = self.shape_operation.lmove(self.x, self.y, xmove, ymove)
            if self.flag == True:
                self.operation.append([xmove, ymove])
            ax = figure.add_subplot(1, 1, 1)
            ax.clear()
            ax.imshow(grid_z0.T, extent=(1, xlen, ylen, 1),norm=norm)
            self.x.append(self.x[0])
            self.y.append(self.y[0])
            ax.plot(self.x[0], self.y[0], 'r+', ms=20)
            ax.plot(self.x[1:], self.y[1:], 'r*')
            ax.plot(self.x, self.y, 'w', linestyle='-.')
            self.x = self.x[:-1]
            self.y = self.y[:-1]
            figure.canvas.draw()

    def reelect(self,event):
        ax = figure.add_subplot(1, 1, 1)
        ax.clear()
        plt.imshow(grid_z0.T, extent=(1,xlen, ylen,1),norm=norm)
        self.x = []
        self.y = []
        self.loc = np.array([[0, 0]])

    def RemoveLearning(self,event):
        ax = figure.add_subplot(1, 1, 1)
        ax.clear()
        plt.imshow(grid_z0.T, extent=(1, xlen, ylen, 1),norm=norm)
        self.flag = False
        self.x = []
        self.y = []
        self.loc = np.array([[0, 0]])
        self.operation = []
        LearningStatus_box.set_val('Off')
        self.msg = My_Message_Form()
        self.msg.label.setText('Learning Removed!')
        self.msg.show()
        self.clo = close_widget_thread(1)
        self.clo.start()
        self.clo.trigger.connect(self.close_window)

    def StartLearning(self,event):
        try:
            self.flag = True
            LearningStatus_box.set_val('On')
            self.msg = My_Message_Form()
            self.msg.label.setText('Start Learning!')
            self.msg.show()
            self.clo = close_widget_thread(1)
            self.clo.start()
            self.clo.trigger.connect(self.close_window)
        except Exception as e:
            print(str(e))

    def close_window(self,val):
        if val==100:
            self.msg.close()

    def close_draw_window(self,val):
        if val==100:
            self.msg.close()
            figure2.show()

    def EndLearning(self,event):
        self.flag = False
        LearningStatus_box.set_val('Off')
        self.msg = My_Message_Form()
        self.msg.label.setText('Stop Learning')
        self.msg.show()
        self.clo = close_widget_thread(1)
        self.clo.start()
        self.clo.trigger.connect(self.close_window)

    def zuoyou_reverseshape(self,event):
        self.x , self.y = self.shape_operation.lzuoyoureverse(self.x,self.y,self.x[0],self.y[0])
        if self.flag==True:
            self.operation.append([self.x[0],self.y[0],2])
        ax = figure.add_subplot(1, 1, 1)
        ax.clear()
        ax.imshow(grid_z0.T, extent=(1, xlen, ylen, 1),norm=norm)
        self.x.append(self.x[0])
        self.y.append(self.y[0])
        ax.plot(self.x[0], self.y[0], 'r+', ms=20)
        ax.plot(self.x[1:], self.y[1:], 'r*')
        ax.plot(self.x, self.y, 'w', linestyle='-.')
        self.x = self.x[:-1]
        self.y = self.y[:-1]
        figure.canvas.draw()

    def reverseshape(self,event):
        self.x , self.y = self.shape_operation.lreverse(self.x,self.y,self.x[0],self.y[0])
        if self.flag==True:
            self.operation.append([self.x[0],self.y[0],1])
        ax = figure.add_subplot(1, 1, 1)
        ax.clear()
        ax.imshow(grid_z0.T, extent=(1, xlen, ylen, 1),norm=norm)
        self.x.append(self.x[0])
        self.y.append(self.y[0])
        ax.plot(self.x[0], self.y[0], 'r+', ms=20)
        ax.plot(self.x[1:], self.y[1:], 'r*')
        ax.plot(self.x, self.y, 'w', linestyle='-.')
        self.x = self.x[:-1]
        self.y = self.y[:-1]
        figure.canvas.draw()

    def rotateshape(self,event):
        angle = eval(Rotate_box.text)
        self.x, self.y = self.shape_operation.lrotate(self.x, self.y, self.x[0], self.y[0],angle)
        if self.flag == True:
            self.operation.append([self.x[0],self.y[0],angle,1])
        ax = figure.add_subplot(1, 1, 1)
        ax.clear()
        ax.imshow(grid_z0.T, extent=(1, xlen, ylen, 1),norm=norm)
        self.x.append(self.x[0])
        self.y.append(self.y[0])
        ax.plot(self.x[0], self.y[0], 'r+', ms=20)
        ax.plot(self.x[1:], self.y[1:], 'r*')
        ax.plot(self.x, self.y, 'w', linestyle='-.')
        self.x = self.x[:-1]
        self.y = self.y[:-1]
        figure.canvas.draw()

    def reflect(self,x, y, operation):
        m = Shape_Operation()
        x1, y1 = x, y
        for i in range(0, len(operation)):
            if len(operation[i]) == 2:
                x1, y1 = m.lmove(x1, y1, operation[i][0], operation[i][1])
            elif len(operation[i]) == 3:
                if operation[i][-1]==1:
                    x1, y1 = m.lreverse(x1, y1, operation[i][0], operation[i][1])
                else:
                    x1, y1 = m.lzuoyoureverse(x1, y1, operation[i][0], operation[i][1])
            elif len(operation[i]) == 4:
                x1, y1 = m.lrotate(x1, y1, operation[i][0], operation[i][1], operation[i][2])
        return x1, y1

    def do_operation(self,event):
        print('Reflect Operation!')
        self.x1 ,self.y1 = self.reflect(self.x,self.y,self.operation)
        ax = figure.add_subplot(1, 1, 1)
        ax.clear()
        plt.imshow(grid_z0.T, extent=(1, xlen, ylen, 1), norm = norm)
        self.x.append(self.x[0])
        self.y.append(self.y[0])
        ax.plot(self.x[0], self.y[0], 'r+', ms=20)
        ax.plot(self.x[1:], self.y[1:], 'r*')
        ax.plot(self.x, self.y, 'w', linestyle='-.')
        self.x = self.x[:-1]
        self.y = self.y[:-1]
        self.x1.append(self.x1[0])
        self.y1.append(self.y1[0])
        ax.plot(self.x1[0], self.y1[0], 'r+', ms=20)
        ax.plot(self.x1[1:], self.y1[1:], 'r*')
        ax.plot(self.x1, self.y1, 'w', linestyle='-.')
        self.x1 = self.x1[:-1]
        self.y1 = self.y1[:-1]
        figure.canvas.draw()

    def close_progressbar(self,val):
        if val==100:
            self.progressBar.close()

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

    def Operation_import(self,event):
        try:
            data = xlrd.open_workbook('Learning Recording Data.xls')
            table = data.sheets()[0]
            nrows = table.nrows
            temp = []
            for i in range(0,nrows):
                m=table.row_values(i)
                rtemp = []
                for j in m:
                    if type(j) == float :
                        rtemp.append(j)
                temp.append(rtemp)
            self.operation = temp
            print('Learning Recording Data Imported Successfully!')
            self.msg = My_Message_Form()
            self.msg.label.setText('Learning Recording Data Imported Successfully!')
            self.msg.show()
            self.clo = close_widget_thread(1)
            self.clo.start()
            self.clo.trigger.connect(self.close_window)
        except Exception as e :
            m='Running error, info: '+str(e)
            self.error(m)

    def Operation_export(self,event):
        f = xlwt.Workbook()
        sheet1 = f.add_sheet(u'sheet1', cell_overwrite_ok=True)
        if self.operation is not None:
            for i in range(0,len(self.operation)):
                for j in range(0,len(self.operation[i])):
                    sheet1.write(i,j,self.operation[i][j])
        f.save('Learning Recording Data.xls')
        print('Learning Recording Data Exported Successfully!')
        self.msg = My_Message_Form()
        self.msg.label.setText('Learning Recording Data Exported Successfully!')
        self.msg.show()
        self.clo = close_widget_thread(1)
        self.clo.start()
        self.clo.trigger.connect(self.close_window)

    def thread_terminate(self):
        self.reflect_export_thread.terminate()

    def Reflect_export(self,event):
        global Token,export_meta_list,export_meta_names
        if self.cho2.radioButton.isChecked()==True:
            Token = 1
        else:
            Token = 2
        export_meta_list = self.cho2.comboBox.SelectIndex()
        export_meta_names  = self.cho2.comboBox.Selectlist()
        file_name = self.cho2.lineEdit.text()
        if file_name.split('.')[-1]=='xls':
            self.cho2.close()
            self.reflect_export_thread = export_thread(file_name,self.x,self.y,self.points,self.intensity,self.operation,export_meta_list)
            self.progressBar = My_Progress_Form()
            self.progressBar.progressBar.setValue(0)
            self.progressBar.pushButton.setVisible(True)
            self.progressBar.pushButton.setText('Cancel')
            self.progressBar.pushButton.clicked.connect(self.thread_terminate)
            self.progressBar.show()
            self.reflect_export_thread.start()
            self.reflect_export_thread.trigger.connect(self.progress_update)
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

    def Subtract_draw(self):
        global figure2,cmap2,norm,som1,som2,startcolor,midcolor,endcolor,radio,grid_z1,cb,N_xlen,N_ylen
        c_index = self.cho.comboBox.currentIndex()
        self.cho.close()
        print('Generating Image, please wait....')
        self.msg = My_Message_Form()
        self.msg.label.setText('Generating Image, please wait....')
        self.msg.show()
        self.clo = close_widget_thread(4)
        self.clo.start()
        self.clo.trigger.connect(self.close_draw_window)
        data = xlrd.open_workbook(self.cho.lineEdit_2.text())
        table = data.sheets()[0]
        nrows = table.nrows
        subtract_points = np.array([[table.cell_value(1, 1), table.cell_value(1, 2)]])
        subtract_intensity = np.array([table.cell_value(1, 3*c_index+7)])

        for i in range(2, nrows):
            subtract_points = np.append(subtract_points, [[table.cell_value(i, 1), table.cell_value(i, 2)]], axis=0)
            subtract_intensity = np.append(subtract_intensity, [table.cell_value(i, 3*c_index+7)], axis=0)

        total_points = np.array([[1, 1]])
        total_intensity = np.array([0])

        blank_intensity = 1

        N_xlen = int(subtract_points[:,0].max()+5)
        N_ylen = int(subtract_points[:,1].max()+5)
        print(N_xlen,N_ylen)
        for x in range(1, N_xlen + 1):
            for y in range(1, N_ylen + 1):
                index = points_confirm(x, y, subtract_points)
                total_points = np.append(total_points, [[x, y]], axis=0)
                if index == -1:
                    total_intensity = np.append(total_intensity, [blank_intensity], axis=0)
                else:
                    total_intensity = np.append(total_intensity, [subtract_intensity[index]], axis=0)

        grid_x, grid_y = np.mgrid[1:N_xlen:(N_xlen * 10j), 1:N_ylen:(N_ylen * 10j)]

        grid_z1 = griddata(total_points, total_intensity, (grid_x, grid_y), method='linear')

        figure2 = plt.figure()
        plt.subplots_adjust(bottom=0.2, left=0.3)  # 调整子图间距
        #plt.rcParams['font.sans-serif'] = ['KaiTi']  # 指定默认字体
        plt.rcParams['axes.unicode_minus'] = False  # 解决保存图像是负号'-'显示为方块的问题

        cmap2 = col.LinearSegmentedColormap.from_list('own2', ['#ffffff','#00ebff','#00bfff','#007cd9','#000000','#e95b5b','#fc8b8b','#feabab','#ffffff'])
        img = plt.imshow(grid_z1.T, extent=(1, xlen, ylen, 1), cmap=cmap2)
        cb = plt.colorbar()
        plt.xticks(np.arange(1, N_xlen + 1, 5))
        plt.yticks(np.arange(1, N_ylen + 1, 5))
        om1 = plt.axes([0.25, 0.1, 0.65, 0.03])  # 第一slider的位置
        om2 = plt.axes([0.25, 0.05, 0.65, 0.03])  # 第一slider的位置
        som1 = Slider(om1, u'min: ', 0, 100, valstep=1, valinit=0)  # 产生第一slider
        som2 = Slider(om2, u'max: ', 0, 100, valstep=1, valinit=100)  # 产生第一slider
        som1.on_changed(update)
        som2.on_changed(update)

        maxval, minval = float(grid_z1.max()), float(grid_z1.min())
        s1 = som1.val
        s2 = som2.val
        vmin = minval + (maxval - minval) * s1 / 100
        vmax = maxval - (maxval - minval) * (100 - s2) / 100
        norm = col.Normalize(vmin=vmin, vmax=vmax)

        cc = plt.axes([0.025, 0.5, 0.2, 0.15])
        radio = RadioButtons(cc, ('re', 'ue', 'een'), active=0)
        radio.on_clicked(cmp_change)

        RemoveLearning_button_position = plt.axes([0.07, 0.590, 0.15, 0.05])
        RemoveLearning_button = MyButton(RemoveLearning_button_position,
                                         image=np.array(PIL.Image.open('colormap/RtoB.png')),
                                         image_pressed=np.array(PIL.Image.open('colormap/RtoB.png')))
        LRemoveLearning_button_position = plt.axes([0.07, 0.550, 0.15, 0.05])
        LRemoveLearning_button = MyButton(LRemoveLearning_button_position,
                                          image=np.array(PIL.Image.open('colormap/YtoB.png')),
                                          image_pressed=np.array(PIL.Image.open('colormap/YtoB.png')))
        GRemoveLearning_button_position = plt.axes([0.07, 0.513, 0.15, 0.05])
        GRemoveLearning_button = MyButton(GRemoveLearning_button_position,
                                          image=np.array(PIL.Image.open('colormap/GtoR.png')),
                                          image_pressed=np.array(PIL.Image.open('colormap/GtoR.png')))

        #plt.show()

    def Subtract_draw_choose(self,event):
        self.cho = My_choose_Form()
        self.cho.pushButton_3.clicked.connect(self.xls_open)
        self.cho.pushButton.clicked.connect(self.Subtract_draw)
        self.cho.show()

    def xls_open(self):
        try:
            self.cho.comboBox.clear()
            path = os.getcwd()
            file_name, _ = QtWidgets.QFileDialog.getOpenFileName(self.cho, 'Choose file', path, 'Excel files (*.xls *.xlsx)')
            self.cho.lineEdit_2.setText(file_name)
            if file_name != '':
                data = xlrd.open_workbook(file_name)
                table = data.sheets()[0]
                nrows = table.nrows
                items = table.col_values(0,1)
                f_items=[]
                for k in items:
                    if k!='': f_items.append(k)
                self.cho.comboBox.addItems(f_items)
        except Exception as e :
            m='Running error, info: '+str(e)
            self.error(m)

    def import_selection(self,event):
        try:
            selection_data = xlrd.open_workbook('Selection_point_export.xls')
            table = selection_data.sheets()[0]
            nrows = table.nrows
            temp_x,temp_y = [] , []
            for i in range(0,nrows):
                m=table.row_values(i)
                temp_x.append(m[0])
                temp_y.append(m[1])
            self.x , self.y = temp_x, temp_y
            ax = figure.add_subplot(1, 1, 1)
            ax.lines.clear()
            self.x.append(self.x[0])
            self.y.append(self.y[0])
            ax.plot(self.x[0], self.y[0], 'r+', ms=20)
            ax.plot(self.x[1:], self.y[1:], 'r*')
            ax.plot(self.x, self.y, 'w', linestyle='-.')
            self.x = self.x[:-1]
            self.y = self.y[:-1]
            figure.canvas.draw()
            self.msg = My_Message_Form()
            self.msg.label.setText('Selction Recording Data Imported Successfully!')
            self.msg.show()
            self.clo = close_widget_thread(1)
            self.clo.start()
            self.clo.trigger.connect(self.close_window)
        except Exception as e :
            m='Running error, info: '+str(e)
            self.error(m)

    def export_selection(self,event):
        try:
            f = xlwt.Workbook()
            sheet1 = f.add_sheet(u'sheet1', cell_overwrite_ok=True)
            for i in range(0,len(self.x)):
                sheet1.write(i,0,self.x[i])
                sheet1.write(i,1,self.y[i])
            f.save('Selection_point_export.xls')
            self.msg = My_Message_Form()
            self.msg.label.setText('Selection Recording Data Exported Successfully!')
            self.msg.show()
            self.clo = close_widget_thread(1)
            self.clo.start()
            self.clo.trigger.connect(self.close_window)
        except Exception as e :
            m='Running error, info: '+str(e)
            self.error(m)

class s_Form(object):
    def setupUi(self, Form):
        Form.setObjectName("Form")
        Form.resize(731, 352)
        icon = QtGui.QIcon()
        icon.addPixmap(QtGui.QPixmap(":/newPrefix/ui/icon.png"), QtGui.QIcon.Normal, QtGui.QIcon.Off)
        Form.setWindowIcon(icon)
        self.verticalLayout_2 = QtWidgets.QVBoxLayout(Form)
        self.verticalLayout_2.setObjectName("verticalLayout_2")
        self.verticalLayout = QtWidgets.QVBoxLayout()
        self.verticalLayout.setObjectName("verticalLayout")
        spacerItem = QtWidgets.QSpacerItem(20, 40, QtWidgets.QSizePolicy.Minimum, QtWidgets.QSizePolicy.Expanding)
        self.verticalLayout.addItem(spacerItem)
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
        sizePolicy = QtWidgets.QSizePolicy(QtWidgets.QSizePolicy.Minimum, QtWidgets.QSizePolicy.Maximum)
        sizePolicy.setHorizontalStretch(0)
        sizePolicy.setVerticalStretch(0)
        sizePolicy.setHeightForWidth(self.pushButton_2.sizePolicy().hasHeightForWidth())
        self.pushButton_2.setSizePolicy(sizePolicy)
        self.pushButton_2.setMinimumSize(QtCore.QSize(160, 0))
        self.pushButton_2.setMaximumSize(QtCore.QSize(120, 40))
        font = QtGui.QFont()
        font.setFamily("Arial")
        font.setBold(True)
        font.setPointSize(12)
        self.pushButton_2.setFont(font)
        icon1 = QtGui.QIcon()
        icon1.addPixmap(QtGui.QPixmap(":/newPrefix/Button_Image/file-open.png"), QtGui.QIcon.Normal, QtGui.QIcon.Off)
        self.pushButton_2.setIcon(icon1)
        self.pushButton_2.setIconSize(QtCore.QSize(35, 35))
        self.pushButton_2.setObjectName("pushButton_2")
        self.horizontalLayout.addWidget(self.pushButton_2)
        self.verticalLayout.addLayout(self.horizontalLayout)
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
        sizePolicy = QtWidgets.QSizePolicy(QtWidgets.QSizePolicy.Minimum, QtWidgets.QSizePolicy.Maximum)
        sizePolicy.setHorizontalStretch(0)
        sizePolicy.setVerticalStretch(0)
        sizePolicy.setHeightForWidth(self.pushButton_3.sizePolicy().hasHeightForWidth())
        self.pushButton_3.setSizePolicy(sizePolicy)
        self.pushButton_3.setMinimumSize(QtCore.QSize(140, 0))
        self.pushButton_3.setMaximumSize(QtCore.QSize(120, 40))
        font = QtGui.QFont()
        font.setFamily("Arial")
        font.setBold(True)
        font.setPointSize(12)
        self.pushButton_3.setFont(font)
        icon2 = QtGui.QIcon()
        icon2.addPixmap(QtGui.QPixmap(":/newPrefix/Button_Image/excel.png"), QtGui.QIcon.Normal, QtGui.QIcon.Off)
        self.pushButton_3.setIcon(icon2)
        self.pushButton_3.setIconSize(QtCore.QSize(35, 35))
        self.pushButton_3.setObjectName("pushButton_3")
        self.horizontalLayout_2.addWidget(self.pushButton_3)
        self.verticalLayout.addLayout(self.horizontalLayout_2)
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
        self.verticalLayout.addLayout(self.horizontalLayout_3)
        spacerItem1 = QtWidgets.QSpacerItem(513, 13, QtWidgets.QSizePolicy.Minimum, QtWidgets.QSizePolicy.Expanding)
        self.verticalLayout.addItem(spacerItem1)
        self.horizontalLayout_4 = QtWidgets.QHBoxLayout()
        self.horizontalLayout_4.setContentsMargins(-1, 10, -1, -1)
        self.horizontalLayout_4.setObjectName("horizontalLayout_4")
        spacerItem2 = QtWidgets.QSpacerItem(40, 20, QtWidgets.QSizePolicy.Expanding, QtWidgets.QSizePolicy.Minimum)
        self.horizontalLayout_4.addItem(spacerItem2)
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
        spacerItem3 = QtWidgets.QSpacerItem(40, 20, QtWidgets.QSizePolicy.Expanding, QtWidgets.QSizePolicy.Minimum)
        self.horizontalLayout_4.addItem(spacerItem3)
        self.horizontalLayout_4.setStretch(0, 2)
        self.horizontalLayout_4.setStretch(1, 3)
        self.horizontalLayout_4.setStretch(2, 2)
        self.verticalLayout.addLayout(self.horizontalLayout_4)
        self.verticalLayout_2.addLayout(self.verticalLayout)

        self.retranslateUi(Form)
        self.pushButton.clicked.connect(Form.imzml_plot)
        self.pushButton_2.clicked.connect(Form.imzml_file_open)
        self.pushButton_3.clicked.connect(Form.xls_file_open)
        QtCore.QMetaObject.connectSlotsByName(Form)

    def retranslateUi(self, Form):
        _translate = QtCore.QCoreApplication.translate
        Form.setWindowTitle(_translate("Form", "imzml corresponding selection data extraction"))
        self.label.setText(_translate("Form", "MSI data file path: "))
        self.pushButton_2.setText(_translate("Form", "Choose file"))
        self.label_2.setText(_translate("Form", "Metabolites data xls file path: "))
        self.pushButton_3.setText(_translate("Form", "Choose file"))
        self.label_3.setText(_translate("Form", "Target Metabolite: "))
        self.pushButton.setText(_translate("Form", "Start Drawing"))

class My_Form(QtWidgets.QWidget,s_Form):
    def __init__(self):
        super().__init__()
        self.setupUi(self)

    def close_progressbar(self,val):
        if val==100:
            self.progressBar.close()

    def error(self,m):
        self.eW=My_Error_Form()
        self.eW.label.setText(m)
        self.eW.show()

    def thread_terminate(self):
        self.mbt.terminate()

    def imzml_plot(self):
        try:
            global fig,index,reverse_button,Meta_Names,Meta_Names_M,Meta_Names_M_start,Meta_Names_M_end
            index = self.comboBox.currentIndex()
            Start = Meta_Names_M_start[index]
            End = Meta_Names_M_end[index]
            print(index,Start,End)
            self.mbt = Imzml_draw_thread(self.lineEdit.text(),float(Start),float(End))
            self.progressBar = My_Progress_Form()
            self.progressBar.progressBar.setValue(0)
            self.progressBar.pushButton.setVisible(True)
            self.progressBar.pushButton.setText('Cancel')
            self.progressBar.pushButton.clicked.connect(self.thread_terminate)
            self.progressBar.show()
            self.mbt.trigger.connect(self.progress_update)
            self.mbt.trigger2.connect(self.plot)
            self.mbt.start()
        except Exception as e :
            m='Running error, info: '+str(e)
            self.error(m)

    def imzml_file_open(self):
        path = os.getcwd()
        file_name,_= QtWidgets.QFileDialog.getOpenFileName(self,u'Choose imzml file',path,'imzml files (*.imzml)')
        self.lineEdit.setText(file_name)

    def xls_file_open(self):
        try:
            global Meta_Names,Meta_Names_M,Meta_Names_M_start,Meta_Names_M_end
            self.comboBox.clear()
            path = os.getcwd()
            file_name,_= QtWidgets.QFileDialog.getOpenFileName(self,u'Choose xls file',path,'Excel files (*.xls *.xlsx)')
            self.lineEdit_2.setText(file_name)
            if file_name!='':
                data = xlrd.open_workbook(file_name)
                table = data.sheets()[0]
                nrows = table.nrows
                Meta_Names = table.col_values(0,1)
                Meta_Names_M = table.col_values(1,1)
                Meta_Names_M_start = table.col_values(2,1)
                Meta_Names_M_end = table.col_values(3,1)
                self.comboBox.addItems(Meta_Names)
        except Exception as e :
            m='Running error, info: '+str(e)
            self.error(m)

    def progress_update(self,val,stry):
        if val!=-1:
            self.progressBar.progressBar.setValue(val)
            if val==100:
                self.progressBar.label.setText('Finished!')
                self.clo = close_widget_thread(1)
                self.clo.start()
                self.clo.trigger.connect(self.close_progressbar)
        else:
            self.progressBar.label.setText('Running error, info: '+stry)
            self.progressBar.progressBar.setValue(0)
            self.progressBar.pushButton.setVisible(True)

    def close_window(self,val):
        if val==100:
            self.progressBar.close()

    def plot(self,grid_z0,xlen,ylen,points,intensity):
        try:
            global figure,img,LearningStatus_box,reelect_button,RemoveLearning_button,StartLearning_button,norm
            global EndLearning_button,reverse_button,zuoyou_reverse_button,Rotate_box,Rotate_button
            global Operation_import_button,Operation_export_button,Operation_button,Reflect_export_button
            global Subtract_export_button,cb1,img,DDsom1,DDsom2,Import_selection_button,Export_selection_button
            figure = plt.figure()
            plt.subplots_adjust(left=0, right=0.7, top=0.9, bottom=0.2)
            img = plt.imshow(grid_z0.T, extent=(1, xlen, ylen, 1))
            cb1 = plt.colorbar()
            plt.xticks(np.arange(1, xlen + 1, 5))
            plt.yticks(np.arange(1, ylen + 1, 5))
            DDom1 = plt.axes([0.25, 0.1, 0.45, 0.03])  # 第一slider的位置
            DDom2 = plt.axes([0.25, 0.05, 0.45, 0.03])  # 第一slider的位置
            DDsom1 = Slider(DDom1, u'min: ', 0, 100, valstep=1, valinit=0)  # 产生第一slider
            DDsom2 = Slider(DDom2, u'max: ', 0, 100, valstep=1, valinit=100)  # 产生第一slider
            maxval ,minval = float(grid_z0.max()),float(grid_z0.min())
            s1 = DDsom1.val
            s2 = DDsom2.val
            vmin = minval+(maxval-minval)*s1/100
            vmax = maxval - (maxval - minval) * (100-s2) / 100
            norm = col.Normalize(vmin = vmin,vmax = vmax)
            DDsom1.on_changed(update2)
            DDsom2.on_changed(update2)
            print('x  ', xlen)
            print('y  ', ylen)
            callback = ButtonHandler(figure, points, intensity)
            callback.connect()

            LearningStatus_box_position = plt.axes([0.8, 0.9, 0.1, 0.08])
            LearningStatus_box = TextBox(LearningStatus_box_position, u'Learning status:', label_pad=0.2)
            LearningStatus_box.set_val('Off')

            reelect_button_position = plt.axes([0.75, 0.8, 0.1, 0.08])
            reelect_button = MyButton(reelect_button_position, '',
                                      np.array(PIL.Image.open(u'Button_Image_e/Reselect.png')),
                                      np.array(PIL.Image.open('Button_Image_e/Reselect_pressed.png')))
            reelect_button.on_clicked(callback.reelect)

            RemoveLearning_button_position = plt.axes([0.85, 0.8, 0.1, 0.08])
            RemoveLearning_button = MyButton(RemoveLearning_button_position, '',
                                             np.array(PIL.Image.open(u'Button_Image_e/Remove_Learning.png')),
                                             np.array(PIL.Image.open('Button_Image_e/Remove_Learning_pressed.png')))
            RemoveLearning_button.on_clicked(callback.RemoveLearning)

            StartLearing_button_position = plt.axes([0.75, 0.7, 0.1, 0.08])
            StartLearning_button = MyButton(StartLearing_button_position, '',
                                            np.array(PIL.Image.open(u'Button_Image_e/Starting_Learning.png')),
                                            np.array(PIL.Image.open('Button_Image_e/Start_Learning_pressed.png')))
            StartLearning_button.on_clicked(callback.StartLearning)

            EndLearning_button_position = plt.axes([0.85, 0.7, 0.1, 0.08])
            EndLearning_button = MyButton(EndLearning_button_position, '',
                                          np.array(PIL.Image.open(u'Button_Image_e/Stop_Learning.png')),
                                          np.array(PIL.Image.open('Button_Image_e/Stop_Learning_pressed.png')))
            EndLearning_button.on_clicked(callback.EndLearning)

            reverse_button_position = plt.axes([0.75, 0.6, 0.1, 0.08])
            reverse_button = MyButton(reverse_button_position, '',
                                      np.array(PIL.Image.open(u'Button_Image_e/Up-down-flip.png')),
                                      np.array(PIL.Image.open('Button_Image_e/Up-down-flip-pressed.png')))
            reverse_button.on_clicked(callback.reverseshape)

            zuoyou_reverse_button_position = plt.axes([0.85, 0.6, 0.1, 0.08])
            zuoyou_reverse_button = MyButton(zuoyou_reverse_button_position , '',
                                             np.array(PIL.Image.open(u'Button_Image_e/left-right-flip.png')),
                                             np.array(PIL.Image.open('Button_Image_e/left-right-flip-pressed.png')))
            zuoyou_reverse_button.on_clicked(callback.zuoyou_reverseshape)

            Import_selection_button_position = plt.axes([0.75, 0.5, 0.1, 0.08])
            Import_selection_button = MyButton(Import_selection_button_position , '',
                                             np.array(PIL.Image.open(u'Button_Image_e/Import_selection.png')),
                                             np.array(PIL.Image.open('Button_Image_e/Import_selection_pressed.png')))
            Import_selection_button.on_clicked(callback.import_selection)

            Export_selection_button_position = plt.axes([0.85, 0.5, 0.1, 0.08])
            Export_selection_button = MyButton(Export_selection_button_position , '',
                                             np.array(PIL.Image.open(u'Button_Image_e/Export_selection.png')),
                                             np.array(PIL.Image.open('Button_Image_e/Export_selection_pressed.png')))
            Export_selection_button.on_clicked(callback.export_selection)

            Rotate_box_position = plt.axes([0.8, 0.4, 0.1, 0.08])
            Rotate_box = TextBox(Rotate_box_position, u'Angle (clockwise positive):', label_pad=0.2)

            Rotate_button_position = plt.axes([0.75, 0.3, 0.1, 0.08])
            Rotate_button = MyButton(Rotate_button_position , '',
                                     np.array(PIL.Image.open(u'Button_Image_e/Rotation_Transform.png')),
                                     np.array(PIL.Image.open('Button_Image_e/Rotation_Transform_pressed.png')))
            Rotate_button.on_clicked(callback.rotateshape)

            Operation_button_position = plt.axes([0.85, 0.3, 0.1, 0.08])
            Operation_button = MyButton(Operation_button_position , '',
                                        np.array(PIL.Image.open(u'Button_Image_e/Reflect_generate.png')),
                                        np.array(PIL.Image.open('Button_Image_e/Reflect_generate_pressed.png')))
            Operation_button.on_clicked(callback.do_operation)

            Operation_import_button_position = plt.axes([0.75, 0.2, 0.1, 0.08])
            Operation_import_button = MyButton(Operation_import_button_position , '',
                                               np.array(PIL.Image.open(u'Button_Image_e/Import_Learning.png')),
                                               np.array(PIL.Image.open('Button_Image_e/Import_Learning_pressed.png')))
            Operation_import_button.on_clicked(callback.Operation_import)

            Operation_export_button_position = plt.axes([0.85, 0.2, 0.1, 0.08])
            Operation_export_button = MyButton(Operation_export_button_position  , '',
                                               np.array(PIL.Image.open(u'Button_Image_e/Export_Learning.png')),
                                               np.array(PIL.Image.open('Button_Image_e/Export_Learning_pressed.png')))
            Operation_export_button.on_clicked(callback.Operation_export)

            Reflect_export_button_position = plt.axes([0.75, 0.1, 0.1, 0.08])
            Reflect_export_button = MyButton(Reflect_export_button_position  , '',
                                             np.array(PIL.Image.open(u'Button_Image_e/Selection_export.png')),
                                             np.array(PIL.Image.open('Button_Image_e/Selection_export_pressed.png')))
            Reflect_export_button.on_clicked(callback.Reflect_export_choose)

            Subtract_export_button_position = plt.axes([0.85, 0.1, 0.1, 0.08])
            Subtract_export_button = MyButton(Subtract_export_button_position , '',
                                              np.array(PIL.Image.open(u'Button_Image_e/Subtract-Image-Generate.png')),
                                              np.array(PIL.Image.open('Button_Image_e/Subtract-Image-Generate-pressed.png')))
            Subtract_export_button.on_clicked(callback.Subtract_draw_choose)
            plt.show()
        except Exception as e:
            m = 'Running error, info: ' + str(e)
            self.error(m)

class Imzml_draw_thread(QtCore.QThread):
    trigger = QtCore.pyqtSignal(int,str)
    trigger2 = QtCore.pyqtSignal(np.ndarray,int,int,np.ndarray,np.ndarray)

    def __init__(self,a,b,c):
        super().__init__()
        self.inputpath = a
        self.left = b
        self.right = c

    def run(self):
        try:
            global p,points,intensity,grid_z0,xlen,ylen
            self.trigger.emit(5,'')
            p = ImzMLParser(self.inputpath)
            Coor = p.coordinates
            total = len(Coor)

            m = p.getspectrum(0)
            points = np.array([list(Coor[0][:-1])])
            Tempinte = PeakIntensitySum(m, self.left, self.right)
            intensity = np.array([Tempinte])
            self.trigger.emit(10,'')
            kkk = 10
            for indecount in range(1, total):
                v1= int((indecount / (total - 1)) * 100)
                if v1-kkk>=1:
                    self.trigger.emit(v1,'')
                    kkk = v1
                m = p.getspectrum(indecount)
                points = np.append(points, [list(Coor[indecount][:-1])], axis=0)
                intensity = np.append(intensity, [PeakIntensitySum(m, self.left, self.right)], axis=0)

            xlen = np.array(Coor)[:,0].max()
            ylen = np.array(Coor)[:,1].max()

            grid_x, grid_y = np.mgrid[1:xlen:(xlen * 10j), 1:ylen:(ylen * 10j)]

            grid_z0 = griddata(points, intensity, (grid_x, grid_y), method='linear')
            self.trigger2.emit(grid_z0,xlen,ylen,points,intensity)
        except Exception as e:
            m = 'Running error, info: ' + str(e)
            self.trigger.emit(-1, m)

class export_thread(QtCore.QThread):
    trigger = QtCore.pyqtSignal(int,str)
    def __init__(self,file_name,x, y, point, intensity,operation,export_list):
        super().__init__()
        self.file_name = file_name
        self.x = x
        self.y = y
        self.point = point
        self.intensity = intensity
        self.operation = operation
        self.export_meta_list = export_list

    def reflect(self,x, y, operation):
        m = Shape_Operation()
        x1, y1 = x, y
        for i in range(0, len(operation)):
            if len(operation[i]) == 2:
                x1, y1 = m.lmove(x1, y1, operation[i][0], operation[i][1])
            elif len(operation[i]) == 3:
                if operation[i][-1]==1:
                    x1, y1 = m.lreverse(x1, y1, operation[i][0], operation[i][1])
                else:
                    x1, y1 = m.lzuoyoureverse(x1, y1, operation[i][0], operation[i][1])
            elif len(operation[i]) == 4:
                x1, y1 = m.lrotate(x1, y1, operation[i][0], operation[i][1], operation[i][2])
        return x1, y1

    def run(self):
        try:
            global Token,p,index,Meta_Names,Meta_Names_M,Meta_Names_M_start,Meta_Names_M_end
            print('Start exporting selection data, please wait....')
            loc = np.array([[0, 0]])
            for i in range(0, len(self.x)):
                loc = np.append(loc, [[self.x[i], self.y[i]]], axis=0)
            polygon = Polygon(loc[1:], True)
            iden = polygon.contains_points(self.point)
            export_points_x, export_points_y, export_points_intensity = [], [], []
            for i in range(0, len(self.point)):
                if iden[i] == True:
                    export_points_x.append(self.point[i][0])
                    export_points_y.append(self.point[i][1])
                    export_points_intensity.append(intensity[i])
            export_reflect_x, export_reflect_y = self.reflect(export_points_x, export_points_y, self.operation)
            export_reflect_intensity = []
            key = 0
            for i in range(0, len(export_points_y)):
                key1 = int((i / (5*len(export_points_y))) * 100)
                if key1-key>=1 and key1<=48:
                    self.trigger.emit(key1,'')
                    key = key1
                flag_point = int(export_reflect_y[i] - 1) * (xlen - 1)
                flag = False
                while not flag:
                    if ((self.point[flag_point][0] - export_reflect_x[i]) ** 2 <= 0.25) and (
                            (self.point[flag_point][1] - export_reflect_y[i]) ** 2 <= 0.25):
                        export_reflect_x[i], export_reflect_y[i] = self.point[flag_point][0], self.point[flag_point][1]
                        export_reflect_intensity.append(intensity[flag_point])
                        flag = True
                    flag_point += 1
                    if flag_point >= len(self.point):
                        print(export_reflect_x[i],export_reflect_y[i])
                        flag = True
            f = xlwt.Workbook()
            sheet1 = f.add_sheet(u'sheet1', cell_overwrite_ok=True)
            row0 = ['Metabolites Names','Selection x', 'Selection y', 'Reflect x', 'Reflect y']
            for i in self.export_meta_list:
                row0.append(str(Meta_Names_M[i])+'Selection intensity')
                row0.append(str(Meta_Names_M[i]) + 'Reflect intensity')
                if Token==1:
                    row0.append(str(Meta_Names_M[i]) +'Intensity difference')
                else:
                    row0.append(str(Meta_Names_M[i]) + u'Intensity ratio')
            for i in range(0,len(self.export_meta_list)):
                nm = self.export_meta_list[i]
                sheet1.write(i+1,0,Meta_Names[nm])
            for i in range(0, len(row0)):
                sheet1.write(0, i, row0[i])
            key = 48
            print(index)
            print('dfeoijeorje')
            for i in range(0, len(export_reflect_x)):
                key1 = int((i / (2 * len(export_reflect_x))) * 100)+48
                if key1 - key >= 1:
                    self.trigger.emit(key1, '')
                    key = key1
                sheet1.write(i + 1, 1, float(export_points_x[i]))
                sheet1.write(i + 1, 2, float(export_points_y[i]))
                sheet1.write(i + 1, 3, float(export_reflect_x[i]))
                sheet1.write(i + 1, 4, float(export_reflect_y[i]))
                for j in range(0,len(self.export_meta_list)):
                    meta = self.export_meta_list[j]
                    if meta ==index:
                        sheet1.write(i + 1, 3*j+5, float(export_points_intensity[i]))
                        sheet1.write(i + 1, 3*j+6, float(export_reflect_intensity[i]))
                        if Token == 1:
                            sheet1.write(i + 1, 3*j+7, -float(export_points_intensity[i]) + float(export_reflect_intensity[i]))
                        else:
                            if float(export_points_intensity[i])!=0:
                                sheet1.write(i + 1, 3 * j + 7,float(export_reflect_intensity[i])/float(export_points_intensity[i]))
                            else:
                                sheet1.write(i + 1, 3 * j + 7,float(export_reflect_intensity[i]))
                    else:
                        export_i = imzml_intensity(export_points_x[i],export_points_y[i],float(Meta_Names_M_start[meta]),float(Meta_Names_M_end[meta]))
                        export_reflect_i = imzml_intensity(export_reflect_x[i],export_reflect_y[i],float(Meta_Names_M_start[meta]),float(Meta_Names_M_end[meta]))
                        sheet1.write(i + 1, 3 * j + 5, float(export_i))
                        sheet1.write(i + 1, 3 * j + 6, float(export_reflect_i))
                        if Token ==1 :
                            sheet1.write(i + 1, 3 * j + 7,-float(export_i) + float(export_reflect_i))
                        else:
                            if float(export_i)!=0:
                                sheet1.write(i + 1, 3 * j + 7, float(export_reflect_i)/float(export_i))
                            else:
                                sheet1.write(i + 1, 3 * j + 7, float(export_reflect_i))

            f.save(self.file_name)
            print('Selection MSI data exported successfully!')
            self.trigger.emit(100,'')
        except Exception as e:
            m = 'Running error, info: ' + str(e)
            self.trigger.emit(-1,m)

class close_widget_thread(QtCore.QThread):
    trigger = QtCore.pyqtSignal(int)

    def __init__(self,seconds):
        super().__init__()
        self.second = seconds

    def run(self):
        self.sleep(self.second)
        self.trigger.emit(100)

if __name__ == "__main__" :
    QtCore.QCoreApplication.setAttribute(QtCore.Qt.AA_EnableHighDpiScaling)
    app = QtWidgets.QApplication(sys.argv)
    my_pyqt_form = My_Form()
    my_pyqt_form.show()
    sys.exit(app.exec_())
