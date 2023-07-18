# -*- coding: utf-8 -*-

# Form implementation generated from reading ui file 'Main.ui'
#
# Created by: PyQt5 UI code generator 5.15.9
#
# WARNING: Any manual changes made to this file will be lost when pyuic5 is
# run again.  Do not edit this file unless you know what you are doing.


from PyQt5 import QtCore, QtGui, QtWidgets
import os


class Button(QtWidgets.QLineEdit):

	def __init__(self, parent):
		super(Button, self).__init__(parent)

		self.setAcceptDrops(True)

	def dragEnterEvent(self, e):

		if e.mimeData().hasUrls():
			e.accept()
		else:
			super(Button, self).dragEnterEvent(e)

	def dragMoveEvent(self, e):

		super(Button, self).dragMoveEvent(e)

	def dropEvent(self, e):

		if e.mimeData().hasUrls():
			for url in e.mimeData().urls():
				self.setText(os.path.normcase(url.toLocalFile()))
				e.accept()
		else:
			super(Button, self).dropEvent(e)


class Ui_MainWindow(object):
    def setupUi(self, MainWindow):
        MainWindow.setObjectName("MainWindow")
        MainWindow.resize(494, 276)
        self.centralwidget = QtWidgets.QWidget(MainWindow)
        self.centralwidget.setObjectName("centralwidget")
        self.gridLayout = QtWidgets.QGridLayout(self.centralwidget)
        self.gridLayout.setObjectName("gridLayout")
        self.lineEdit_path_finish_folder = Button(self.centralwidget)
        font = QtGui.QFont()
        font.setPointSize(10)
        self.lineEdit_path_finish_folder.setFont(font)
        self.lineEdit_path_finish_folder.setObjectName("lineEdit_path_finish_folder")
        self.gridLayout.addWidget(self.lineEdit_path_finish_folder, 1, 1, 1, 1)
        self.pushButton_open_data_file = QtWidgets.QPushButton(self.centralwidget)
        font = QtGui.QFont()
        font.setPointSize(10)
        self.pushButton_open_data_file.setFont(font)
        self.pushButton_open_data_file.setObjectName("pushButton_open_data_file")
        self.gridLayout.addWidget(self.pushButton_open_data_file, 0, 2, 1, 1)
        self.progressBar = QtWidgets.QProgressBar(self.centralwidget)
        font = QtGui.QFont()
        font.setPointSize(10)
        self.progressBar.setFont(font)
        self.progressBar.setProperty("value", 0)
        self.progressBar.setObjectName("progressBar")
        self.gridLayout.addWidget(self.progressBar, 4, 0, 1, 3)
        self.lineEdit_path_data_file = Button(self.centralwidget)
        font = QtGui.QFont()
        font.setPointSize(10)
        self.lineEdit_path_data_file.setFont(font)
        self.lineEdit_path_data_file.setObjectName("lineEdit_path_data_file")
        self.gridLayout.addWidget(self.lineEdit_path_data_file, 0, 1, 1, 1)
        self.pushButton_open_finish_folder = QtWidgets.QPushButton(self.centralwidget)
        font = QtGui.QFont()
        font.setPointSize(10)
        self.pushButton_open_finish_folder.setFont(font)
        self.pushButton_open_finish_folder.setObjectName("pushButton_open_finish_folder")
        self.gridLayout.addWidget(self.pushButton_open_finish_folder, 1, 2, 1, 1)
        self.pushButton_start = QtWidgets.QPushButton(self.centralwidget)
        self.pushButton_start.setObjectName("pushButton_start")
        self.gridLayout.addWidget(self.pushButton_start, 3, 0, 1, 3)
        self.label_data_file = QtWidgets.QLabel(self.centralwidget)
        font = QtGui.QFont()
        font.setPointSize(10)
        self.label_data_file.setFont(font)
        self.label_data_file.setObjectName("label_data_file")
        self.gridLayout.addWidget(self.label_data_file, 0, 0, 1, 1)
        self.label_finish_folder = QtWidgets.QLabel(self.centralwidget)
        font = QtGui.QFont()
        font.setPointSize(10)
        self.label_finish_folder.setFont(font)
        self.label_finish_folder.setObjectName("label_finish_folder")
        self.gridLayout.addWidget(self.label_finish_folder, 1, 0, 1, 1)
        self.label_file_name = QtWidgets.QLabel(self.centralwidget)
        font = QtGui.QFont()
        font.setPointSize(10)
        self.label_file_name.setFont(font)
        self.label_file_name.setObjectName("label_file_name")
        self.gridLayout.addWidget(self.label_file_name, 2, 0, 1, 1)
        self.lineEdit_file_name = QtWidgets.QLineEdit(self.centralwidget)
        font = QtGui.QFont()
        font.setPointSize(10)
        self.lineEdit_file_name.setFont(font)
        self.lineEdit_file_name.setObjectName("lineEdit_file_name")
        self.gridLayout.addWidget(self.lineEdit_file_name, 2, 1, 1, 2)
        MainWindow.setCentralWidget(self.centralwidget)
        self.menubar = QtWidgets.QMenuBar(MainWindow)
        self.menubar.setGeometry(QtCore.QRect(0, 0, 494, 21))
        self.menubar.setObjectName("menubar")
        self.menu = QtWidgets.QMenu(self.menubar)
        self.menu.setObjectName("menu")
        MainWindow.setMenuBar(self.menubar)
        self.statusbar = QtWidgets.QStatusBar(MainWindow)
        self.statusbar.setObjectName("statusbar")
        MainWindow.setStatusBar(self.statusbar)
        self.action_settings_default = QtWidgets.QAction(MainWindow)
        self.action_settings_default.setObjectName("action_settings_default")
        self.action_settings_table = QtWidgets.QAction(MainWindow)
        self.action_settings_table.setObjectName("action_settings_table")
        self.menu.addAction(self.action_settings_default)
        self.menubar.addAction(self.menu.menuAction())

        self.retranslateUi(MainWindow)
        QtCore.QMetaObject.connectSlotsByName(MainWindow)

    def retranslateUi(self, MainWindow):
        _translate = QtCore.QCoreApplication.translate
        MainWindow.setWindowTitle(_translate("MainWindow", "Главное окно"))
        self.pushButton_open_data_file.setText(_translate("MainWindow", "Открыть"))
        self.pushButton_open_finish_folder.setText(_translate("MainWindow", "Открыть"))
        self.pushButton_start.setText(_translate("MainWindow", "Создать файл"))
        self.label_data_file.setText(_translate("MainWindow", "Файл выгрузки"))
        self.label_finish_folder.setText(_translate("MainWindow", "Конечная папка"))
        self.label_file_name.setText(_translate("MainWindow", "Имя файла"))
        self.menu.setTitle(_translate("MainWindow", "Настройки"))
        self.action_settings_default.setText(_translate("MainWindow", "Настройки по умолчанию"))
        self.action_settings_table.setText(_translate("MainWindow", "Настройки таблицы"))
