# -*- coding: utf-8 -*-

# Form implementation generated from reading ui file 'get_input_files.ui'
#
# Created by: PyQt5 UI code generator 5.9.2
#
# WARNING! All changes made in this file will be lost!
import logging
import os
import time

from PyQt5 import QtCore, QtGui, QtWidgets
import MainReportGenerator

class Ui_ReadFiles(object):
    def setupUi(self, Dialog):
        Dialog.setObjectName("Generate Report : Enter File Names")
        Dialog.resize(817, 180)
        self.buttonBox = QtWidgets.QDialogButtonBox(Dialog)
        self.buttonBox.setGeometry(QtCore.QRect(30, 130, 761, 32))
        self.buttonBox.setOrientation(QtCore.Qt.Horizontal)
        self.buttonBox.setStandardButtons(QtWidgets.QDialogButtonBox.Cancel | QtWidgets.QDialogButtonBox.Save)
        self.buttonBox.setObjectName("buttonBox")
        self.label = QtWidgets.QLabel(Dialog)
        self.label.setGeometry(QtCore.QRect(20, 25, 55, 21))
        self.label.setObjectName("label")
        self.label_2 = QtWidgets.QLabel(Dialog)
        self.label_2.setGeometry(QtCore.QRect(20, 80, 55, 21))
        self.label_2.setObjectName("label_2")
        self.lineEdit = QtWidgets.QLineEdit(Dialog)
        self.lineEdit.setGeometry(QtCore.QRect(90, 21, 601, 31))
        self.lineEdit.setObjectName("lineEdit")
        self.lineEdit_2 = QtWidgets.QLineEdit(Dialog)
        self.lineEdit_2.setGeometry(QtCore.QRect(90, 70, 601, 31))
        self.lineEdit_2.setObjectName("lineEdit_2")
        self.pushButton = QtWidgets.QPushButton(Dialog)
        self.pushButton.setGeometry(QtCore.QRect(710, 20, 93, 28))
        self.pushButton.setObjectName("pushButton")
        self.pushButton_2 = QtWidgets.QPushButton(Dialog)
        self.pushButton_2.setGeometry(QtCore.QRect(710, 70, 93, 28))
        self.pushButton_2.setObjectName("pushButton_2")

        self.retranslateUi(Dialog)
        self.buttonBox.accepted.connect(Dialog.accept)
        self.buttonBox.rejected.connect(Dialog.reject)
        QtCore.QMetaObject.connectSlotsByName(Dialog)

        self.pushButton.clicked.connect(self.get_excel_file)
        output = self.pushButton_2.clicked.connect(self.get_output_location)
        Dialog.accepted.connect(lambda: self.call_generate_report(self.lineEdit.text(), self.lineEdit_2.text()))

    def retranslateUi(self, Dialog):
        _translate = QtCore.QCoreApplication.translate
        Dialog.setWindowTitle(_translate("Dialog", "Dialog"))
        self.label.setText(_translate("Dialog", "Excel File"))
        self.label_2.setText(_translate("Dialog", "Output File"))
        self.pushButton.setText(_translate("Dialog", "Browse"))
        self.pushButton_2.setText(_translate("Dialog", "Browse"))

    def get_excel_file(self):
        fileinput = QtWidgets.QFileDialog.getOpenFileName(self.pushButton, 'Select Input Excel File', "C:\\", "Excel (*.xls *.xlsx)")
        self.lineEdit.setText(os.path.normpath(str(fileinput[0])))
        return os.path.normpath(str(fileinput[0]))

    def get_output_location(self):
        fileinput = QtWidgets.QFileDialog.getExistingDirectory(self.pushButton_2, 'Select Output Directory')
        self.lineEdit_2.setText(os.path.normpath(str(fileinput)))
        return os.path.normpath(str(fileinput))

    def call_generate_report(self, filename, output):
        MainReportGenerator.Pass_Params(filename, output)
        return True

if __name__ == "__main__":
    import sys
    app = QtWidgets.QApplication(sys.argv)
    Dialog = QtWidgets.QDialog()
    ui = Ui_ReadFiles()
    ui.setupUi(Dialog)
    Dialog.show()
    result = app.exec()
    sys.exit(result)
