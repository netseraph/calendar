# -*- coding: utf-8 -*-

################################################################################
## Form generated from reading UI file 'MyDialog.ui'
##
## Created by: Qt User Interface Compiler version 6.10.1
##
## WARNING! All changes made in this file will be lost when recompiling UI file!
################################################################################

from PySide6.QtCore import (QCoreApplication, QDate, QDateTime, QLocale,
    QMetaObject, QObject, QPoint, QRect,
    QSize, QTime, QUrl, Qt)
from PySide6.QtGui import (QBrush, QColor, QConicalGradient, QCursor,
    QFont, QFontDatabase, QGradient, QIcon,
    QImage, QKeySequence, QLinearGradient, QPainter,
    QPalette, QPixmap, QRadialGradient, QTransform)
from PySide6.QtWidgets import (QAbstractButton, QApplication, QDialog, QDialogButtonBox,
    QLabel, QLineEdit, QSizePolicy, QSpinBox,
    QToolButton, QWidget)

class Ui_Dialog(object):
    def setupUi(self, Dialog):
        if not Dialog.objectName():
            Dialog.setObjectName(u"Dialog")
        Dialog.resize(400, 250)
        self.buttonBox = QDialogButtonBox(Dialog)
        self.buttonBox.setObjectName(u"buttonBox")
        self.buttonBox.setGeometry(QRect(160, 180, 221, 41))
        font = QFont()
        font.setFamilies([u"Arial"])
        font.setPointSize(9)
        font.setBold(False)
        self.buttonBox.setFont(font)
        self.buttonBox.setOrientation(Qt.Orientation.Horizontal)
        self.buttonBox.setStandardButtons(QDialogButtonBox.StandardButton.Cancel|QDialogButtonBox.StandardButton.Ok)
        self.buttonBox.setCenterButtons(True)
        self.label_date = QLabel(Dialog)
        self.label_date.setObjectName(u"label_date")
        self.label_date.setGeometry(QRect(10, 10, 38, 22))
        self.label_date.setFont(font)
        self.toolButton = QToolButton(Dialog)
        self.toolButton.setObjectName(u"toolButton")
        self.toolButton.setGeometry(QRect(350, 50, 29, 22))
        self.toolButton.setFont(font)
        self.lineEdit_folder = QLineEdit(Dialog)
        self.lineEdit_folder.setObjectName(u"lineEdit_folder")
        self.lineEdit_folder.setGeometry(QRect(70, 50, 280, 22))
        self.lineEdit_folder.setFont(font)
        self.lineEdit_folder.setReadOnly(True)
        self.label_folder = QLabel(Dialog)
        self.label_folder.setObjectName(u"label_folder")
        self.label_folder.setGeometry(QRect(10, 50, 71, 22))
        self.label_folder.setFont(font)
        self.spinBox = QSpinBox(Dialog)
        self.spinBox.setObjectName(u"spinBox")
        self.spinBox.setGeometry(QRect(70, 10, 80, 22))
        self.spinBox.setMinimum(2020)
        self.spinBox.setMaximum(2030)

        self.retranslateUi(Dialog)
        self.buttonBox.accepted.connect(Dialog.accept)
        self.buttonBox.rejected.connect(Dialog.reject)

        QMetaObject.connectSlotsByName(Dialog)
    # setupUi

    def retranslateUi(self, Dialog):
        Dialog.setWindowTitle(QCoreApplication.translate("Dialog", u"\u6211\u7684\u65e5\u7a0b\u8868", None))
        self.label_date.setText(QCoreApplication.translate("Dialog", u"\u65e5\u671f:", None))
        self.toolButton.setText(QCoreApplication.translate("Dialog", u"...", None))
        self.label_folder.setText(QCoreApplication.translate("Dialog", u"\u5de5\u4f5c\u76ee\u5f55:", None))
    # retranslateUi

