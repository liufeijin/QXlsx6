#
# sat_gui.pro
#
# Copyright (c) 2018 edosad(github) All rights reserved.
#
# Some code is fixed by j2doll(github)

QT += core
QT += gui
QT += widgets

TARGET = Chromatogram
TEMPLATE = app

# NOTE: Here you can change path to QXlsx sources

QXLSX_PARENTPATH=../../QXlsx/
QXLSX_HEADERPATH=$${QXLSX_PARENTPATH}header/ # should be path to QXlsx/header directory
QXLSX_SOURCEPATH=$${QXLSX_PARENTPATH}source/ # should be path to QXlsx/source directory

include($${QXLSX_PARENTPATH}QXlsx.pri)

DEFINES += QT_DEPRECATED_WARNINGS

SOURCES += \
main.cpp \
mainwindow.cpp \
sat_calc.cpp

HEADERS += \
mainwindow.h \
sat_calc.h

FORMS += \
mainwindow.ui

RESOURCES += \
test.qrc

#OTHER_FILES += \
#Chromatogram_example.txt


