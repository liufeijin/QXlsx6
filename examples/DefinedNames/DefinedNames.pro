# DefinedNames.pro

TARGET = DefinedNames
TEMPLATE = app

QT += core
QT += gui

CONFIG += console
CONFIG -= app_bundle

# NOTE: Here you can change path to QXlsx sources

QXLSX_PARENTPATH=../../QXlsx/
QXLSX_HEADERPATH=$${QXLSX_PARENTPATH}header/ # should be path to QXlsx/header directory
QXLSX_SOURCEPATH=$${QXLSX_PARENTPATH}source/ # should be path to QXlsx/source directory

include($${QXLSX_PARENTPATH}QXlsx.pri)

##########################################################################
# The following define makes your compiler emit warnings if you use
# any feature of Qt which as been marked deprecated (the exact warnings
# depend on your compiler). Please consult the documentation of the
# deprecated API in order to know how to port your code away from it.
DEFINES += QT_DEPRECATED_WARNINGS

##########################################################################
# You can also make your code fail to compile if you use deprecated APIs.
# In order to do so, uncomment the following line.
# You can also select to disable deprecated APIs only up to a certain
# version of Qt.
#DEFINES += QT_DISABLE_DEPRECATED_BEFORE=0x060000   
# disables all the APIs deprecated before Qt 6.0.0

##########################################################################
# source code

SOURCES += definedNames.cpp
