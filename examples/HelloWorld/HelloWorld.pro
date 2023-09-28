# HelloWorld.pro
 
TARGET = HelloWorld
TEMPLATE = app

QT += core

CONFIG += console
CONFIG -= app_bundle

DEFINES += QT_DEPRECATED_WARNINGS

##########################################################################
# NOTE: Here you can change path to QXlsx sources

#  QXLSX_PARENTPATH=./
#  QXLSX_HEADERPATH=./header/ # should be path to QXlsx/header directory
#  QXLSX_SOURCEPATH=./source/ # should be path to QXlsx/source directory

include(../../QXlsx/QXlsx.pri)

SOURCES += main.cpp
