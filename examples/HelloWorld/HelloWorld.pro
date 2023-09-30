# HelloWorld.pro
 
TARGET = HelloWorld
TEMPLATE = app

QT += core

CONFIG += console
CONFIG -= app_bundle

DEFINES += QT_DEPRECATED_WARNINGS

##########################################################################
# NOTE: Here you can change path to QXlsx sources

QXLSX_PARENTPATH=../../QXlsx/
QXLSX_HEADERPATH=$${QXLSX_PARENTPATH}header/ # should be path to QXlsx/header directory
QXLSX_SOURCEPATH=$${QXLSX_PARENTPATH}source/ # should be path to QXlsx/source directory

include($${QXLSX_PARENTPATH}QXlsx.pri)

SOURCES += main.cpp
