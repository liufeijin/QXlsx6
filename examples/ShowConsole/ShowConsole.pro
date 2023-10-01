# ShowConsole.pro
 
TARGET = ShowConsole
TEMPLATE = app

QT += core

CONFIG += console
CONFIG -= app_bundle

DEFINES += QT_DEPRECATED_WARNINGS

# Default rules for deployment.
qnx: target.path = /tmp/$${TARGET}/bin
else: unix:!android: target.path = /opt/$${TARGET}/bin
!isEmpty(target.path): INSTALLS += target

##########################################################################
# NOTE: Here you can change path to QXlsx sources

QXLSX_PARENTPATH=../../QXlsx/
QXLSX_HEADERPATH=$${QXLSX_PARENTPATH}header/ # should be path to QXlsx/header directory
QXLSX_SOURCEPATH=$${QXLSX_PARENTPATH}source/ # should be path to QXlsx/source directory

include($${QXLSX_PARENTPATH}QXlsx.pri)

# libfort (c)Seleznev Anton, MIT license
HEADERS += fort.hpp fort.h
SOURCES += fort.c
SOURCES += main.cpp
