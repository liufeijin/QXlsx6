# WebServer.pro
 
TARGET = WebServer
TEMPLATE = app

QT += core
QT += network
QT -= gui

CONFIG += console
CONFIG -= app_bundle

macx {
    QMAKE_CXXFLAGS += -stdlib=libc++
}

# NOTE: Here you can change path to QXlsx sources

QXLSX_PARENTPATH=../../QXlsx/
QXLSX_HEADERPATH=$${QXLSX_PARENTPATH}header/ # should be path to QXlsx/header directory
QXLSX_SOURCEPATH=$${QXLSX_PARENTPATH}source/ # should be path to QXlsx/source directory

include($${QXLSX_PARENTPATH}QXlsx.pri)

# source code

RESOURCES += \
ws.qrc

HEADERS += \
recurse.hpp \
request.hpp \
response.hpp \
context.hpp

INCLUDEPATH += .

SOURCES += \
main.cpp


