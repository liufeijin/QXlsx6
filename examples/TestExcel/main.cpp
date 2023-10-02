// main.cpp

#include <QtGlobal>
#include <QCoreApplication>
#include <QtCore>
#include <QDebug>

#include <iostream>
using namespace std;

extern int datavalidation();
extern int documentproperty();
extern int hyperlink();
extern int image();
extern int richtext();
extern int rowcolumn();
extern int style();
extern int worksheetoperations();
extern int readStyle();
extern int readextlist();

int main()
{
    qDebug() << "**** readStyle() ****";
    readStyle();
    qDebug() << "**** datavalidation() ****";
    datavalidation();
    qDebug() << "**** documentproperty() ****";
    documentproperty();
    qDebug() << "**** hyperlink() ****";
    hyperlink();
    qDebug() << "**** image() ****";
    image();
    qDebug() << "**** richtext() ****";
    richtext();
    qDebug() << "**** rowcolumn() ****";
    rowcolumn();
    qDebug() << "**** style() ****";
    style();
    qDebug() << "**** worksheetoperations() ****";
    worksheetoperations();
    qDebug() << "**** readextlist() ****";
    readextlist();
    qDebug() << "**** end of main() ****";

    return 0;
}
