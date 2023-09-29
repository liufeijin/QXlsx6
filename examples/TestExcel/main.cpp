// main.cpp

#include <QtGlobal>
#include <QCoreApplication>
#include <QtCore>
#include <QDebug>

#include <iostream>
using namespace std;

extern int datavalidation();
extern int definename();
extern int documentproperty();
extern int extractdata();
extern int formula();
extern int hyperlink();
extern int image();
extern int richtext();
extern int rowcolumn();
extern int style();
extern int worksheetoperations();
extern int readStyle();
extern int pages();
extern int readextlist();
extern int autofilter();

int main(int argc, char *argv[])
{
    QCoreApplication app(argc, argv);

    qDebug() << "**** readStyle() ****";
    readStyle();


    qDebug() << "**** datavalidation() ****";
    datavalidation();
    qDebug() << "**** definename() ****";
    definename();
    qDebug() << "**** documentproperty() ****";
    
    documentproperty();
    qDebug() << "**** extractdata() ****";
    extractdata();
    qDebug() << "**** formula() ****";
    formula();
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
    qDebug() << "**** pages() ****";
    pages();
    qDebug() << "**** readextlist() ****";
    readextlist();
    qDebug() << "**** autofilter() ****";
    autofilter();

    qDebug() << "**** end of main() ****";

    return 0;
}
