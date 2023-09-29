// main.cpp

#include <QtGlobal>
#include <QCoreApplication>
#include <QtCore>
#include <QDebug>

#include <iostream>
using namespace std;

extern int pages();

int main(int argc, char *argv[])
{
    QCoreApplication app(argc, argv);

    qDebug() << "**** pages() ****";
    pages();

    qDebug() << "**** end of main() ****";

    return 0;
}
