// main.cpp

#include <QtGlobal>
#include <QCoreApplication>
#include <QtCore>
#include <QDebug>

#include <iostream>
using namespace std;

extern int chartsheet();

int main(int argc, char *argv[])
{
    QCoreApplication app(argc, argv);

    chartsheet();

    qDebug() << "**** end of main() ****";

    return 0;
}
