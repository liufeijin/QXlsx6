// main.cpp

#include <QtGlobal>
#include <QCoreApplication>
#include <QtCore>
#include <QDebug>

#include <iostream>
using namespace std;

extern int chart();
extern int chartExtended();
extern int barChart();
extern int chartlinefill();

int main(int argc, char *argv[])
{
    QCoreApplication app(argc, argv);

    qDebug() << "**** chart() ****";
    chart();
    qDebug() << "**** chartExtended() ****";
    chartExtended();
    qDebug() << "**** barChart() ****";
    barChart();
    qDebug() << "**** chartlinefill() ****";
    chartlinefill();

    qDebug() << "**** end of main() ****";

    return 0;
}
