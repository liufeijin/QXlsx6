#include <QtGlobal>
#include <QCoreApplication>
#include <QtCore>
#include <QVariant>
#include <QDir>
#include <QDebug>

#include <iostream>
using namespace std;

#include "xlsxdocument.h"
using namespace QXlsx;

int main(int argc, char *argv[])
{
    QCoreApplication app(argc, argv);

    Document xlsx2("chartsheet1.xlsx");
    if (xlsx2.load()) {
        xlsx2.saveAs("chartsheet2.xlsx");
    }

    return 0;
}
