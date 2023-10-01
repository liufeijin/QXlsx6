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

    // Copying source file as is.
    QFile::copy(":/chartsheet1.xlsx", "chartsheet1.xlsx");

    // Reading source file with QXlsx
    Document xlsx2(":/chartsheet1.xlsx");
    if (xlsx2.isLoaded()) {
        xlsx2.saveAs("chartsheet2.xlsx");
    }

    return 0;
}
