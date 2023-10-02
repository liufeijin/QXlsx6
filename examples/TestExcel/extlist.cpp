// extlist.cpp

#include <QtGlobal>
#include <QtCore>
#include <QDebug>

#include "xlsxdocument.h"
#include "xlsxchart.h"

int readextlist()
{
    QFile::copy(":/extlist.xlsx", "extlist1.xlsx");
    QXlsx::Document xlsx(":/extlist.xlsx");
    xlsx.saveAs("extlist2.xlsx");

    return 0;
}
