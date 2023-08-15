// extlist.cpp

#include <QtGlobal>
#include <QtCore>
#include <QDebug>

#include "xlsxdocument.h"
#include "xlsxchart.h"

int readextlist()
{
    //![0]
    QXlsx::Document xlsx("extlist.xlsx");
    //![0]

    return 0;
}
