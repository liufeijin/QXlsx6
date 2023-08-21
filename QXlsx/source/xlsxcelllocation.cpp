// xlsxcelllocation.cpp

#include <QtGlobal>
#include <QObject>
#include <QString>
#include <QVector>
#include <QList>

#include "xlsxglobal.h"
#include "xlsxcell.h"
#include "xlsxcelllocation.h"

namespace QXlsx {

CellLocation::CellLocation()
{
    col = -1;
    row = -1;

    cell.reset();
}

}
