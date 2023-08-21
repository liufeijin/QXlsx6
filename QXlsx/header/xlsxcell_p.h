// xlsxcell_p.h

#ifndef XLSXCELL_P_H
#define XLSXCELL_P_H

#include <QtGlobal>
#include <QObject>
#include <QList>

#include "xlsxglobal.h"
#include "xlsxcell.h"
#include "xlsxcellrange.h"
#include "xlsxrichstring.h"
#include "xlsxcellformula.h"

namespace QXlsx {

class CellPrivate
{
    Q_DECLARE_PUBLIC(Cell)
public:
    CellPrivate(Cell *p);
    CellPrivate(const CellPrivate * const cp);
public:
    Worksheet *parent;
    Cell *q_ptr;
public:
    Cell::Type cellType;
    QVariant value;

    CellFormula formula;
    Format format;

    RichString richString;

    qint32 styleNumber;
};

}

#endif // XLSXCELL_P_H
