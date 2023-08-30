// xlsxcellformula_p.h

#ifndef XLSXCELLFORMULA_P_H
#define XLSXCELLFORMULA_P_H

#include "xlsxglobal.h"
#include "xlsxcellformula.h"
#include "xlsxcellrange.h"

#include <QSharedData>
#include <QString>

namespace QXlsx {

class CellFormulaPrivate : public QSharedData
{
public:
    CellFormulaPrivate(const QString &formula, const CellRange &reference, CellFormula::Type type);
    CellFormulaPrivate(const CellFormulaPrivate &other);
    ~CellFormulaPrivate();

    QString formula; //formula contents
    CellFormula::Type type;
    CellRange reference;
    bool ca; //Calculate Cell
    int si;  //Shared group index
};

}

#endif // XLSXCELLFORMULA_P_H
