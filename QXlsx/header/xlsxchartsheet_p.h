// xlsxchartsheet_p.h

#ifndef XLSXCHARTSHEET_P_H
#define XLSXCHARTSHEET_P_H

#include <QtGlobal>

#include "xlsxglobal.h"
#include "xlsxchartsheet.h"
#include "xlsxabstractsheet_p.h"

namespace QXlsx {

class ChartsheetPrivate : public AbstractSheetPrivate
{
    Q_DECLARE_PUBLIC(Chartsheet)
public:
    ChartsheetPrivate(Chartsheet *p, Chartsheet::CreateFlag flag);
    ~ChartsheetPrivate();

    Chart *chart;
};

}
#endif // XLSXCHARTSHEET_P_H
