// chartsheet.cpp

#include <QtGlobal>

#include "xlsxdocument.h"
#include "xlsxchartsheet.h"
#include "xlsxcellrange.h"
#include "xlsxchart.h"

QXLSX_USE_NAMESPACE

int chartsheet()
{
    //![0]
    Document xlsx;
    for (int i=1; i<10; ++i)
    {
        xlsx.write(i, 1, i*i);
    }
    //![0]

    //![1]
    xlsx.addSheet("Chart1", AbstractSheet::Type::Chartsheet);
    Chartsheet *sheet = static_cast<Chartsheet*>(xlsx.currentSheet());
    Chart *barChart = sheet->chart();
    barChart->setType(Chart::Type::Bar);
    barChart->addSeries(CellRange("A1:A9"), xlsx.sheet("Sheet1"));
    //![1]

    //![2]
    xlsx.saveAs("chartsheet1.xlsx");
    //![2]

    Document xlsx2("chartsheet1.xlsx");
    if ( xlsx2.load() )
    {
        xlsx2.saveAs("chartsheet2.xlsx");
    }

    return 0;
}
