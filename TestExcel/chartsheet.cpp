// chartsheet.cpp

#include <QtGlobal>

#include "xlsxdocument.h"
#include "xlsxchartsheet.h"
#include "xlsxcellrange.h"
#include "xlsxchart.h"

QXLSX_USE_NAMESPACE

int chartsheet()
{
    Document xlsx;
    for (int i=1; i<10; ++i)
        xlsx.write(i, 1, i*i);


    //Add the first chartsheet
    xlsx.addSheet("Chart1", AbstractSheet::Type::Chartsheet);
    Chartsheet *sheet = static_cast<Chartsheet*>(xlsx.currentSheet());
    Chart *barChart = sheet->chart();
    barChart->setType(Chart::Type::Bar);
    barChart->addSeries(CellRange("A1:A9"), xlsx.sheet("Sheet1"));
    barChart->setTitle("Chart #1");

    //Testing the copy of the 1st chartsheet
    xlsx.copySheet("Chart1","Chart2");
    sheet = static_cast<Chartsheet*>(xlsx.sheet("Chart2"));
    barChart = sheet->chart();
    barChart->setTitle("Chart #2");

    //Testing set blip fill
    QImage img;
    if (img.load(":/background1.jpg")) {
        barChart->chartShape().fill().setType(FillFormat::FillType::BlipFill);
        barChart->chartShape().fill().setPicture(img);
    }

    //Saving result
    xlsx.saveAs("chartsheet1.xlsx");

    //Testing load result
    Document xlsx2("chartsheet1.xlsx");
    //TODO: fix loading blip fills

    if (xlsx2.load())
        xlsx2.saveAs("chartsheet2.xlsx");

    return 0;
}
