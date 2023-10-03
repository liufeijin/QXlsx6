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
    Chartsheet *sheet = xlsx.addChartsheet("Chart1");
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
        barChart->chartShape().fill().setType(FillFormat::FillType::PictureFill);
        barChart->chartShape().fill().setPicture(img);
        barChart->chartShape().fill().setPictureAlpha(70); //70% opacity
        barChart->plotAreaShape().setFill(FillFormat(FillFormat::FillType::NoFill));
    }

    //Saving result
    xlsx.saveAs("chartsheet1.xlsx");

    //Testing load result
    Document xlsx2("chartsheet1.xlsx");
    {
        sheet = static_cast<Chartsheet*>(xlsx2.sheet("Chart2"));
        barChart = sheet->chart();
        Q_ASSERT(barChart->chartShape().isValid() && barChart->chartShape().fill().isValid()
                 && !barChart->chartShape().fill().picture().isNull());

        //Testing remove blip fill
        barChart->chartShape().fill().setPicture({});
    }


    if (xlsx2.isLoaded())
        xlsx2.saveAs("chartsheet2.xlsx");

    //chartsheet2.xlsx shouldn't have image1.png inside

    return 0;
}
