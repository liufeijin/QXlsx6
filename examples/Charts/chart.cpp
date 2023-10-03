// chart.cpp

#include <QtGlobal>
#include <QtCore>
#include <QDebug>

#include "xlsxdocument.h"
#include "xlsxcellrange.h"
#include "xlsxchart.h"

QXLSX_USE_NAMESPACE

int chart()
{
    Document xlsx;
    auto sheet = xlsx.activeWorksheet();

    for (int i=1; i<10; ++i)
    {
        sheet->write(i, 1, i*i*i);   //A1:A9
        sheet->write(i, 2, i*i); //B1:B9
        sheet->write(i, 3, i*i-1); //C1:C9
    }

    Chart *pieChart = sheet->insertChart(3, 3, QSize(300, 300));
    pieChart->setType(Chart::Type::Pie);
    pieChart->addSeries(CellRange("A1:A9"));
    pieChart->addSeries(CellRange("B1:B9"));
    pieChart->addSeries(CellRange("C1:C9"));

    Chart *pie3DChart = sheet->insertChart(3, 9, QSize(300, 300));
    pie3DChart->setType(Chart::Type::Pie3D);
    pie3DChart->addSeries(CellRange("A1:C9"));

    Chart *barChart = sheet->insertChart(23, 3, QSize(300, 300));
    barChart->setType(Chart::Type::Bar);
    barChart->addSeries(CellRange("A1:C9"));

    Chart *bar3DChart = sheet->insertChart(23, 9, QSize(300, 300));
    bar3DChart->setType(Chart::Type::Bar3D);
    bar3DChart->addSeries(CellRange("A1:C9"));

    Chart *lineChart = sheet->insertChart(43, 3, QSize(300, 300));
    lineChart->setType(Chart::Type::Line);
    lineChart->addSeries(CellRange("A1:C9"));

    Chart *line3DChart = sheet->insertChart(43, 9, QSize(300, 300));
    line3DChart->setType(Chart::Type::Line3D);
    line3DChart->addSeries(CellRange("A1:C9"));

    Chart *areaChart = sheet->insertChart(63, 3, QSize(300, 300));
    areaChart->setType(Chart::Type::Area);
    areaChart->addSeries(CellRange("A1:C9"));

    Chart *area3DChart = sheet->insertChart(63, 9, QSize(300, 300));
    area3DChart->setType(Chart::Type::Area3D);
    area3DChart->addSeries(CellRange("A1:C9"));
    // }}

    Chart *scatterChart = sheet->insertChart(83, 3, QSize(300, 300));
    scatterChart->setType(Chart::Type::Scatter);
    //Will generate three lines.
    scatterChart->addSeries(CellRange("A1:A9"));
    scatterChart->addSeries(CellRange("B1:B9"));
    scatterChart->addSeries(CellRange("C1:C9"));

    Chart *scatterChart_2 = sheet->insertChart(83, 9, QSize(300, 300));
    scatterChart_2->setType(Chart::Type::Scatter);
    //Will generate two lines.
    scatterChart_2->addSeries(CellRange("A1:C9"));

    Chart *doughnutChart = sheet->insertChart(103, 3, QSize(300, 300));
    doughnutChart->setType(Chart::Type::Doughnut);
    doughnutChart->addSeries(CellRange("A1:C9"));

    //Testing copying of worksheets with charts
    xlsx.copySheet("Sheet1", "Sheet2");

    xlsx.saveAs("chart1.xlsx");

    Document xlsx2("chart1.xlsx");
    if ( xlsx2.isLoaded() )
    {
        xlsx2.saveAs("chart2.xlsx");
    }

    return 0;
}
