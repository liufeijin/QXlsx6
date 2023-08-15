// chartextended.cpp

#include <QtGlobal>
#include <QtCore>
#include <QDebug>

#include "xlsxdocument.h"
#include "xlsxcellrange.h"
#include "xlsxchart.h"

QXLSX_USE_NAMESPACE

/*
 * Test Chart with title, gridlines, legends, multiple headers
 *
 */

int chartExtended()
{
    Document xlsx;
    for (int i=1; i<10; ++i)
    {
        xlsx.write(1, i+1, "Pos " + QString::number(i));
    }

    xlsx.write(2, 1, "Set 1");
    xlsx.write(3, 1, "Set 2");

    for (int i=1; i<10; ++i)
    {
        xlsx.write(2, i+1, i*i*i);   //A2:A10
        xlsx.write(3, i+1, i*i);    //B2:B10
    }


    Chart *barChart1 = xlsx.insertChart(4, 3, QSize(300, 300));
    barChart1->setType(Chart::Type::Bar);
    barChart1->setLegend(Legend::Position::Right);
    barChart1->setTitle("Test1");
    // Messreihen
    barChart1->addSeries(CellRange(1,1,3,10), NULL, true, true);

    Chart *barChart2 = xlsx.insertChart(4, 9, QSize(300, 300));
    barChart2->setType(Chart::Type::Bar);
    barChart2->setLegend(Legend::Position::Right);
    barChart2->setTitle("Test2");
    barChart2->addDefaultAxes();
    // Messreihen
    barChart2->addSeries(CellRange(1,1,3,10), NULL, true, true);
    //axes by index
    barChart2->axis(0)->majorGridLines().setLine(LineFormat(FillFormat::FillType::SolidFill, 0.15, QColor(Qt::black)));
    barChart2->axis(1)->majorGridLines().setLine(LineFormat(FillFormat::FillType::SolidFill, 0.15, QColor(Qt::black)));



    Chart *barChart3 = xlsx.insertChart(24, 3, QSize(300, 300));
    barChart3->setType(Chart::Type::Bar);
    barChart3->setLegend(Legend::Position::Left);
    barChart3->setTitle("Test3");
    // Messreihen
    barChart3->addSeries(CellRange(1,1,3,10));
    //axes by position
    barChart3->axis(Axis::Position::Bottom)->majorGridLines().setLine(LineFormat(FillFormat::FillType::SolidFill, 0.15, QColor(Qt::red)));
    barChart3->axis(Axis::Position::Left)->majorGridLines().setLine(LineFormat(FillFormat::FillType::SolidFill, 0.15, QColor(Qt::black)));

    Chart *barChart4 = xlsx.insertChart(24, 9, QSize(300, 300));
    barChart4->setType(Chart::Type::Bar);
    barChart4->setLegend(Legend::Position::Top);
    barChart4->setTitle("Test4");
    // Messreihen
    barChart4->addSeries(CellRange(1,1,3,10));
    //axes by type
    barChart4->axis(Axis::Type::Category)->majorGridLines().setLine(LineFormat(FillFormat::FillType::SolidFill, 0.15, QColor(Qt::red)));
    barChart4->axis(Axis::Type::Value)->majorGridLines().setLine(LineFormat(FillFormat::FillType::SolidFill, 0.15, QColor(Qt::black)));
    if (auto ax = barChart4->axis(Axis::Type::Series)) {
        qDebug() << "Found a series axis, none should be found";
    }

    Chart *barChart5 = xlsx.insertChart(44, 9, QSize(300, 300));
    barChart5->setType(Chart::Type::Bar);
    barChart5->setLegend(Legend::Position::Bottom);
    barChart5->setTitle("Test5");
    // Messreihen
    barChart5->addSeries(CellRange(1,1,3,10));


    //![2]
    xlsx.saveAs("chartExtended1.xlsx");
    //![2]

    Document xlsx2("chartExtended1.xlsx");
    if ( xlsx2.load() )
    {
        xlsx2.saveAs("chartExtended2.xlsx");
    }

    return 0;
}
