// barchart.cpp

#include <QtGlobal>
#include <QtCore>
#include <QDebug>

#include "xlsxdocument.h"
#include "xlsxcellrange.h"
#include "xlsxchart.h"

QXLSX_USE_NAMESPACE

/*
 * Test bar charts
 *
 */

int barChart()
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

    Chart *barChart11 = xlsx.insertChart(5, 4, QSize(300, 300));
    barChart11->setType(Chart::Type::Bar);
    barChart11->setLegend(Legend::Position::Right);
    barChart11->setTitle("Row-based, no options");
    barChart11->addSeries(CellRange(1,1,3,10), nullptr, true, true, false);

    //demonstrates how to use Coordinate to set size in pixels
    Chart *barChart12 = xlsx.insertChart(5, 10, Coordinate("300px"), Coordinate("300px"));
    barChart12->setType(Chart::Type::Bar);
    barChart12->setLegend(Legend::Position::Right);
    barChart12->setTitle("Row-based, Clustered grouping");
    barChart12->addSeries(CellRange(1,1,3,10), nullptr, true, true, false);
    barChart12->setBarGrouping(Chart::BarGrouping::Clustered);

    //demonstrates how to use Coordinate to set size in pixels
    Chart *barChart13 = xlsx.insertChart(5, 16, Coordinate("80mm"), Coordinate("80mm"));
    barChart13->setType(Chart::Type::Bar);
    barChart13->setLegend(Legend::Position::Right);
    barChart13->setTitle("Row-based, Stacked grouping");
    barChart13->addSeries(CellRange(1,1,3,10), nullptr, true, true, false);
    barChart13->setBarGrouping(Chart::BarGrouping::Stacked);

    Chart *barChart14 = xlsx.insertChart(5, 22, QSize(300, 300));
    barChart14->setType(Chart::Type::Bar);
    barChart14->setLegend(Legend::Position::Right);
    barChart14->setTitle("Row-based, PercentStacked grouping");
    barChart14->addSeries(CellRange(1,1,3,10), nullptr, true, true, false);
    barChart14->setBarGrouping(Chart::BarGrouping::PercentStacked);

    Chart *barChart21 = xlsx.insertChart(25, 4, QSize(300, 300));
    barChart21->setType(Chart::Type::Bar);
    barChart21->setLegend(Legend::Position::Right);
    barChart21->setTitle("Column-based, no options");
    barChart21->addSeries(CellRange(1,1,3,10), nullptr, true, true, true);

    Chart *barChart22 = xlsx.insertChart(25, 10, QSize(300, 300));
    barChart22->setType(Chart::Type::Bar);
    barChart22->setLegend(Legend::Position::Right);
    barChart22->setTitle("Column-based, Clustered grouping");
    barChart22->addSeries(CellRange(1,1,3,10), nullptr, true, true, true);
    barChart22->setBarGrouping(Chart::BarGrouping::Clustered);

    Chart *barChart23 = xlsx.insertChart(25, 16, QSize(300, 300));
    barChart23->setType(Chart::Type::Bar);
    barChart23->setLegend(Legend::Position::Right);
    barChart23->setTitle("Column-based, Stacked grouping");
    barChart23->addSeries(CellRange(1,1,3,10), nullptr, true, true, true);
    barChart23->setBarGrouping(Chart::BarGrouping::Stacked);

    Chart *barChart24 = xlsx.insertChart(25, 22, QSize(300, 300));
    barChart24->setType(Chart::Type::Bar);
    barChart24->setLegend(Legend::Position::Right);
    barChart24->setTitle("Column-based, PercentStacked grouping");
    barChart24->addSeries(CellRange(1,1,3,10), nullptr, true, true, true);
    barChart24->setBarGrouping(Chart::BarGrouping::PercentStacked);

    Chart *barChart31 = xlsx.insertChart(45, 4, QSize(300, 300));
    barChart31->setType(Chart::Type::Bar);
    barChart31->setLegend(Legend::Position::Right);
    barChart31->setTitle("Row-based, default bar direction");
    barChart31->addSeries(CellRange(1,1,3,10), nullptr, true, true, false);

    Chart *barChart32 = xlsx.insertChart(45, 10, QSize(300, 300));
    barChart32->setType(Chart::Type::Bar);
    barChart32->setLegend(Legend::Position::Right);
    barChart32->setTitle("Row-based, horizontal");
    barChart32->addSeries(CellRange(1,1,3,10), nullptr, true, true, false);
    barChart32->setBarDirection(Chart::BarDirection::Bar);

    Chart *barChart33 = xlsx.insertChart(45, 16, QSize(300, 300));
    barChart33->setType(Chart::Type::Bar);
    barChart33->setLegend(Legend::Position::Right);
    barChart33->setTitle("Row-based, no gap");
    barChart33->addSeries(CellRange(1,1,3,10), nullptr, true, true, false);
    barChart33->setGapWidth(0);

    Chart *barChart34 = xlsx.insertChart(45, 22, QSize(300, 300));
    barChart34->setType(Chart::Type::Bar);
    barChart34->setLegend(Legend::Position::Right);
    barChart34->setTitle("Row-based, 100% gap");
    barChart34->addSeries(CellRange(1,1,3,10), nullptr, true, true, false);
    barChart34->setGapWidth(100);

    //3D charts
    Chart *barChart41 = xlsx.insertChart(65, 4, QSize(300, 300));
    barChart41->setType(Chart::Type::Bar3D);
    barChart41->setLegend(Legend::Position::Right);
    barChart41->setTitle("3D bar");
    barChart41->addSeries(CellRange(1,1,3,10), nullptr, true, true, false);

    Chart *barChart42 = xlsx.insertChart(65, 10, QSize(300, 300));
    barChart42->setType(Chart::Type::Bar3D);
    barChart42->setLegend(Legend::Position::Right);
    barChart42->setTitle("3D bar, horizontal");
    barChart42->addSeries(CellRange(1,1,3,10), nullptr, true, true, false);
    barChart42->setBarDirection(Chart::BarDirection::Bar);

    Chart *barChart43 = xlsx.insertChart(65, 16, QSize(300, 300));
    barChart43->setType(Chart::Type::Bar3D);
    barChart43->setLegend(Legend::Position::Right);
    barChart43->setTitle("3D bar, no gap");
    barChart43->addSeries(CellRange(1,1,3,10), nullptr, true, true, false);
    barChart43->setGapWidth(0);

    Chart *barChart44 = xlsx.insertChart(65, 221, QSize(300, 300));
    barChart44->setType(Chart::Type::Bar3D);
    barChart44->setLegend(Legend::Position::Right);
    barChart44->setTitle("3D bar, no gap depth");
    barChart44->addSeries(CellRange(1,1,3,10), nullptr, true, true, false);
    barChart44->setGapDepth(0);

    Chart *barChart51 = xlsx.insertChart(85, 4, QSize(300, 300));
    barChart51->setType(Chart::Type::Bar3D);
    barChart51->setLegend(Legend::Position::Right);
    barChart51->setTitle("3D bar, standard shape");
    barChart51->addSeries(CellRange(1,1,3,10), nullptr, true, true, false);

    Chart *barChart52 = xlsx.insertChart(85, 10, QSize(300, 300));
    barChart52->setType(Chart::Type::Bar3D);
    barChart52->setLegend(Legend::Position::Right);
    barChart52->setTitle("3D bar, cone shape");
    barChart52->addSeries(CellRange(1,1,3,10), nullptr, true, true, false);
    barChart52->setBarShape(Series::BarShape::Cone);

    Chart *barChart53 = xlsx.insertChart(85, 16, QSize(300, 300));
    barChart53->setType(Chart::Type::Bar3D);
    barChart53->setLegend(Legend::Position::Right);
    barChart53->setTitle("3D bar, pyramid shape");
    barChart53->addSeries(CellRange(1,1,3,10), nullptr, true, true, false);
    barChart53->setBarShape(Series::BarShape::Pyramid);

    Chart *barChart54 = xlsx.insertChart(85, 22, QSize(300, 300));
    barChart54->setType(Chart::Type::Bar3D);
    barChart54->setLegend(Legend::Position::Right);
    barChart54->setTitle("3D bar, only 1st series cylinder");
    barChart54->addSeries(CellRange(1,1,3,10), nullptr, true, true, false);
    barChart54->series(0)->setBarShape(Series::BarShape::Cylinder);

    //![2]
    xlsx.saveAs("barCharts1.xlsx");
    //![2]

    Document xlsx2("barCharts1.xlsx");
    if ( xlsx2.isLoaded() )
    {
        xlsx2.saveAs("barCharts2.xlsx");
    }

    return 0;
}
