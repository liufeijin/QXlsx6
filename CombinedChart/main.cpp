#include <QtGlobal>
#include <QCoreApplication>
#include <QtCore>
#include <QVariant>
#include <QDir>
#include <QDebug>

#include <iostream>
using namespace std;

// [0] include QXlsx headers
#include "xlsxdocument.h"
#include "xlsxchartsheet.h"
#include "xlsxcellrange.h"
#include "xlsxchart.h"
#include "xlsxrichstring.h"
#include "xlsxworkbook.h"
using namespace QXlsx;

int main(int argc, char *argv[])
{
    QCoreApplication app(argc, argv);

    QXlsx::Document xlsx;
    for (int i=1; i<10; ++i) {
        xlsx.write(1, i+1, "Pos " + QString::number(i));
    }

    xlsx.write(2, 1, "Set 1");
    xlsx.write(3, 1, "Set 2");

    for (int i=1; i<10; ++i) {
        xlsx.write(2, i+1, i*i*i);   //A2:A10
        xlsx.write(3, i+1, i*i);    //B2:B10
    }
    /**

    An example of a chart with series of different types.

    */

    Chart *chart1 = xlsx.insertChart(4, 3, QSize(600, 600));
    chart1->setLegend(Legend::Position::Right);
    chart1->setTitle("Combined chart");

    //The first series - bar series
    chart1->setType(Chart::Type::Bar);
    chart1->addSeries(CellRange(1,1,2,10), NULL, true, true, false);

    //The second series - line chart
    chart1->setType(Chart::Type::Line);
    chart1->addSeries(CellRange(1,1,1,10), CellRange(3,1,3,10), NULL, true);

    /**

    An example of a chart with series of different types. The first series is placed 
    on the left axis, the second one is on the right axis.

    */

    Chart *chart2 = xlsx.insertChart(4, 15, QSize(600, 600));
    chart2->setLegend(Legend::Position::Right);
    chart2->setTitle("Combined chart with three axes");

    // add axes
    auto &bottomAxis = chart2->addAxis(Axis::Type::Category, Axis::Position::Bottom, "bottom axis");
    auto &leftAxis = chart2->addAxis(Axis::Type::Value, Axis::Position::Left, "left axis");
    auto &rightAxis = chart2->addAxis(Axis::Type::Value, Axis::Position::Right, "right axis");
    //we need the 2nd bottom axis as Excel requires each subchart have its own subset of axes
    auto &bottom2Axis = chart2->addAxis(Axis::Type::Category, Axis::Position::Bottom);
    //hide the second bottom axis
    bottom2Axis.setVisible(false);
    //link axes together
    bottomAxis.setCrossAxis(&leftAxis);
    rightAxis.setCrossAxis(&bottom2Axis);

    chart2->title().textProperties().textShape = TextShapeType::textChevron;

    //The first series - bar series
    chart2->setType(Chart::Type::Bar);
    chart2->addSeries(CellRange(1,1,2,10), nullptr, true, true, false);
    chart2->setSeriesAxesIDs({bottomAxis.id(), leftAxis.id()});

    //The second series - line chart
    chart2->setType(Chart::Type::Line);
    chart2->addSeries(CellRange(1,1,1,10), CellRange(3,1,3,10), nullptr, true);
    chart2->setSeriesAxesIDs({bottom2Axis.id(), rightAxis.id()});

    //right now both left and right axes are plotted on the left. We need to properly
    //position the right axis.
    rightAxis.setCrossesType(Axis::CrossesType::Maximum); // Comment out this line to see what changes.

    xlsx.saveAs("combinedChart1.xlsx");

    Document xlsx2("combinedChart1.xlsx");
    if (xlsx2.load()) {
        xlsx2.saveAs("combinedChart2.xlsx");
    }

    return 0;
}
