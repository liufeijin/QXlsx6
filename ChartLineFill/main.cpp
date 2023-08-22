#include <QtGlobal>
#include <QCoreApplication>
#include <QtCore>
#include <QVariant>
#include <QDir>
#include <QDebug>

#include <iostream>
using namespace std;

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

    for (int row = 1; row < 13; ++row) {
        xlsx.write(row+1, 1, "Set "+QString::number(row));

        for (int i=1; i<10; ++i) {
            xlsx.write(row+1, i+1, (i % 2)+row);
        }
    }

    /**
    An example of chart lines and filling.
    */

    Chart *chart1 = xlsx.insertChart(12, 0, QSize(600, 600));
    chart1->setLegend(Legend::Position::Right);
    chart1->setTitle("Line stroke style");

    chart1->setType(Chart::Type::Line);
    chart1->addSeries(CellRange(1,1,11,10), NULL, true, true, false);
    for (int i=0; i<chart1->seriesCount(); ++i) {
        chart1->series(i)->setMarker(MarkerFormat::MarkerType::None);
        chart1->series(i)->line().setStrokeType(static_cast<LineFormat::StrokeType>(i));
        QString s; LineFormat::toString(static_cast<LineFormat::StrokeType>(i), s);
        chart1->series(i)->setName(s);
    }

    Chart *chart2 = xlsx.insertChart(12, 10, QSize(600, 600));
    chart2->setLegend(Legend::Position::Right);
    chart2->setTitle("custom dash pattern");

    chart2->setType(Chart::Type::Line);
    chart2->addSeries(CellRange(1,1,12,10), NULL, true, true, false);
    //dashes and spaces lengths are in %
    QVector<double> customDashPattern {300,100,100,100,300,100,100,100,100,100,
                                       100,100,300,100,100,100,100,100,100,100,
                                       300,100,300,100,300,100,300,100,300,100,
                                       300,100};
    for (int i=0; i<5; ++i) {
        chart2->series(i)->setMarker(MarkerFormat::MarkerType::None);
        chart2->series(i)->line().setDashPattern(customDashPattern);
        chart2->series(i)->line().setCompoundLineType(static_cast<LineFormat::CompoundLineType>(i));
        chart2->series(i)->line().setWidth(5.0);
        QString s; LineFormat::toString(static_cast<LineFormat::CompoundLineType>(i), s);
        chart2->series(i)->setName(s);
    }
    for (int i=5; i<8; ++i) {
        chart2->series(i)->setMarker(MarkerFormat::MarkerType::None);
        chart2->series(i)->line().setDashPattern(customDashPattern);
        chart2->series(i)->line().setLineCap(static_cast<LineFormat::LineCap>(i-5));
        chart2->series(i)->line().setWidth(5.0);
        QString s; LineFormat::toString(static_cast<LineFormat::LineCap>(i-5), s);
        chart2->series(i)->setName(s);
    }
    for (int i=8; i<11; ++i) {
        chart2->series(i)->setMarker(MarkerFormat::MarkerType::None);
        chart2->series(i)->line().setDashPattern(customDashPattern);
        chart2->series(i)->line().setLineJoin(static_cast<LineFormat::LineJoin>(i-8));
        chart2->series(i)->line().setWidth(5.0);
        QString s; LineFormat::toString(static_cast<LineFormat::LineJoin>(i-8), s);
        chart2->series(i)->setName(s);
    }

    Chart *chart3 = xlsx.insertChart(12, 20, QSize(600, 600));
    chart3->setLegend(Legend::Position::Right);
    chart3->setTitle("Line ending");

    chart3->setType(Chart::Type::Line);
    chart3->addSeries(CellRange(1,1,13,10), NULL, true, true, false);
    for (int i=0; i<9; ++i) {
        chart3->series(i)->setMarker(MarkerFormat::MarkerType::None);
        chart3->series(i)->line().setLineStartType(LineFormat::LineEndType::Arrow);
        chart3->series(i)->line().setLineEndType(LineFormat::LineEndType::Triangle);
    }
    for (int i=9; i<12; ++i) {
        chart3->series(i)->setMarker(MarkerFormat::MarkerType::None);
        chart3->series(i)->line().setLineStartType(LineFormat::LineEndType::Oval);
        chart3->series(i)->line().setLineEndType(LineFormat::LineEndType::Stealth);
    }

    for (int i=0; i<3; ++i) {
        for (int j=0; j<3; ++j) {
            chart3->series(i*3+j)->line().setLineStartWidth(static_cast<LineFormat::LineEndSize>(i));
            chart3->series(i*3+j)->line().setLineStartLength(static_cast<LineFormat::LineEndSize>(j));
            chart3->series(i*3+j)->line().setLineEndWidth(static_cast<LineFormat::LineEndSize>(i));
            chart3->series(i*3+j)->line().setLineEndLength(static_cast<LineFormat::LineEndSize>(j));
        }
    }
    for (int j=0; j<3; ++j) {
        chart3->series(j+9)->line().setLineStartWidth(LineFormat::LineEndSize::Large);
        chart3->series(j+9)->line().setLineStartLength(static_cast<LineFormat::LineEndSize>(j));
        chart3->series(j+9)->line().setLineEndWidth(LineFormat::LineEndSize::Large);
        chart3->series(j+9)->line().setLineEndLength(static_cast<LineFormat::LineEndSize>(j));
    }

    Chart *chart4 = xlsx.insertChart(44, 0, QSize(600, 300));
    chart4->setType(Chart::Type::Bar); //required
    chart4->addDefaultAxes(); //required

    FillFormat f(FillFormat::FillType::GradientFill);
    //the gradient fill is applied to the range [30%..70%] of the gradient path
    //the first color is red, the last color is blue, so the gradient is from red to blue
    f.addGradientStop(30, Color("red")); //first 30% of the shape will be filled with red
    f.addGradientStop(70, Color("blue")); //last 30% of the shape will be filled with blue
    //set the linear gradient angle of 45 degrees
    f.setLinearShadeAngle(45.0);
    //make this angle be actually from the top left corner to the bottom right corner of the shape
    f.setLinearShadeScaled(true);
    chart4->chartShape().setFill(f);

    Chart *chart5 = xlsx.insertChart(44, 10, QSize(600, 300));
    chart5->setType(Chart::Type::Bar); //required
    chart5->addDefaultAxes(); //required

    FillFormat f1(FillFormat::FillType::GradientFill);
    //the gradient fill is applied to the range [0%..100%] of the gradient path
    //the first color is red, the last color is blue, so the gradient is from red to blue
    f1.addGradientStop(0, Color("red"));
    f1.addGradientStop(100, Color("blue"));
    //set the linear gradient angle of 45 degrees
    f1.setLinearShadeAngle(45.0);
    //make this angle be actually from the top left corner to the bottom right corner of the shape
    f1.setLinearShadeScaled(true);
    //shrink the gradient by applying the edge margins of 20% at the top and bottom,
    //and 45% at left and right
    RelativeRect rr(45,20,45,20); f1.setTileRect(rr);
    //THis parameter is actually ignored by Excel.
    f1.setTileFlipMode(FillFormat::TileFlipMode::XY);
    chart5->chartShape().setFill(f1);

    xlsx.saveAs("LineAndFill1.xlsx");

    Document xlsx2("LineAndFill1.xlsx");
    if (xlsx2.load()) {
        xlsx2.saveAs("LineAndFill2.xlsx");
    }

    return 0;
}
