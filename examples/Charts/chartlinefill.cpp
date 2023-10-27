#include <QtGlobal>
#include <QCoreApplication>
#include <QtCore>
#include <QVariant>
#include <QDir>
#include <QDebug>

#include <iostream>
using namespace std;

#include "xlsxdocument.h"
#include "xlsxcellrange.h"
#include "xlsxchart.h"
using namespace QXlsx;

int chartlinefill()
{
    QXlsx::Document xlsx;
    auto sheet = xlsx.activeWorksheet();

    for (int i=1; i<10; ++i) {
        sheet->write(1, i+1, "Pos " + QString::number(i));
    }

    for (int row = 1; row < 13; ++row) {
        sheet->write(row+1, 1, "Set "+QString::number(row));

        for (int i=1; i<10; ++i) {
            sheet->write(row+1, i+1, (i % 2)+row);
        }
    }

    /**
    An example of chart lines and filling.
    */

    Chart *chart1 = sheet->insertChart(13, 1, QSize(600, 600));
    chart1->setLegend(Legend::Position::Right);
    chart1->setTitle("Line stroke style");

    chart1->addSubchart(Chart::Type::Line);
    chart1->addSeries(CellRange(1,1,11,10), NULL, true, true, false);
    for (int i=0; i<chart1->seriesCount(); ++i) {
        chart1->series(i)->setMarker(MarkerFormat::MarkerType::None);
        chart1->series(i)->line().setStrokeType(static_cast<LineFormat::StrokeType>(i));
        chart1->series(i)->setName(LineFormat::toString(static_cast<LineFormat::StrokeType>(i)));
    }

    Chart *chart2 = sheet->insertChart(13, 11, QSize(600, 600));
    chart2->setLegend(Legend::Position::Right);
    chart2->setTitle("custom dash pattern");

    chart2->addSubchart(Chart::Type::Line);
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
        chart2->series(i)->setName(LineFormat::toString(static_cast<LineFormat::CompoundLineType>(i)));
    }
    for (int i=5; i<8; ++i) {
        chart2->series(i)->setMarker(MarkerFormat::MarkerType::None);
        chart2->series(i)->line().setDashPattern(customDashPattern);
        chart2->series(i)->line().setLineCap(static_cast<LineFormat::LineCap>(i-5));
        chart2->series(i)->line().setWidth(5.0);
        chart2->series(i)->setName(LineFormat::toString(static_cast<LineFormat::LineCap>(i-5)));
    }
    for (int i=8; i<11; ++i) {
        chart2->series(i)->setMarker(MarkerFormat::MarkerType::None);
        chart2->series(i)->line().setDashPattern(customDashPattern);
        chart2->series(i)->line().setLineJoin(static_cast<LineFormat::LineJoin>(i-8));
        chart2->series(i)->line().setWidth(5.0);
        chart2->series(i)->setName(LineFormat::toString(static_cast<LineFormat::LineJoin>(i-8)));
    }

    Chart *chart3 = sheet->insertChart(13, 21, QSize(600, 600));
    chart3->setLegend(Legend::Position::Right);
    chart3->setTitle("Line ending");

    chart3->addSubchart(Chart::Type::Line);
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

    Chart *chart4 = sheet->insertChart(45, 1, QSize(600, 300));
    chart4->addSubchart(Chart::Type::Bar); //required
    xlsx.write(44,1,"Linear gradient");

    FillFormat f(FillFormat::FillType::GradientFill);
    //the gradient fill is applied to the range [30%..70%] of the gradient path
    //the first color is red, the last color is blue, so the gradient is from red to blue
    f.addGradientStop(30, "red"); //first 30% of the shape will be filled with red
    f.addGradientStop(70, "blue"); //last 30% of the shape will be filled with blue
    //set the linear gradient angle of 45 degrees
    f.setLinearShadeAngle(45.0);
    //make this angle be actually from the top left corner to the bottom right corner of the shape
    f.setLinearShadeScaled(true);
    chart4->chartShape().setFill(f);


    //Chart 5 will have path gradient fill
    Chart *chart5 = sheet->insertChart(45, 11, QSize(300, 300));
    chart5->addSubchart(Chart::Type::Bar); //required
    xlsx.write(44,11,"Circle gradient");

    //the gradient fill is applied to the range [0%..100%] of the gradient path
    //the first color is red, the last color is blue, so the gradient is from red to blue
    FillFormat f1({{0, "red"},{100, "blue"}}, FillFormat::PathType::Circle);
    //expand the gradient by applying the edge margin of -50% at bottom
    f1.setTileRect(RelativeRect(0,0,0,-50));
    //move the focus point of the gradient to the center of the bottom edge.
    f1.setPathRect(RelativeRect(50,100,50,0));
    //This parameter is actually ignored by Excel.
    chart5->chartShape().setFill(f1);

#if (QT_VERSION >= QT_VERSION_CHECK(5, 12, 0))
    //Chart 6 will use QGradient preset
    Chart *chart6 = sheet->insertChart(45, 21, QSize(300, 300));
    chart6->addSubchart(Chart::Type::Bar); //required
    xlsx.write(44,21,"Preset gradient");

    FillFormat f2((QGradient(QGradient::KindSteel)));
    chart6->chartShape().setFill(f2);
#endif

    xlsx.saveAs("LineAndFill1.xlsx");

    Document xlsx2("LineAndFill1.xlsx");
    if (xlsx2.isLoaded()) {
        xlsx2.saveAs("LineAndFill2.xlsx");
    }

    return 0;
}
