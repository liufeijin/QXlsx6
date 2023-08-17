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
        xlsx.write(1, i+1, i);
    }

    xlsx.write(2, 1, "Set 1");
    xlsx.write(3, 1, "Set 2");

    for (int i=1; i<10; ++i)
    {
        xlsx.write(2, i+1, i*i*i);   //A2:A10
        xlsx.write(3, i+1, i*i);    //B2:B10
    }
    xlsx.write(4,1,"Referenced");
    xlsx.write(5,1,"title");

    /// Gridlines
    // no gridlines
    Chart *chart11 = xlsx.insertChart(4, 3, QSize(300, 300));
    chart11->setType(Chart::Type::Scatter);
    chart11->setLegend(Legend::Position::Right);
    chart11->setTitle("No gridlines");
    chart11->addSeries(CellRange(1,1,3,10), NULL, true, true, false);

    // standarg gridlines
    Chart *chart12 = xlsx.insertChart(4, 9, QSize(300, 300));
    chart12->setType(Chart::Type::Scatter);
    chart12->setLegend(Legend::Position::Right);
    chart12->setTitle("Standard major gridlines");
    chart12->addSeries(CellRange(1,1,3,10), NULL, true, true, false);
    chart12->axis(0)->setMajorGridLines(true);
    chart12->axis(1)->setMajorGridLines(true);

    // custom gridlines
    Chart *chart13 = xlsx.insertChart(4, 15, QSize(300, 300));
    chart13->setType(Chart::Type::Scatter);
    chart13->setLegend(Legend::Position::Right);
    chart13->setTitle("Custom major gridlines");
    chart13->addSeries(CellRange(1,1,3,10), NULL, true, true, false);
    chart13->axis(0)->setMajorGridLines(Qt::red, 1.5, LineFormat::StrokeType::Dot);
    chart13->axis(1)->setMajorGridLines(Qt::blue, 0.5, LineFormat::StrokeType::DashDot);

    // standard gridlines
    Chart *chart14 = xlsx.insertChart(4, 21, QSize(300, 300));
    chart14->setType(Chart::Type::Scatter);
    chart14->setLegend(Legend::Position::Right);
    chart14->setTitle("Major and minor gridlines");
    chart14->addSeries(CellRange(1,1,3,10), NULL, true, true, false);
    chart14->axis(0)->setMajorGridLines(true);
    chart14->axis(1)->setMajorGridLines(true);
    chart14->axis(0)->setMinorGridLines(true);
    chart14->axis(1)->setMinorGridLines(true);

    /// Titles

    // plain string title
    Chart *chart21 = xlsx.insertChart(20, 3, QSize(300, 300));
    chart21->setType(Chart::Type::Scatter);
    chart21->setLegend(Legend::Position::Right);
    chart21->setTitle("Plain title");
    chart21->addSeries(CellRange(1,1,3,10), NULL, true, true, false);

    // referenced title
    //note that the reference range includes 2 rows
    Chart *chart22 = xlsx.insertChart(20, 9, QSize(300, 300));
    chart22->setType(Chart::Type::Scatter);
    chart22->setLegend(Legend::Position::Right);
    chart22->title().setStringReference(CellRange(4,1,5,1), xlsx.currentSheet());
    chart22->addSeries(CellRange(1,1,3,10), NULL, true, true, false);

    // rich formatted title
    Chart *chart23 = xlsx.insertChart(20, 15, QSize(300, 300));
    chart23->setType(Chart::Type::Scatter);
    chart23->setLegend(Legend::Position::Right);
    chart23->title().setHtml("<i>Rich</i> <b><font color=\"red\">formatted</font> title");
    chart23->addSeries(CellRange(1,1,3,10), NULL, true, true, false);

    // referenced title with some custom formatting
    Chart *chart24 = xlsx.insertChart(20, 21, QSize(300, 300));
    chart24->setType(Chart::Type::Scatter);
    chart24->setLegend(Legend::Position::Right);
    chart24->title().setStringReference(CellRange(4,1,5,1), xlsx.currentSheet());
    chart24->title().defaultCharacterProperties().underline = CharacterProperties::UnderlineType::Double;
    chart24->title().defaultCharacterProperties().capitalization = CharacterProperties::CapitalizationType::SmallCaps;
    chart24->title().defaultCharacterProperties().highlightColor = Color(Color::SchemeColor::Accent3);
    chart24->addSeries(CellRange(1,1,3,10), NULL, true, true, false);

    /// Title positioning and other properties

    // title moved to the top-right corner
    Chart *chart31 = xlsx.insertChart(36, 3, QSize(300, 300));
    chart31->setType(Chart::Type::Scatter);
    chart31->setLegend(Legend::Position::Right);
    chart31->setTitle("Title right-aligned");
    chart31->title().moveTo({1,0});
    chart31->addSeries(CellRange(1,1,3,10), NULL, true, true, false);

    // referenced title
    //note that the reference range includes 2 rows
    Chart *chart32 = xlsx.insertChart(36, 9, QSize(300, 300));
    chart32->setType(Chart::Type::Scatter);
    chart32->setLegend(Legend::Position::Right);
    chart32->setTitle("Title overlayed");
    chart32->title().setOverlay(true);
    chart32->addSeries(CellRange(1,1,3,10), NULL, true, true, false);

    // rich formatted title
    Chart *chart33 = xlsx.insertChart(36, 15, QSize(300, 300));
    chart33->setType(Chart::Type::Scatter);
    chart33->setLegend(Legend::Position::Right);
    chart33->setTitle("Title with line and fill");
    chart33->title().shape().line().setStrokeType(LineFormat::StrokeType::Solid);
    chart33->title().shape().line().setColor(Color::SystemColor::Highlight);
    chart33->title().shape().fill().setType(FillFormat::FillType::GradientFill);
    chart33->title().shape().fill().setGradientList({{0, Color::SchemeColor::Accent1},{1, Color::SchemeColor::Accent2}});

    chart33->addSeries(CellRange(1,1,3,10), NULL, true, true, false);

    // referenced title with some custom formatting
    Chart *chart34 = xlsx.insertChart(36, 21, QSize(300, 300));
    chart34->setType(Chart::Type::Scatter);
    chart34->setLegend(Legend::Position::Right);
    chart34->title().setStringReference(CellRange(4,1,5,1), xlsx.currentSheet());
    chart34->title().shape().setPresetGeometry(ShapeType::can);
    chart34->title().shape().line().setStrokeType(LineFormat::StrokeType::Solid);
    chart34->title().shape().line().setColor(Color::SchemeColor::Accent6);
    chart34->addSeries(CellRange(1,1,3,10), NULL, true, true, false);

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
