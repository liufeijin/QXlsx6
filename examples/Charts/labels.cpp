#include <QtGlobal>
#include <QtCore>
#include <QDebug>

#include "xlsxdocument.h"
#include "xlsxcellrange.h"
#include "xlsxchart.h"
#include "xlsxmain.h"

QXLSX_USE_NAMESPACE

int chartlabels()
{
    Document xlsx;
    auto sheet = xlsx.activeWorksheet();

    for (int i=1; i<10; ++i)
    {
        sheet->write(1, i+1, i);
    }

    sheet->write(2, 1, "Set 1");
    sheet->write(3, 1, "Set 2");

    for (int i=1; i<10; ++i)
    {
        sheet->write(2, i+1, i*i*i);   //A2:A10
        sheet->write(3, i+1, i*i);    //B2:B10
    }

    /// Default labels in two series

    Chart *chart11 = sheet->insertChart(4, 3, QSize(300, 300));
    chart11->addSubchart(Chart::Type::Line);
    chart11->setLegend(Legend::Position::Right);
    chart11->setTitle("Default labels");
    chart11->addSeries(CellRange(1,1,3,10), NULL, true, true, false);
    // adds labels to all series in the subchart
    chart11->labels(0).setVisible(true);
    chart11->labels(0).setShowParameters(Label::ShowParameter::ShowValue | Label::ShowParameter::ShowCategory);
    chart11->labels(0).setPosition(Label::Position::Top);


    /// Labels only in the first series
    Chart *chart12 = sheet->insertChart(4, 9, QSize(300, 300));
    chart12->addSubchart(Chart::Type::Scatter);
    chart12->setLegend(Legend::Position::Right);
    chart12->setTitle("Labels in 1st series");
    chart12->addSeries(CellRange(1,1,3,10), NULL, true, true, false);
    // adds labels to 1st series
    chart12->series(0)->defaultLabels().setVisible(true);
    chart12->series(0)->defaultLabels().setShowParameters(Label::ShowParameter::ShowValue | Label::ShowParameter::ShowCategory);

    /// Label to the 6th data point of 1st series
    Chart *chart13 = sheet->insertChart(4, 15, QSize(300, 300));
    chart13->addSubchart(Chart::Type::Scatter);
    chart13->setLegend(Legend::Position::Right);
    chart13->setTitle("Default labels");
    chart13->addSeries(CellRange(1,1,3,10), NULL, true, true, false);
    // adds label to 1st series
    auto &label = chart13->series(0)->label(5)->get();
    label.setVisible(true);
    label.setPosition(Label::Position::Center);
    label.setShowParameters(Label::ShowParameter::ShowValue | Label::ShowParameter::ShowCategory);

    xlsx.saveAs("chartLabels1.xlsx");

    Document xlsx2("chartLabels1.xlsx");
    if ( xlsx2.isLoaded() )
    {
        xlsx2.saveAs("chartLabels2.xlsx");
    }

    return 0;
}
