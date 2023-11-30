#include <QtGlobal>
#include <QObject>
#include <QString>
#include <QDebug>

#include "xlsxdocument.h"
#include "xlsxchartsheet.h"
#include "xlsxcellrange.h"
#include "xlsxchart.h"
#include "xlsxrichstring.h"
#include "xlsxworkbook.h"

QXLSX_USE_NAMESPACE

int pages()
{
    Document xlsx;
    auto sheet = xlsx.activeWorksheet();

    // Set page margins of 5 mm each, no header / footer

    sheet->write("A1", "This sheet has page margins of 5 mm");
    sheet->setPageMarginsMm(5,5,5,5);

    // Set zero page margins

    sheet = xlsx.addWorksheet();
    sheet->write("A1", "This sheet has page margins equal to zero");
    sheet->setPageMarginsMm(0,0,0,0);

    // Set default page margins, specify print options, view type

    sheet = xlsx.addWorksheet();
    sheet->write("A1", "Hello QtXlsx!");
    sheet->write("A2", 12345);
    sheet->write("A3", "=44+33"); // cell value is 77.
    sheet->write("A4", true);
    sheet->write("A5", "http://qt-project.org");
    sheet->write("A6", QDate(2013, 12, 27));
    sheet->write("A7", QTime(6, 30));

    sheet->write("G1", "Hello QtXlsx!");
    sheet->write("G2", 12345);
    sheet->write("G3", "=44+33"); // cell value is 77.
    sheet->write("G4", true);
    sheet->write("G5", "http://qt-project.org");
    sheet->write("G6", QDate(2013, 12, 27));
    sheet->write("G7", QTime(6, 30));

    sheet->write("M1", "Hello QtXlsx!");
    sheet->write("M2", 12345);
    sheet->write("M3", "=44+33"); // cell value is 77.
    sheet->write("M4", true);
    sheet->write("M5", "http://qt-project.org");
    sheet->write("M6", QDate(2013, 12, 27));
    sheet->write("M7", QTime(6, 30));

    sheet->setDefaultPageMargins();

    sheet->setPaperSize(PageSetup::PaperSize::A5);
    sheet->setPageOrientation(PageSetup::Orientation::Portrait);

    sheet->setFitToWidth(2);
    sheet->setPrintCellComments(PageSetup::CellComments::AsDisplayed);
    sheet->setFirstPageNumber(300);

    sheet->setViewType(SheetView::Type::PageLayout);
    sheet->setShowAutoPageBreaks(true);

    sheet->setHeaders("odd page", "even page", "1st page");
    sheet->setFooters("&RPage &P of &N", "&L&D &T","&C&F");

    // Two views:
    // 1. no grid lines, no ruler visible
    // 2. split by two panes

    sheet = xlsx.addWorksheet();
    // 1st view
    sheet->setGridLinesVisible(false, 0);
    sheet->setRulerVisible(false, 0);
    // 2nd view
    sheet->addView();
    sheet->setViewColorIndex(6, 1);
    sheet->setDefaultGridColorUsed(false, 1);
//    sheet->splitViewVertically(2, true, 1);
    sheet->write(1,1,"The first two rows are frozen");
//    sheet->splitViewHorizontally(2, false, 1);
    sheet->splitView(5,4,false, 1);

    if (!xlsx.saveAs("PageMargins.xlsx")) {
        qDebug() << "[PageMargins] failed to save excel file";
        return (-2);
    }

    Document xlsx2;
    for (int i=0; i<=64; ++i) {
        auto sheet = xlsx2.addWorksheet(QString("color %1").arg(i));
        sheet->setViewColorIndex(i);
        sheet->setDefaultGridColorUsed(false);
    }
    xlsx2.saveAs("colors.xlsx");

    return 0;
}
