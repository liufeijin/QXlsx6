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

    sheet->write("A1", "Hello QtXlsx!");
    sheet->write("A2", 12345);
    sheet->write("A3", "=44+33"); // cell value is 77.
    sheet->write("A4", true);
    sheet->write("A5", "http://qt-project.org");
    sheet->write("A6", QDate(2013, 12, 27));
    sheet->write("A7", QTime(6, 30));

    sheet->setPageMarginsMm(5,5,5,5);

    // Set zero page margins

    auto sheet1 = xlsx.addWorksheet();
    sheet1->write("A1", "Hello QtXlsx!");
    sheet1->write("A2", 12345);
    sheet1->write("A3", "=44+33"); // cell value is 77.
    sheet1->write("A4", true);
    sheet1->write("A5", "http://qt-project.org");
    sheet1->write("A6", QDate(2013, 12, 27));
    sheet1->write("A7", QTime(6, 30));
    sheet1->setPageMarginsMm(0,0,0,0);

    // Set default page margins, specify print options, view type

    auto sheet2 = xlsx.addWorksheet();
    sheet2->write("A1", "Hello QtXlsx!");
    sheet2->write("A2", 12345);
    sheet2->write("A3", "=44+33"); // cell value is 77.
    sheet2->write("A4", true);
    sheet2->write("A5", "http://qt-project.org");
    sheet2->write("A6", QDate(2013, 12, 27));
    sheet2->write("A7", QTime(6, 30));

    sheet2->write("G1", "Hello QtXlsx!");
    sheet2->write("G2", 12345);
    sheet2->write("G3", "=44+33"); // cell value is 77.
    sheet2->write("G4", true);
    sheet2->write("G5", "http://qt-project.org");
    sheet2->write("G6", QDate(2013, 12, 27));
    sheet2->write("G7", QTime(6, 30));

    sheet2->write("M1", "Hello QtXlsx!");
    sheet2->write("M2", 12345);
    sheet2->write("M3", "=44+33"); // cell value is 77.
    sheet2->write("M4", true);
    sheet2->write("M5", "http://qt-project.org");
    sheet2->write("M6", QDate(2013, 12, 27));
    sheet2->write("M7", QTime(6, 30));

    sheet2->setDefaultPageMargins();

    sheet2->setPaperSize(PageSetup::PaperSize::A5);
    sheet2->setPageOrientation(PageSetup::Orientation::Portrait);

    sheet2->setViewType(SheetView::Type::PageLayout);
    sheet2->setShowAutoPageBreaks(true);

    sheet2->setHeaders("odd page", "even page", "1st page");
    sheet2->setFooters("&RPage &P of &N", "&L&D &T","&C&F");

    if (!xlsx.saveAs("PageMargins.xlsx")) {
        qDebug() << "[PageMargins] failed to save excel file";
        return (-2);
    }

    return 0;
}
