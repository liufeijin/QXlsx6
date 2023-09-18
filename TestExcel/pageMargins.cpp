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

    xlsx.write("A1", "Hello QtXlsx!"); // current file is utf-8 character-set.
    xlsx.write("A2", 12345);
    xlsx.write("A3", "=44+33"); // cell value is 77.
    xlsx.write("A4", true);
    xlsx.write("A5", "http://qt-project.org");
    xlsx.write("A6", QDate(2013, 12, 27));
    xlsx.write("A7", QTime(6, 30));

    auto sheet = xlsx.currentWorksheet();
    sheet->setPageMarginsMm(5,5,5,5);

    auto sheet1 = xlsx.addWorksheet();
    sheet1->write("A1", "Hello QtXlsx!"); // current file is utf-8 character-set.
    sheet1->write("A2", 12345);
    sheet1->write("A3", "=44+33"); // cell value is 77.
    sheet1->write("A4", true);
    sheet1->write("A5", "http://qt-project.org");
    sheet1->write("A6", QDate(2013, 12, 27));
    sheet1->write("A7", QTime(6, 30));
    sheet1->setPageMarginsMm(0,0,0,0);

    auto sheet2 = xlsx.addWorksheet();
    sheet2->write("A1", "Hello QtXlsx!"); // current file is utf-8 character-set.
    sheet2->write("A2", 12345);
    sheet2->write("A3", "=44+33"); // cell value is 77.
    sheet2->write("A4", true);
    sheet2->write("A5", "http://qt-project.org");
    sheet2->write("A6", QDate(2013, 12, 27));
    sheet2->write("A7", QTime(6, 30));
    sheet2->setDefaultPageMargins();

    PageSetup pageSetup;
    pageSetup.paperSize = PageSetup::PaperSize::A5;
    pageSetup.orientation = PageSetup::Orientation::Portrait;
    sheet2->setPageSetup(pageSetup);

    sheet2->setViewType(SheetView::Type::PageLayout);

    if (!xlsx.saveAs("PageMargins.xlsx")) {
        qDebug() << "[PageMargins] failed to save excel file";
        return (-2);
    }

    return 0;
}
