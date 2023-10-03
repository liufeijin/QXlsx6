// autofilter.cpp

#include <QtGlobal>
#include <QtCore>

#include "xlsxdocument.h"
#include "xlsxformat.h"

int main()
{
    QXlsx::Document xlsx;

    // 1. Filtering by values
    auto sheet = xlsx.activeWorksheet();
    sheet->rename("Filter by values"); //rename default sheet

    for (int i=1; i<=10; ++i) {
        sheet->write(1,i, QString("column %1").arg(i));
        for (int j=2; j<=31; ++j) {
            sheet->write(j,i,i+j);
        }
    }

    sheet->autofilter().setRange(QXlsx::CellRange(1,1,31,10));
    sheet->autofilter().setFilterByValues(0, {2,4,6,8,10,12});

    // 2. Filtering by predicates
    sheet = xlsx.addWorksheet("filter by predicates");
    for (int i=1; i<=10; ++i) {
        sheet->write(1,i, QString("column %1").arg(i));
        for (int j=2; j<=31; ++j) {
            sheet->write(j,i,i+j);
        }
    }
    sheet->autofilter().setRange(QXlsx::CellRange(1,1,31,10));
    sheet->autofilter().setPredicate(3,
                                     QXlsx::Filter::Predicate::GreaterThanOrEqual, 10,
                                     QXlsx::Filter::Operation::And,
                                     QXlsx::Filter::Predicate::LessThanOrEqual, 40
                                     );

    //3. Dynamic filtering of dates
    sheet = xlsx.addWorksheet(("dynamic filter"));
    sheet->write(1,1, "dates");
    sheet->write(1,2, "dates");
    const auto now = QDateTime::currentDateTime();
    for (int i=-5; i<=5; ++i) {
        auto dt = now.addMonths(i);
        sheet->write(i+7, 1, dt);
        dt = now.addDays(i*2);
        sheet->write(i+7, 2, dt);
    }
    sheet->autosizeColumnsWidth(1,2);
    sheet->autofilter().setRange(QXlsx::CellRange(1,1,12,2));
    sheet->autofilter().setDynamicFilter(0, QXlsx::Filter::DynamicFilterType::LastQuarter);
    sheet->autofilter().setDynamicFilter(1, QXlsx::Filter::DynamicFilterType::LastWeek);

    xlsx.saveAs("autofilter.xlsx");
    return 0;
}
