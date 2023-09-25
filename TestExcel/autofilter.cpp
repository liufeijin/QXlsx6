// autofilter.cpp

#include <QtGlobal>
#include <QtCore>

#include "xlsxdocument.h"
#include "xlsxformat.h"

int autofilter()
{
    QXlsx::Document xlsx;

    // 1. Filtering by values
    auto sheet = xlsx.currentWorksheet();
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
    sheet->autofilter().setCustomPredicate(3, 10, QXlsx::Filter::Predicate::GreaterThanOrEqual,
                                              40, QXlsx::Filter::Predicate::LessThanOrEqual,
                                              QXlsx::Filter::Operation::And);

    xlsx.saveAs("autofilter.xlsx");
    return 0;
}
