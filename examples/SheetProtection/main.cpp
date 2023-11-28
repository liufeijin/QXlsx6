// main.cpp

#include <QtGlobal>
#include <QCoreApplication>
#include <QtCore>
#include <QDebug>
#include <QCryptographicHash>
#include <QPasswordDigestor>
#include <QFile>
#include <QBuffer>
#include <QDataStream>

#include "xlsxdocument.h"
#include "xlsxworkbook.h"
#include "xlsxformat.h"
#include "xlsxcellrange.h"
#include "xlsxworksheet.h"
#include "xlsxcellformula.h"
#include "xlsxsheetprotection.h"

int main(int argc, char *argv[])
{
    QCoreApplication app(argc, argv);

    QXlsx::Document xlsx;

    auto sheet1 = xlsx.addWorksheet("Protected sheet");
    if (!sheet1) return -1;

    sheet1->write(1, 1, "This sheet is protected with password '12345'");
    sheet1->write(2, 1, "Protected entities:");
    sheet1->write(2, 2, "Adding and removing columns");
    sheet1->write(3, 2, "Formatting cells");

    sheet1->sheetProtection().setProtectInsertColumns(true);
    sheet1->sheetProtection().setProtectDeleteColumns(true);
    sheet1->sheetProtection().setProtectFormatCells(true);

    // some algorithms adjustments
    QXlsx::PasswordProtection::setRandomizedSalt(true);
    sheet1->setPasswordProtection("12345");

    // Example of protecting a specific range

    sheet1->write(5,3, "Cells A5:A10 are protected with password 'a5a10'");
    QXlsx::CellRange r("A5:A10");
    sheet1->addProtectedRange(r, "range1");
    sheet1->protectedRange(0).setPasswordProtection("a5a10");
    QXlsx::Format f;
    f.setPatternBackgroundColor(QColor(Qt::red));
    sheet1->setFormat(r, f);

    xlsx.saveAs("SheetProtection.xlsx");

    // Example of checking the user's password with a value stored in the xlsx.

    auto sheet = xlsx.workbook()->worksheets().at(0);
    qDebug() << sheet->sheetProtection().protection().checkPassword("12345");

    qDebug() << "**** end of main() ****";

    return 0;
}
