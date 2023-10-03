// style.cpp

#include <QtGlobal>
#include <QtCore>

#include "xlsxdocument.h"
#include "xlsxformat.h"

int style()
{
    QXlsx::Document xlsx;
    auto sheet = xlsx.activeWorksheet();
    QXlsx::Format format1;
    format1.setFontColor(QColor(Qt::red));
    format1.setFontSize(15);
    format1.setHorizontalAlignment(QXlsx::Format::AlignHCenter);
    format1.setBorderStyle(QXlsx::Format::BorderDashDotDot);
    sheet->write("A1", "Hello Qt!", format1);
    sheet->write("B3", 12345, format1);

    QXlsx::Format format2;
    format2.setFontBold(true);
    format2.setFontUnderline(QXlsx::Format::FontUnderlineDouble);
    format2.setFillPattern(QXlsx::Format::PatternLightUp);
    sheet->write("C5", "=44+33", format2);
    sheet->write("D7", true, format2);

    QXlsx::Format format3;
    format3.setFontBold(true);
    format3.setFontColor(QColor(Qt::blue));
    format3.setFontSize(20);
    sheet->write(11, 1, "Hello Row Style");
    sheet->write(11, 6, "Blue Color");
    sheet->setRowFormat(11, 41, format3);

    QXlsx::Format format4;
    format4.setFontBold(true);
    format4.setFontColor(QColor(Qt::magenta));
    for (int row=21; row<=40; row++)
        for (int col=9; col<16; col++)
            xlsx.write(row, col, row+col);
    sheet->setColumnFormat(9, 16, format4);

    sheet->write("A5", QDate(2013, 8, 29));

    QXlsx::Format format6;
    format6.setPatternBackgroundColor(QColor(Qt::green));
    sheet->write("A6", "Background color: green", format6);

    xlsx.saveAs("style1.xlsx");

    return 0;
}
