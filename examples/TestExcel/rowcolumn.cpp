// rowcolumn.cpp

#include <QtGlobal>
#include <QtCore>

#include "xlsxdocument.h"
#include "xlsxformat.h"

int rowcolumn()
{
    QXlsx::Document xlsx;
    auto sheet = xlsx.activeWorksheet();
    sheet->write(1, 2, "Row:0, Col:2 ==> (C1)");

    //Set the height of the first row to 50.0(points)
    sheet->setRowHeight(1, 50.0);

    //Set the width of the third column to 40.0(chars)
    sheet->setColumnWidth(3, 3, 40.0);

    //Set style for the row 11th.
    QXlsx::Format format1;
    format1.setFontBold(true);
    format1.setFontColor(QColor(Qt::blue));
    format1.setFontSize(20);
    sheet->write(11, 1, "Hello Row Style");
    sheet->write(11, 6, "Blue Color");
    sheet->setRowFormat(11, format1);
    sheet->setRowHeight(11, 41);

    //Set style for the col [9th, 16th)
    QXlsx::Format format2;
    format2.setFontBold(true);
    format2.setFontColor(QColor(Qt::magenta));
    for (int row=12; row<=30; row++)
        for (int col=9; col<=15; col++)
            xlsx.write(row, col, row+col);
    sheet->setColumnWidth(9, 16, 5.0);
    sheet->setColumnFormat(9, 16, format2);


    // Autosize columns
    sheet = xlsx.addWorksheet("Autosize");
    QXlsx::Format header;
    header.setFontBold(true);
    header.setFontSize(20);

    //Custom number formats
    int row = 2;
    sheet->write(row, 1, "Small", header);
    sheet->write(row, 2, "Large header", header);
    for (int i=0; i<10; ++i)
    {
        sheet->write(row+1+i, 1, i);
        sheet->write(row+1+i, 2, i);
    }
    sheet->autosizeColumnsWidth(1, 2);

    sheet->write(2, 4, "Small", header);
    sheet->write(2, 5, "<--- Large header --->", header);
    for (int i=0; i<10; ++i)
    {
        sheet->write(row+1+i, 4, exp(i));
        sheet->write(row+1+i, 5, exp(i));
    }
    sheet->autosizeColumnsWidth(4, 5);


    sheet->write(1, 7, "<--- Very big first line --->", header);
    sheet->write(row, 7, "Small", header);
    sheet->write(row, 8, "<--- Large header --->", header);
    for (int i=0; i<10; ++i)
    {
        sheet->write(row+1+i, 7, exp(i));
        sheet->write(row+1+i, 8, exp(i));
    }
    sheet->autosizeColumnsWidth(QXlsx::CellRange(2, 7, 1000, 8));


    xlsx.saveAs("rowcolumn.xlsx");
    return 0;
}
