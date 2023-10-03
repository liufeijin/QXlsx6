// demo.cpp

#include <QtGlobal>
#include <QtCore>
#include <QDebug>

#if QT_VERSION >= QT_VERSION_CHECK(5,15,0) // Qt 5.15 or higher
#include <QRandomGenerator>
#endif

#include "xlsxdocument.h"
#include "xlsxformat.h"
#include "xlsxcellrange.h"
#include "xlsxworksheet.h"
#include "xlsxcellformula.h"

QXLSX_USE_NAMESPACE

void writeHorizontalAlignCell(Worksheet *sheet, const QString &cell, const QString &text, Format::HorizontalAlignment align)
{
   Format format;
   format.setHorizontalAlignment(align);
   format.setBorderStyle(Format::BorderThin);
   sheet->write(cell, text, format);
}

void writeVerticalAlignCell(Worksheet *sheet, const QString &range, const QString &text, Format::VerticalAlignment align)
{
   Format format;
   format.setVerticalAlignment(align);
   format.setBorderStyle(Format::BorderThin);
   CellRange r(range);
   sheet->write(r.firstRow(), r.firstColumn(), text);
   sheet->mergeCells(r, format);
}

void writeBorderStyleCell(Worksheet *sheet, const QString &cell, const QString &text, Format::BorderStyle bs)
{
   Format format;
   format.setBorderStyle(bs);
   sheet->write(cell, text, format);
}

void writeSolidFillCell(Worksheet *sheet, const QString &cell, const QColor &color)
{
   Format format;
   format.setPatternBackgroundColor(color);
   sheet->write(cell, QVariant(), format);
}

void writePatternFillCell(Worksheet *sheet, const QString &cell, Format::FillPattern pattern, const QColor &color)
{
   Format format;
   format.setPatternForegroundColor(color);
   format.setFillPattern(pattern);
   sheet->write(cell, QVariant(), format);
}

void writeBorderAndFontColorCell(Worksheet *sheet, const QString &cell, const QString &text, const QColor &color)
{
   Format format;
   format.setBorderStyle(Format::BorderThin);
   format.setBorderColor(color);
   format.setFontColor(color);
   sheet->write(cell, text, format);
}

void writeFontNameCell(Worksheet *sheet, const QString &cell, const QString &text)
{
    Format format;
    format.setFontName(text);
    format.setFontSize(16);
    sheet->write(cell, text, format);
}

void writeFontSizeCell(Worksheet *sheet, const QString &cell, int size)
{
    Format format;
    format.setFontSize(size);
    sheet->write(cell, "Qt Xlsx", format);
}

void writeInternalNumFormatsCell(Worksheet *sheet, int row, int col, double value, int numFmt)
{
    Format format;
    format.setNumberFormatIndex(numFmt);
    sheet->write(row, col, value);
    sheet->write(row, col+1, QString("Builtin NumFmt %1").arg(numFmt));
    sheet->write(row, col+2, value, format);
}

void writeCustomNumFormatsCell(Worksheet *sheet, int row, double value, const QString &numFmt)
{
    Format format;
    format.setNumberFormat(numFmt);
    sheet->write(row, 1, value);
    sheet->write(row, 2, numFmt);
    sheet->write(row, 3, value, format);
}

int main(int argc, char *argv[])
{
    QCoreApplication app(argc, argv);

    Document xlsx;


    //---------------------------------------------------------------
    //Create the first sheet (Otherwise, default "Sheet1" will be created)
    xlsx.addSheet("Aligns & Borders");
    auto sheet = xlsx.activeWorksheet();
    sheet->setColumnWidth(2, 20); //Column B
    sheet->setColumnWidth(8, 12); //Column H
    sheet->setGridLinesVisible(false);
    sheet->setTabColor(Qt::magenta);

    //Alignment
    writeHorizontalAlignCell(sheet, "B3", "AlignLeft", Format::AlignLeft);
    writeHorizontalAlignCell(sheet, "B5", "AlignHCenter", Format::AlignHCenter);
    writeHorizontalAlignCell(sheet, "B7", "AlignRight", Format::AlignRight);
    writeVerticalAlignCell(sheet, "D3:D7", "AlignTop", Format::AlignTop);
    writeVerticalAlignCell(sheet, "F3:F7", "AlignVCenter", Format::AlignVCenter);
    writeVerticalAlignCell(sheet, "H3:H7", "AlignBottom", Format::AlignBottom);

    //Border
    writeBorderStyleCell(sheet, "B13", "BorderMedium", Format::BorderMedium);
    writeBorderStyleCell(sheet, "B15", "BorderDashed", Format::BorderDashed);
    writeBorderStyleCell(sheet, "B17", "BorderDotted", Format::BorderDotted);
    writeBorderStyleCell(sheet, "B19", "BorderThick", Format::BorderThick);
    writeBorderStyleCell(sheet, "B21", "BorderDouble", Format::BorderDouble);
    writeBorderStyleCell(sheet, "B23", "BorderDashDot", Format::BorderDashDot);

    //Fill
    writeSolidFillCell(sheet, "D13", Qt::red);
    writeSolidFillCell(sheet, "D15", Qt::blue);
    writeSolidFillCell(sheet, "D17", Qt::yellow);
    writeSolidFillCell(sheet, "D19", Qt::magenta);
    writeSolidFillCell(sheet, "D21", Qt::green);
    writeSolidFillCell(sheet, "D23", Qt::gray);
    writePatternFillCell(sheet, "F13", Format::PatternMediumGray, Qt::red);
    writePatternFillCell(sheet, "F15", Format::PatternDarkHorizontal, Qt::blue);
    writePatternFillCell(sheet, "F17", Format::PatternDarkVertical, Qt::yellow);
    writePatternFillCell(sheet, "F19", Format::PatternDarkDown, Qt::magenta);
    writePatternFillCell(sheet, "F21", Format::PatternLightVertical, Qt::green);
    writePatternFillCell(sheet, "F23", Format::PatternLightTrellis, Qt::gray);

    writeBorderAndFontColorCell(sheet, "H13", "Qt::red", Qt::red);
    writeBorderAndFontColorCell(sheet, "H15", "Qt::blue", Qt::blue);
    writeBorderAndFontColorCell(sheet, "H17", "Qt::yellow", Qt::yellow);
    writeBorderAndFontColorCell(sheet, "H19", "Qt::magenta", Qt::magenta);
    writeBorderAndFontColorCell(sheet, "H21", "Qt::green", Qt::green);
    writeBorderAndFontColorCell(sheet, "H23", "Qt::gray", Qt::gray);

    //---------------------------------------------------------------
    //Create the second sheet.
    xlsx.addSheet("Fonts");
    sheet = xlsx.activeWorksheet();
    sheet->setTabColor(Qt::yellow);

    sheet->write("B3", "Normal");
    Format font_bold;
    font_bold.setFontBold(true);
    sheet->write("B4", "Bold", font_bold);
    Format font_italic;
    font_italic.setFontItalic(true);
    sheet->write("B5", "Italic", font_italic);
    Format font_underline;
    font_underline.setFontUnderline(Format::FontUnderlineSingle);
    sheet->write("B6", "Underline", font_underline);
    Format font_strikeout;
    font_strikeout.setFontStrikeOut(true);
    sheet->write("B7", "StrikeOut", font_strikeout);

    writeFontNameCell(sheet, "D3", "Arial");
    writeFontNameCell(sheet, "D4", "Arial Black");
    writeFontNameCell(sheet, "D5", "Comic Sans MS");
    writeFontNameCell(sheet, "D6", "Courier New");
    writeFontNameCell(sheet, "D7", "Impact");
    writeFontNameCell(sheet, "D8", "Times New Roman");
    writeFontNameCell(sheet, "D9", "Verdana");

    writeFontSizeCell(sheet, "G3", 10);
    writeFontSizeCell(sheet, "G4", 12);
    writeFontSizeCell(sheet, "G5", 14);
    writeFontSizeCell(sheet, "G6", 16);
    writeFontSizeCell(sheet, "G7", 18);
    writeFontSizeCell(sheet, "G8", 20);
    writeFontSizeCell(sheet, "G9", 25);

    Format font_vertical;
    font_vertical.setRotation(255);
    font_vertical.setFontSize(16);
    sheet->write("J3", "vertical", font_vertical);
    sheet->mergeCells("J3:J9");

    //---------------------------------------------------------------
    //Create the third sheet.
    xlsx.addSheet("Formulas");
    sheet = xlsx.activeWorksheet();
    sheet->setTabColor(Qt::darkCyan);
    sheet->setColumnWidth(1, 2, 40);
    Format rAlign;
    rAlign.setHorizontalAlignment(Format::AlignRight);
    Format lAlign;
    lAlign.setHorizontalAlignment(Format::AlignLeft);
    sheet->write("B3", 40, lAlign);
    sheet->write("B4", 30, lAlign);
    sheet->write("B5", 50, lAlign);
    sheet->write("A7", "SUM(B3:B5)=", rAlign);
    sheet->write("B7", "=SUM(B3:B5)", lAlign);
    sheet->write("A8", "AVERAGE(B3:B5)=", rAlign);
    sheet->write("B8", "=AVERAGE(B3:B5)", lAlign);
    sheet->write("A9", "MAX(B3:B5)=", rAlign);
    sheet->write("B9", "=MAX(B3:B5)", lAlign);
    sheet->write("A10", "MIN(B3:B5)=", rAlign);
    sheet->write("B10", "=MIN(B3:B5)", lAlign);
    sheet->write("A11", "COUNT(B3:B5)=", rAlign);
    sheet->write("B11", "=COUNT(B3:B5)", lAlign);

    sheet->write("A13", "IF(B7>100,\"large\",\"small\")=", rAlign);
    sheet->write("B13", "=IF(B7>100,\"large\",\"small\")", lAlign);

    sheet->write("A15", "SQRT(25)=", rAlign);
    sheet->write("B15", "=SQRT(25)", lAlign);
    sheet->write("A16", "RAND()=", rAlign);
    sheet->write("B16", "=RAND()", lAlign);
    sheet->write("A17", "2*PI()=", rAlign);
    sheet->write("B17", "=2*PI()", lAlign);

    sheet->write("A19", "UPPER(\"qtxlsx\")=", rAlign);
    sheet->write("B19", "=UPPER(\"qtxlsx\")", lAlign);
    sheet->write("A20", "LEFT(\"ubuntu\",3)=", rAlign);
    sheet->write("B20", "=LEFT(\"ubuntu\",3)", lAlign);
    sheet->write("A21", "LEN(\"Hello Qt!\")=", rAlign);
    sheet->write("B21", "=LEN(\"Hello Qt!\")", lAlign);

    Format dateFormat;
    dateFormat.setHorizontalAlignment(Format::AlignLeft);
    dateFormat.setNumberFormat("yyyy-mm-dd");
    sheet->write("A23", "DATE(2013,8,13)=", rAlign);
    sheet->write("B23", "=DATE(2013,8,13)", dateFormat);
    sheet->write("A24", "DAY(B23)=", rAlign);
    sheet->write("B24", "=DAY(B23)", lAlign);
    sheet->write("A25", "MONTH(B23)=", rAlign);
    sheet->write("B25", "=MONTH(B23)", lAlign);
    sheet->write("A26", "YEAR(B23)=", rAlign);
    sheet->write("B26", "=YEAR(B23)", lAlign);
    sheet->write("A27", "DAYS360(B23,TODAY())=", rAlign);
    sheet->write("B27", "=DAYS360(B23,TODAY())", lAlign);

    sheet->write("A29", "B3+100*(2-COS(0)))=", rAlign);
    sheet->write("B29", "=B3+100*(2-COS(0))", lAlign);
    sheet->write("A30", "ISNUMBER(B29)=", rAlign);
    sheet->write("B30", "=ISNUMBER(B29)", lAlign);
    sheet->write("A31", "AND(1,0)=", rAlign);
    sheet->write("B31", "=AND(1,0)", lAlign);

    sheet->write("A33", "HYPERLINK(\"http://qt-project.org\")=", rAlign);
    sheet->write("B33", "=HYPERLINK(\"http://qt-project.org\")", lAlign);

    // Array formulas
    sheet->write(2,4,"Array formulas");
    for (int row=3; row<21; ++row) {
        sheet->write(row, 4, row*2);
        sheet->write(row, 5, row*3);
    }
    sheet->writeFormula("F3", CellFormula("D3:D20+E3:E20", "F3:F20", CellFormula::Type::Array));
    sheet->writeFormula("G3", CellFormula("=CONCATENATE(\"The total is \",D3:D20,\" units\")", "G3:G20", CellFormula::Type::Array));

    sheet->write("H2","Shared formulas");
    sheet->writeFormula("H3", CellFormula("=D3+E3", "H3:H20", CellFormula::Type::Shared));
    sheet->writeFormula("I3", CellFormula("=CONCATENATE(\"The total is \",D3,\" units\")", "I3:I20", CellFormula::Type::Shared));


    //---------------------------------------------------------------
    //Create the fourth sheet.
    xlsx.addSheet("NumFormats");
    sheet = xlsx.activeWorksheet();
    sheet->setTabColor(Qt::darkRed);
    sheet->write(3, 1, "Raw data");
    sheet->write(3, 2, "Builtin Format");
    sheet->write(3, 3, "Shown value");
    writeInternalNumFormatsCell(sheet, 4, 1, 2.5681, 2);
    writeInternalNumFormatsCell(sheet, 5, 1, 2500000, 3);
    writeInternalNumFormatsCell(sheet, 6, 1, -500, 5);
    writeInternalNumFormatsCell(sheet, 7, 1, -0.25, 9);
    writeInternalNumFormatsCell(sheet, 8, 1, 890, 11);
    writeInternalNumFormatsCell(sheet, 9, 1, 0.75, 12);
    writeInternalNumFormatsCell(sheet, 10, 1, 41499, 14);
    writeInternalNumFormatsCell(sheet, 11, 1, 41499, 17);

    sheet->write(3, 5, "Raw data");
    sheet->write(3, 6, "Builtin Format");
    sheet->write(3, 7, "Shown value");
    for (int i=0; i<50; ++i) {
        writeInternalNumFormatsCell(sheet, i+4, 5, 100, i);
    }

    writeCustomNumFormatsCell(sheet, 13, 20.5627, "#.###");
    writeCustomNumFormatsCell(sheet, 14, 4.8, "#.00");
    writeCustomNumFormatsCell(sheet, 15, 1.23, "0.00 \"RMB\"");
    writeCustomNumFormatsCell(sheet, 16, 60, "[Red][<=100];[Green][>100]");
    writeCustomNumFormatsCell(sheet, 17, 4000, "yyyy-mmm-dd");

    sheet->autosizeColumnsWidth(1,7);

    //---------------------------------------------------------------
    //Create the fifth sheet.
    xlsx.addSheet("Grouping");
    sheet = xlsx.activeWorksheet();
    sheet->setTabColor(Qt::darkBlue);

#if QT_VERSION >= QT_VERSION_CHECK(5,15,0) // Qt 5.15 or over
    QRandomGenerator rgen;
#else
    qsrand(QDateTime::currentMSecsSinceEpoch());
#endif

    for (int row=2; row<31; ++row) {
        for (int col=1; col<=10; ++col) {
#if QT_VERSION >= QT_VERSION_CHECK(5,15,0) // Qt 5.15 or over
            sheet->write(row, col, rgen.generate() % 100);
#else
            sheet->write(row, col, qrand() % 100);
#endif
        }
    }
    sheet->groupRows(4, 7);
    sheet->groupRows(11, 26, false);
    sheet->groupRows(15, 17, false);
    sheet->groupRows(20, 22, false);
    sheet->setColumnWidth(1, 10, 5.0);
    qDebug() << "grouping columns [1-20]" << sheet->groupColumns(1, 20, false);
    qDebug() << "grouping columns [2-19]" << sheet->groupColumns(2, 19, false);
    qDebug() << "grouping columns [3-18]" << sheet->groupColumns(3, 18, false);
    qDebug() << "grouping columns [4-17]" << sheet->groupColumns(4, 17, false);
    qDebug() << "grouping columns [5-16]" << sheet->groupColumns(5, 16, false);
    qDebug() << "grouping columns [6-15]" << sheet->groupColumns(6, 15, false);
    qDebug() << "grouping columns [7-14]" << sheet->groupColumns(7, 14, false);
    qDebug() << "grouping columns [8-13]" << sheet->groupColumns(8, 13, false);

    sheet->setRowHeight(2,10);
    sheet->setRowHeight(3,10,20);
    sheet->setRowHeight(30,40,30);

    //select cells
    for (int row=30; row<=40; ++row) {
        for (int column = 1; column <=20; ++column) {
            if ((row+column) % 2 == 0) sheet->addSelection(CellRange(row,column,row,column));
        }
    }

    sheet->setFitToPage(true);

    xlsx.saveAs("demo1.xlsx");

    //Make sure that read/write works well.
    Document xlsx2("demo1.xlsx");
    xlsx2.saveAs("demo2.xlsx"); 

    qDebug() << "**** end of main() ****";

    return 0;
}
