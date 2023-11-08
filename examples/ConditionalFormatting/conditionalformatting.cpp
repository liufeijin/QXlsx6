#include <QtGlobal>
#include <QtCore>

#include "xlsxdocument.h"
#include "xlsxformat.h"
#include "xlsxconditionalformatting.h"
#include "xlsxcellrange.h"

int main()
{
    QXlsx::Document xlsx;
    auto sheet = xlsx.activeWorksheet();
        sheet->setName("Highlight cells"); //rename default sheet

    // 1. Formatting in the range
    sheet->write(1, 1, "Highlight cells with values in the range [400-800]");
    for (int i=2; i<=11; ++i) {
        sheet->write(i, 1, QRandomGenerator::global()->bounded(0,1000));
    }
    QXlsx::Format f1;
    f1.setFontColor(QColor(0x9C0006));
    f1.setPatternForegroundColor(QColor(0xFFC7CE));
    QXlsx::ConditionalFormatting format1;
    format1.setRange(QXlsx::CellRange(2, 1, 11, 1));
    format1.addHighlightCellsRule(QXlsx::ConditionalFormatting::Type::Between, "400", "800", f1);
    sheet->addConditionalFormatting(format1);

    // 2. Formatting of empty cells
    sheet->write(1, 3, "Highlight empty cells");
    for (int i=2; i<=11; ++i) {
        sheet->write(i, 3, QRandomGenerator::global()->bounded(0,1000));
    }
    sheet->writeBlank(QRandomGenerator::global()->bounded(2, 11), 3);
    sheet->writeBlank(QRandomGenerator::global()->bounded(2, 11), 3);
    QXlsx::Format f2;
    f2.setPatternForegroundColor(QColor(0xFFFFEB9C));
    QXlsx::ConditionalFormatting format2;
    format2.setRange(QXlsx::CellRange(2, 3, 11, 3));
    format2.addHighlightCellsRule(QXlsx::ConditionalFormatting::Type::Blanks, f2);
    sheet->addConditionalFormatting(format2);

    // 3. Two conditional formattings in one range
    sheet->write(1, 5, "Two conditional formattings in one range");
    sheet->write(2, 5, "aba");
    sheet->write(3, 5, "bcb");
    sheet->write(4, 5, "cdc");
    sheet->write(5, 5, "ded");
    sheet->write(6, 5, "efe");
    QXlsx::ConditionalFormatting format3;
    format3.setRange(QXlsx::CellRange(2, 5, 6, 5));
    format3.addHighlightCellsRule(QXlsx::ConditionalFormatting::Type::ContainsText, "b", f1);
    format3.addHighlightCellsRule(QXlsx::ConditionalFormatting::Type::ContainsText, "c", f2);
    sheet->addConditionalFormatting(format3);

    // 4. Two conditional formattings in one range, but with priorities
    sheet->write(1, 7, "Two conditional formattings in one range, but with priorities");
    sheet->write(2, 7, "aba");
    sheet->write(3, 7, "bcb");
    sheet->write(4, 7, "cdc");
    sheet->write(5, 7, "ded");
    sheet->write(6, 7, "efe");
    QXlsx::ConditionalFormatting format4;
    format4.setRange(QXlsx::CellRange(2, 7, 6, 7));
    format4.addHighlightCellsRule(QXlsx::ConditionalFormatting::Type::ContainsText, "b",
                                  QXlsx::ConditionalFormatting::PredefinedFormat::Format1, true);
    format4.addHighlightCellsRule(QXlsx::ConditionalFormatting::Type::ContainsText, "c",
                                  QXlsx::ConditionalFormatting::PredefinedFormat::Format2);
    format4.updateRulesPriorities(true);
    sheet->addConditionalFormatting(format4);

    qDebug() << QString("The sheet has %1 conditional formattings").arg(sheet->conditionalFormattingCount());
    qDebug() << QLatin1String("Conditional formatting with index 1 is applied to"
                              " the ranges")<<sheet->conditionalFormatting(1).ranges();

    sheet->autosizeColumnsWidth();

    // 5. Conditional formatting with color scale
    sheet = xlsx.addWorksheet("Color scale");
    sheet->write(1, 1, "Highlight cells with color scaling from 0 to 1000");
    for (int i=2; i<=11; ++i) {
        for (int j = 1; j <= 10; ++j)
            sheet->write(i, j, QRandomGenerator::global()->bounded(0,1000));
    }

    QXlsx::ConditionalFormatting format5;
    format5.setRange(QXlsx::CellRange(2, 1, 11, 10));
    format5.add2ColorScaleRule(QColor(Qt::white), QColor(0xFFC6EFCE));
    sheet->addConditionalFormatting(format5);

    // 6. Conditional formatting with histograms
    sheet = xlsx.addWorksheet("Histograms");
    sheet->write(1, 1, "Show color bars in the range [0, 1000]");
    for (int i=2; i<=11; ++i) {
        sheet->write(i, 1, QRandomGenerator::global()->bounded(0, 1000));
    }
    QXlsx::ConditionalFormatting format6;
    format6.setRange(QXlsx::CellRange(2, 1, 11, 1));
//    format6.addDataBarRule(QColor("#FFC6EFCE"), QXlsx::ConditionalFormatting::ValueObjectType::Min, "0",
//                           QXlsx::ConditionalFormatting::ValueObjectType::Max, "1000");
    format6.addDataBarRule(QColor(0xFFC6EFCE));
    sheet->addConditionalFormatting(format6);

    xlsx.saveAs("ConditionalFormatting.xlsx");
    return 0;
}
