// calendar.cpp

#include <QtGlobal>
#include <QtCore>
#include <QDebug>

#include "xlsxdocument.h"
#include "xlsxformat.h"
#include "xlsxcellrange.h"
#include "xlsxworksheet.h"

QXLSX_USE_NAMESPACE

int main()
{
    Document xlsx;
    QDate today(QDate::currentDate());
    for (int month=1; month<=12; ++month)
    {
        auto sheet = xlsx.addWorksheet(QLocale().monthName(month));
        sheet->setGridLinesVisible(false);

        //the header row
        Format headerStyle;
        headerStyle.setFontSize(48);
        headerStyle.setFontColor(Qt::darkBlue);
        headerStyle.setHorizontalAlignment(Format::AlignHCenter);
        headerStyle.setVerticalAlignment(Format::AlignVCenter);
        sheet->setRowHeight(1, 80);
        sheet->write("A1", QString("%1 %2").arg(QLocale().monthName(month)).arg(today.year()));
        sheet->mergeCells("A1:N1", headerStyle);

        //header with month titles
        for (int day=1; day<=7; ++day) {
            Format monthStyle;
            monthStyle.setFontSize(12);
            monthStyle.setFontColor(Qt::white);
            monthStyle.setFontBold(true);
            monthStyle.setHorizontalAlignment(Format::AlignHCenter);
            monthStyle.setVerticalAlignment(Format::AlignVCenter);
            monthStyle.setFillPattern(Format::PatternSolid);
            monthStyle.setPatternBackgroundColor(Qt::darkBlue);

            sheet->setColumnWidth(day*2-1, 5);
            sheet->setColumnWidth(day*2, 13);
            sheet->write(2, day*2-1, QLocale().dayName(day));
            sheet->mergeCells(CellRange(2, day*2-1, 2, day*2), monthStyle);
        }

        QColor borderColor = QColor(Qt::gray);

        Format weekendLeftStyle;
        weekendLeftStyle.setFontSize(14);
        weekendLeftStyle.setFontBold(true);
        weekendLeftStyle.setHorizontalAlignment(Format::AlignLeft);
        weekendLeftStyle.setVerticalAlignment(Format::AlignTop);
        weekendLeftStyle.setPatternBackgroundColor(QColor("#93CCEA"));
        weekendLeftStyle.setLeftBorderStyle(Format::BorderThin);
        weekendLeftStyle.setLeftBorderColor(borderColor);
        weekendLeftStyle.setBottomBorderStyle(Format::BorderThin);
        weekendLeftStyle.setBottomBorderColor(borderColor);

        Format weekendRightStyle;
        weekendRightStyle.setHorizontalAlignment(Format::AlignHCenter);
        weekendRightStyle.setVerticalAlignment(Format::AlignTop);
        weekendRightStyle.setPatternBackgroundColor(QColor("#93CCEA"));
        weekendRightStyle.setRightBorderStyle(Format::BorderThin);
        weekendRightStyle.setRightBorderColor(borderColor);
        weekendRightStyle.setBottomBorderStyle(Format::BorderThin);
        weekendRightStyle.setBottomBorderColor(borderColor);

        Format holidayLeftStyle;
        holidayLeftStyle.setFontSize(14);
        holidayLeftStyle.setFontBold(true);
        holidayLeftStyle.setHorizontalAlignment(Format::AlignLeft);
        holidayLeftStyle.setVerticalAlignment(Format::AlignTop);
        holidayLeftStyle.setPatternBackgroundColor(QColor("#EACC93"));
        holidayLeftStyle.setLeftBorderStyle(Format::BorderThin);
        holidayLeftStyle.setLeftBorderColor(borderColor);
        holidayLeftStyle.setBottomBorderStyle(Format::BorderThin);
        holidayLeftStyle.setBottomBorderColor(borderColor);

        Format holidayRightStyle;
        holidayRightStyle.setHorizontalAlignment(Format::AlignHCenter);
        holidayRightStyle.setVerticalAlignment(Format::AlignTop);
        holidayRightStyle.setPatternBackgroundColor(QColor("#EACC93"));
        holidayRightStyle.setRightBorderStyle(Format::BorderThin);
        holidayRightStyle.setRightBorderColor(borderColor);
        holidayRightStyle.setBottomBorderStyle(Format::BorderThin);
        holidayRightStyle.setBottomBorderColor(borderColor);

        Format workdayLeftStyle;
        workdayLeftStyle.setHorizontalAlignment(Format::AlignLeft);
        workdayLeftStyle.setVerticalAlignment(Format::AlignTop);
        workdayLeftStyle.setPatternBackgroundColor(Qt::white);
        workdayLeftStyle.setLeftBorderStyle(Format::BorderThin);
        workdayLeftStyle.setLeftBorderColor(borderColor);
        workdayLeftStyle.setBottomBorderStyle(Format::BorderThin);
        workdayLeftStyle.setBottomBorderColor(borderColor);

        Format workdayRightStyle;
        workdayRightStyle.setHorizontalAlignment(Format::AlignHCenter);
        workdayRightStyle.setVerticalAlignment(Format::AlignTop);
        workdayRightStyle.setPatternBackgroundColor(Qt::white);
        workdayRightStyle.setRightBorderStyle(Format::BorderThin);
        workdayRightStyle.setRightBorderColor(borderColor);
        workdayRightStyle.setBottomBorderStyle(Format::BorderThin);
        workdayRightStyle.setBottomBorderColor(borderColor);

        Format greyLeftStyle;
        greyLeftStyle.setPatternBackgroundColor(Qt::lightGray);
        greyLeftStyle.setLeftBorderStyle(Format::BorderThin);
        greyLeftStyle.setLeftBorderColor(borderColor);
        greyLeftStyle.setBottomBorderStyle(Format::BorderThin);
        greyLeftStyle.setBottomBorderColor(borderColor);

        Format greyRightStyle;
        greyRightStyle.setPatternBackgroundColor(Qt::lightGray);
        greyRightStyle.setRightBorderStyle(Format::BorderThin);
        greyRightStyle.setRightBorderColor(borderColor);
        greyRightStyle.setBottomBorderStyle(Format::BorderThin);
        greyRightStyle.setBottomBorderColor(borderColor);

        int rownum = 3;
        for (int day=1; day<=31; ++day) {
            QDate date(today.year(), month, day);
            if (!date.isValid())
                break;
            sheet->setRowHeight(rownum, 100);
            int dow = date.dayOfWeek(); // 1 = Monday to 7 = Sunday
            int colnum = dow*2-1;

            if (dow <= 5) { // 1,2,3,4,5
                sheet->write(rownum, colnum, day, workdayLeftStyle);
                sheet->write(rownum, colnum+1, QVariant(), workdayRightStyle);
            } else if (dow == 6) {
                sheet->write(rownum, colnum, day, weekendLeftStyle);
                sheet->write(rownum, colnum+1, QVariant(), weekendRightStyle);
            } else {
                sheet->write(rownum, colnum, day, holidayLeftStyle);
                sheet->write(rownum, colnum + 1, QVariant(), holidayRightStyle);
            }

            if (day == 1 && dow != 1) {//First day
                for (int i=1; i<dow; ++i) {
                    sheet->write(rownum, i*2-1, QVariant(), greyLeftStyle);
                    sheet->write(rownum, i*2, QVariant(), greyRightStyle);
                }
            } else if (day == date.daysInMonth() && dow != 7) {//Last day
                for (int i=dow+1; i<=7; ++i) {
                    sheet->write(rownum, i*2-1, QVariant(), greyLeftStyle);
                    sheet->write(rownum, i*2, QVariant(), greyRightStyle);
                }
            }

            if (dow == 7)
                rownum++;
        }

    }

    xlsx.saveAs("calendar1.xlsx");
    qDebug()<<"**** end of main() ****";
    return 0;
}
