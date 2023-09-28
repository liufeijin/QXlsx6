// QXlsx
// MIT License
// https://github.com/j2doll/QXlsx

// main.cpp

#include <QtGlobal>
#include <QObject>
#include <QString>
#include <QUrl>
#include <QList>
#include <QVariant>

#include <QDebug>

#include <QGuiApplication>
#include <QQmlApplicationEngine>
#include <QQmlContext>

#include <cstdio>
#include <iostream>

#include "xlsxdocument.h"
#include "xlsxworkbook.h"
using namespace QXlsx;

#include "XlsxTableModel.h"
#include "DynArray2D.h"

std::string convertFromNumberToExcelColumn(int n);

int main(int argc, char *argv[])
{
#if QT_VERSION < QT_VERSION_CHECK(6,0,0)
    QCoreApplication::setAttribute( Qt::AA_EnableHighDpiScaling );
#endif
    QGuiApplication app(argc, argv);

    QQmlApplicationEngine engine;
    QQmlContext* ctxt = engine.rootContext();

    QXlsx::Document xlsx( ":/test.xlsx" ); // load xlsx
    if (!xlsx.isLoaded())
    {
        qDebug() << "[ERROR] Failed to load xlsx";
        return (-1);
    }

    QList<QString> colTitle; // list of column title
    Worksheet* wsheet = (Worksheet*) xlsx.workbook()->activeSheet();
    if (!wsheet )
    {
        qDebug() << "[ERROR] Failed to get active sheet";
        return (-2);
    }

    int maxRow = wsheet->dimension().lastRow() - wsheet->dimension().firstRow();
    int maxCol = wsheet->dimension().lastColumn() - wsheet->dimension().firstColumn();

    for (int ic = 0 ; ic < maxCol ; ic++)
    {
        std::string strCol = convertFromNumberToExcelColumn(ic + 1);
        QString colName = QString::fromStdString( strCol );
        colTitle.append( colName );
    }

    // make cell matrix that has (colMax x rowMax) size.

    DynArray2D< std::string > dynIntArray(maxCol, maxRow);

    for (int row = 0; row < maxRow; row++)  {
        for (int col = 0; col < maxCol; col++) {
            auto val = wsheet->read(row + wsheet->dimension().firstRow(),
                                    col + wsheet->dimension().firstColumn()).toString();
            // set string value to (col, row)
            dynIntArray.setValue( col, row, val.toStdString() );
        }
    }

    QList<VLIST> xlsxData;
    for (int ir = 0; ir < maxRow; ir++ )
    {
        VLIST vl;
        for (int ic = 0; ic < maxCol; ic++)
        {
            std::string value = dynIntArray.getValue( ic, ir );
            vl.append( QString::fromStdString(value) );
        }
        xlsxData.append(vl);
    }

    // set model for tableview
    XlsxTableModel xlsxTableModel(colTitle, xlsxData);
    ctxt->setContextProperty( "xlsxModel", &xlsxTableModel );

    engine.load( QUrl(QStringLiteral("qrc:/main.qml")) ); // load QML
    if ( engine.rootObjects().isEmpty() )
    {
        qDebug() << "Failed to load qml";
        return (-1);
    }

    int ret = app.exec();
    return ret;
}

std::string convertFromNumberToExcelColumn(int n)
{
    // main code from https://www.geeksforgeeks.org/find-excel-column-name-given-number/
    // Function to print Excel column name for a given column number

    std::string stdString;

    char str[1000]; // To store result (Excel column name)
    int i = 0; // To store current index in str which is result

    while ( n > 0 )
    {
        // Find remainder
        int rem = n % 26;

        // If remainder is 0, then a 'Z' must be there in output
        if ( rem == 0 )
        {
            str[i++] = 'Z';
            n = (n/26) - 1;
        }
        else // If remainder is non-zero
        {
            str[i++] = (rem-1) + 'A';
            n = n / 26;
        }
    }
    str[i] = '\0';

    // Reverse the string and print result
    std::reverse( str, str + strlen(str) );

    stdString = str;
    return stdString;
}
