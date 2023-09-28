// main.cpp

#include <iostream>
#include <vector>

#include <QtGlobal>
#include <QCoreApplication>
#include <QtCore>
#include <QVariant>
#include <QDir>
#include <QDebug>

#include <iostream>
using namespace std;

#include "xlsxdocument.h"
#include "xlsxworkbook.h"
using namespace QXlsx;

#include "fort.hpp" // libfort

int main(int argc, char *argv[])
{
    QCoreApplication app(argc, argv);

    if ( argc != 2 )
    {
        std::cout << "[Usage] ShowConsole *.xlsx" << std::endl;
        return 0;
    }

    // get xlsx file name
    QString xlsxFileName = argv[1];
    qDebug() << xlsxFileName;

    // load new xlsx using new document
    QXlsx::Document xlsxDoc(xlsxFileName);
    if (!xlsxDoc.isLoaded()) {
        qCritical() << "Failed to load" << xlsxFileName;
        return (-1); // failed to load
    }

    int sheetIndexNumber = 0;
    foreach (QString currentSheetName, xlsxDoc.sheetNames() )
    {
        QXlsx::AbstractSheet* currentSheet = xlsxDoc.sheet(currentSheetName);
        if (!currentSheet)
            continue;

        // get full cells of sheet

        currentSheet->workbook()->setActiveSheet( sheetIndexNumber );
        Worksheet* wsheet = (Worksheet*) currentSheet->workbook()->activeSheet();
        if (!wsheet)
            continue;

        // display sheet name
        std::cout
                << std::string( currentSheetName.toLocal8Bit() )
                << std::endl;

        int maxRow = wsheet->dimension().lastRow() - wsheet->dimension().firstRow();
        int maxCol = wsheet->dimension().lastColumn() - wsheet->dimension().firstColumn();

        fort::table fortTable;
        for (int row = 0; row < maxRow; row++)  {
            for (int col = 0; col < maxCol; col++) {
                auto val = wsheet->read(row + wsheet->dimension().firstRow(),
                                        col + wsheet->dimension().firstColumn()).toString();
                fortTable << std::string( val.toLocal8Bit() ); // display value
            }
            fortTable << fort::endr; // change to new row
        }

        std::cout << fortTable.to_string() << std::endl; // display forttable row

        sheetIndexNumber++;
    }


    return 0;
}
