#include <QCoreApplication>
#include <QVariant>
#include <QDir>
#include <QDebug>

#include "xlsxdocument.h"

using namespace QXlsx;

int main(int argc, char *argv[])
{
    QCoreApplication app(argc, argv);

    // Create xlsx document

    QXlsx::Document xlsxW;

    // The default sheet is a worksheet named "Sheet1". It is automatically available.

    //Rows and columns numbering starts with 1.
    int row = 1; int col = 1;

    QVariant writeValue = QString("Hello Qt!");
    xlsxW.write(row, col, writeValue); // write "Hello Qt!" to cell(A,1). it's shared string.

    if (xlsxW.saveAs("Test.xlsx")) // save the document as 'Test.xlsx'
    {
        qDebug() << "[debug] success to write xlsx file";
    }
    else
    {
        qDebug() << "[error] failed to write xlsx file";
        exit(-1);
    }

    qDebug() << "[debug] current directory is " << QDir::currentPath();

    // Read excel file (*.xlsx)

    Document xlsxR("Test.xlsx"); // Test.xlsx automatically loads in the constructor.
    if ( xlsxR.isLoaded() ) // returns whether excel file is successfully loaded
    {
        qDebug() << "[debug] successfully loaded xlsx file.";

        if (Cell* cell = xlsxR.activeWorksheet()->cell(row, col)) // get cell pointer.
        {
            QVariant var = cell->readValue(); // read cell value (number(double), QDateTime, QString ...)
            qDebug() << "[debug] cell(1,1) is " << var; // Display value. It is 'Hello Qt!'.
        }
        else
        {
            qDebug() << "[error] cell(1,1) is not set.";
            exit(-2);
        }
    }
    else
    {
        qDebug() << "[error] failed to load xlsx file.";
    }

    return 0;
}
