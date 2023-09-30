
# How to setup QXlsx project

Here's an easy way to add QXlsx to your project. 

This works only with qmake, for cmake look at [the other doc](HowToSetProject-cmake.md).

As QXlsx uses private Qt classes, it is best to statically link QXlsx with your 
program. Steps below describe adding QXlsx to a qmake project directly (statically).

## Steps to set

:one: Clone source code from github

```sh
git clone https://github.com/alexnovichkov/QXlsx.git
```

:two: Add QXlsx to your Qt project (*.pro)

```qmake
QXLSX_PARENTPATH=path_to_QXlsx_QXlsx_directory/
QXLSX_HEADERPATH=path_to_QXlsx_QXlsx_directory/header/ # should be path to QXlsx/header directory
QXLSX_SOURCEPATH=path_to_QXlsx_QXlsx_directory/source/ # should be path to QXlsx/source directory
include($${QXLSX_PARENTPATH}/QXlsx.pri)
```

QXlsx.pri automatically adds dependencies on `core` and `gui-private` modules.

:three: Add header file(s). Here is the sample HelloWorld project

```cpp
// main.cpp

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

        if (Cell* cell = xlsxR.cellAt(row, col)) // get cell pointer.
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

```

:four: Build and run the project. An xlsx file will be generated in the current 
build directory.

