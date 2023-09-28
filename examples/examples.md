# QXlsx Examples

## [HelloWorld](https://github.com/QtExcel/QXlsx/tree/master/examples/HelloWorld)

- Hello world example

```cpp
// main.cpp

#include <QtGlobal>
#include <QCoreApplication>
#include <QtCore>
#include <QVariant>
#include <QDebug>

#include <iostream>
using namespace std;

// [0] include QXlsx headers 
#include "xlsxdocument.h"
#include "xlsxchartsheet.h"
#include "xlsxcellrange.h"
#include "xlsxchart.h"
#include "xlsxrichstring.h"
#include "xlsxworkbook.h"
using namespace QXlsx;

int main(int argc, char *argv[])
{
    QCoreApplication app(argc, argv);

    int row = 1; int col = 1;
	
    // [1]  Writing excel file(*.xlsx)
    QXlsx::Document xlsxW;
	QVariant writeValue = QString("Hello Qt!");
    xlsxW.write(row, col, writeValue); // write "Hello Qt!" to cell(A,1).
    xlsxW.saveAs("Test.xlsx"); // save the document as 'Test.xlsx'

    // [2] Reading excel file(*.xlsx)
    Document xlsxR("Test.xlsx"); 
    if (xlsxR.load()) // load excel file
    { 
        Cell* cell = xlsxR.cellAt(row, col); // get cell pointer.
        if ( cell != NULL )
        {
            QVariant var = cell->readValue(); // read cell value (number(double), QDateTime, QString ...)
            qDebug() << var; // display value. it is 'Hello Qt!'.
        }
    }

    return 0;
}
```

## [TestExcel](https://github.com/QtExcel/QXlsx/tree/master/examples/TestExcel)

- :zap: Basic examples (based on QtXlsx examples)
    - [calendar](https://github.com/QtExcel/QXlsx/blob/master/examples/TestExcel/calendar.cpp)
    - [chart](https://github.com/QtExcel/QXlsx/blob/master/examples/TestExcel/chart.cpp)
    - [chart sheet](https://github.com/QtExcel/QXlsx/blob/master/examples/TestExcel/chartsheet.cpp) - tests adding chartsheets, copying charsheets, using picture fills.
    - [data validation](https://github.com/QtExcel/QXlsx/blob/master/examples/TestExcel/datavalidation.cpp)
    - [demo](https://github.com/QtExcel/QXlsx/blob/master/TestExcel/examples/demo.cpp)
    - [document property](https://github.com/QtExcel/QXlsx/blob/master/examples/TestExcel/documentproperty.cpp)
    - [extract data](https://github.com/QtExcel/QXlsx/blob/master/examples/TestExcel/extractdata.cpp)
    - [formula](https://github.com/QtExcel/QXlsx/blob/master/examples/TestExcel/formulas.cpp)
    - [hyperlink](https://github.com/QtExcel/QXlsx/blob/master/examples/TestExcel/hyperlinks.cpp)
    - [image](https://github.com/QtExcel/QXlsx/blob/master/examples/TestExcel/image.cpp)
    - [merge cells](https://github.com/QtExcel/QXlsx/blob/master/examples/TestExcel/mergecells.cpp)
    - [number format](https://github.com/QtExcel/QXlsx/blob/master/examples/TestExcel/numberformat.cpp)
    - [page margins](https://github.com/QtExcel/QXlsx/blob/master/examples/TestExcel/pageMargins.cpp)
    - [read style](https://github.com/QtExcel/QXlsx/blob/master/examples/TestExcel/readStyle.cpp)
    - [richtext](https://github.com/QtExcel/QXlsx/blob/master/examples/TestExcel/richtext.cpp)
    - [row column](https://github.com/QtExcel/QXlsx/blob/master/examples/TestExcel/rowcolumn.cpp)
    - [style](https://github.com/QtExcel/QXlsx/blob/master/examples/TestExcel/style.cpp)
    - [worksheet operations](https://github.com/QtExcel/QXlsx/blob/master/TestExcel/examples/worksheetoperations.cpp)

![](../markdown.data/testexcel.png)

## [ExcelReading](https://github.com/QtExcel/QXlsx/tree/master/examples/ExcelReading)

Loads xlsx file that was created in Excel and writes it back.

The file chartsheet1.xlsx contains complex shape and line formatting, data labels, data table, Excel extension data etc.

## [CombinedChart](https://github.com/QtExcel/QXlsx/tree/master/examples/CombinedChart)

Demonstrates how to add bar series and line series on the same chart. Demonstrates how to move one of the series to the right axis.

![](../markdown.data/combinedchart.png)

## [ChartLineFill](https://github.com/QtExcel/QXlsx/tree/master/examples/ChartLineFill)

Demonstrates various line parameters.

## [HelloAndroid](https://github.com/QtExcel/QXlsx/tree/master/examples/HelloAndroid)

- See 'HelloAndroid' example using QML and native C++.

- Qt 5.11.1 / gcc 4.9 / QtCreator 4.6.2 
- Android x86 (using Emulator <Android Oreo / API 26>)
- Android Studio 3.1.3 (Android NDK 17.1)

![](../markdown.data/android.jpg)

## [WebServer](https://github.com/QtExcel/QXlsx/tree/master/examples/WebServer)
- Loads xlsx file and displays on Web.
	- Connect to `http://127.0.0.1:3001` 
- C++ 14(17) is required. Old compilers are not supported.

![](../markdown.data/webserver.png)

## [ShowConsole](https://github.com/QtExcel/QXlsx/tree/master/examples/ShowConsole)
- Loads xlsx file and displays in console.
  - [Usage] ShowConsole *.xlsx
- C++ 11 is required. Old compilers is not supported.

![](../markdown.data/show-console.jpg)

## [ReadColor](https://github.com/QtExcel/QXlsx/tree/master/examples/ReadColor) 
- Read cell color

![](../markdown.data/read-color.jpg)

## XlsxFactory 
- Load xlsx file and display on Qt widgets. 
- Moved to personal repository for advanced app.
	- https://j2doll.tistory.com/654
	- The source code of this program cannot be released because it contains a commercial license.

![](../markdown.data/copycat.png)
![](../markdown.data/copycat2.jpg)
