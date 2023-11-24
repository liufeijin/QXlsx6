# QXlsx Examples

## [HelloWorld](HelloWorld)

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

## [Demo](Demo)

Demonstrates basic operations in worksheets:

- how to format cells:
    - text alignment
    - cell borders
    - cell fills
    - text fonts
    - cell number formats
- how to merge, group and select cells in a worksheet
- how to write formulas
    - basic functions
    - array formulas
    - shared formulas

## [Types](Types)

Demonstrates various numeric types that can be written and read.

This example shows that `Worksheet::read()` method returns values with respect to their type, whereas Cell::value() returns data as it is stored in cells (for numeric data it is usually `double`). Compare:

```cpp 
Document doc;
doc.write( "A3", QVariant(QDate(2019, 10, 9)) );
qDebug() << doc.read(3, 1).type() << doc.read(3, 1); // QVariant::QDate QVariant(QDate, QDate("2019-10-09"))
qDebug() << doc.activeWorksheet()->cellAt(3, 1)->value().type() << doc.cellAt(3, 1)->value(); // QVariant::double QVariant(double, 43747)
```



## [DefinedNames](DefinedNames)

Demonstrates how to add defined names to the workbook and use them in formulas.

Defined names are descriptive names to represent cells, ranges of cells, formulas, or constant values. Defined names can be used to represent a range on any worksheet.

Excerpt from [definedNames.cpp](DefinedNames/definedNames.cpp)

```cpp
...
Document xlsx;
xlsx.addDefinedName("MyCol_1", "=Sheet1!$A$1:$A$10");
xlsx.write(11, 1, "=SUM(MyCol_1)");
...
```

## [DataValidation](DataValidation)

Demonstrates how to add validation to worksheet data.

Data validation is used to specify constraints on the data that can be entered into a cell.

## [Calendar](Calendar)

Demonstrates how to create a workbook with the current year calendar.

![](../markdown.data/calendar.png)

## [Autofilter](Autofilter)

Demonstrates how to use autofiltering and sorting features in a worksheet.

:warning: Excel understands autofilters written with QXlsx, but it doesn't automatically hide the filtered-out rows. So for each worksheet press Data -> Filter -> Reapply.

:arrow_right: TO ADD IN FUTURE RELEASES: when adding autofilter, manually hide rows. It requires reading cell values and actually is not an **autofilter** anymore.

## [ConditionalFormatting](ConditionalFormatting)

Demonstrates how to add formatting to cells based on their values.

## [Charts](Charts)

Demonstrates adding various charts to a worksheet.

- [chart.cpp](Charts/chart.cpp) demonstrates how to add charts of different types.
- [chertextended.cpp](Charts/chartextended.cpp) demonstrates how to change title and gridlines of a chart.
- [barchart.cpp](Charts/barchart.cpp) demonstrates various bar chart styles.
- [chartlinefill.cpp](Charts/chartlinefill.cpp) demonstrates how to change lines and fills in charts.

## [CombinedChart](CombinedChart)

Demonstrates how to add bar series and line series on the same chart. Demonstrates how to move one of the series to the right axis.

![](../markdown.data/combinedchart.png)

## [Chartsheets](Chartsheets)

Demonstrates how to add and copy chartsheets, how to use picture fills.

## [Sheets](Sheets)

Demonstrates how to set up sheet parameters such as page properties, print properties, view properties etc.

## [Chromatogram](Chromatogram)

Demonstrates a simple GUI application that uses QXlsx to generate a report based on a file data.

## [TestExcel](TestExcel)

:zap: Basic examples (based on QtXlsx examples)
- [document property](TestExcel/documentproperty.cpp) - tests adding document properties
- [hyperlink](TestExcel/hyperlinks.cpp) - tests writing hyperlinks of different kinds.
- [image](TestExcel/image.cpp) - tests adding, reading and removing images from a worksheet
- [read style](TestExcel/readStyle.cpp) - needs rewriting as now it does nothing useful
- [richtext](TestExcel/richtext.cpp)
- [row column](TestExcel/rowcolumn.cpp)
- [style](TestExcel/style.cpp)
- [worksheet operations](TestExcel/worksheetoperations.cpp)
- [extList](TestExcel/extList.cpp) - tests reading Excel extensions from an xlsx file.

![](../markdown.data/testexcel.png)

## [SheetProtection](SheetProtection)

Demonstrates how to protect a worksheet with a password and how to check the password with a user-supplied input.

## [ExcelReading](ExcelReading)

Demonstrates that loading xlsx file that was created in Excel and writing it back creates almost identical xlsx file.

The file chartsheet1.xlsx contains complex shape and line formatting, data labels, data table, Excel extension data etc. 

![](../markdown.data/ExcelReading.png)

## [HelloAndroid](HelloAndroid)

Demonstrates how to create a QML application that uses QXlsx to fill the model.

This example needs rewriting as it compiles only with Qt5.

- See 'HelloAndroid' example using QML and native C++.

- Qt 5.11.1 / gcc 4.9 / QtCreator 4.6.2 
- Android x86 (using Emulator <Android Oreo / API 26>)
- Android Studio 3.1.3 (Android NDK 17.1)

![](../markdown.data/android.jpg)

## [Pump](Pump)

Serves as a test application that pumps xlsx files through QXlsx (reads files and writes them back) to test QXlsx features.

Originally the example was a part of [libxlsxwriter](https://github.com/jmcnamara/libxlsxwriter).

## [WebServer](WebServer)

Loads xlsx file and displays on Web.

- Connect to `http://127.0.0.1:3001` 
- C++ 14 is required. Old compilers are not supported.

This example needs rewriting.

![](../markdown.data/webserver.png)

## [ShowConsole](ShowConsole)

Loads xlsx file and displays data in console.

Usage: 

```bat
ShowConsole *.xlsx
```

![](../markdown.data/show-console.jpg)

## XlsxFactory 

- Load xlsx file and display on Qt widgets. 
- Moved to personal repository for advanced app.
	- https://j2doll.tistory.com/654
	- The source code of this program cannot be released because it contains a commercial license.

![](../markdown.data/copycat.png)
![](../markdown.data/copycat2.jpg)
