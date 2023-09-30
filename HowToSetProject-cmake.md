# How to setup QXlsx project

Here's an easy way to add QXlsx to your project. 

This works only with cmake, for qmake look at [the other doc](HowToSetProject.md).

As QXlsx uses private Qt classes, it is best to statically link QXlsx with your 
program. Steps below describe adding QXlsx to a cmake project directly (statically).

## Steps to set

:one: Clone source code from github

```sh
git clone https://github.com/alexnovichkov/QXlsx.git
```

:two: Add QXlsx to your Qt project (CMakeLists.txt)

```cmake
set(QXLSX_PARENTPATH path_to_QXlsx_QXlsx_dir)
set(QXLSX_HEADERPATH ${QXLSX_PARENTPATH}/header)
set(QXLSX_SOURCEPATH ${QXLSX_PARENTPATH}/source)
# specify QXlsx subdirectory and where to build QXlsx
add_subdirectory(${QXLSX_PARENTPATH} ${QXLSX_PARENTPATH}/../build)
```

CMakeLists.txt automatically adds dependencies on `Core` and `GuiPrivate` modules.

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


## Adding QXlsx with cmake FetchContent

Add these lines to your CMakeLists.txt:

```cmake
FetchContent_Declare(
  QXlsx
  GIT_REPOSITORY https://github.com/alexnovichkov/QXlsx.git
  GIT_TAG        sha-of-the-commit
  SOURCE_SUBDIR  QXlsx
)
FetchContent_MakeAvailable(QXlsx)
target_link_libraries(myapp PRIVATE QXlsx::QXlsx)
```

## Installing QXlsx locally

If needed, you can install QXlsx to later use. 

:one: Go to QXlsx directory and build QXlsx:

```sh
mkdir build
cd build
cmake ../QXlsx/ -DCMAKE_INSTALL_PREFIX=... -DCMAKE_BUILD_TYPE=Release
cmake --build .
cmake --install .
```

:two: Add QXlsx to your project:

```cmake
find_package(QXlsxQt5 REQUIRED) # or QXlsxQt6
target_link_libraries(myapp PRIVATE QXlsx::QXlsx)
```

:three: Build and run the project. An xlsx file will be generated in the current 
build directory.
