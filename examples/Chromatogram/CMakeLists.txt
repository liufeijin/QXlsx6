# CMakeLists.txt for Console Application

# Set minumum cmake version
cmake_minimum_required(VERSION 3.14)

# Set project name
project(Chromatogram LANGUAGES CXX)

# Set Your C++ version
set(CMAKE_CXX_STANDARD 17)

set(CMAKE_AUTOMOC ON)
set(CMAKE_AUTORCC ON)
set(CMAKE_AUTOUIC ON)

find_package(QT NAMES Qt6 Qt5 COMPONENTS Core Gui Widgets REQUIRED)
find_package(Qt${QT_VERSION_MAJOR} COMPONENTS Core Gui Widgets REQUIRED)

# NOTE: Here you can change path to QXlsx sources

set(QXLSX_PARENTPATH ${CMAKE_CURRENT_SOURCE_DIR}/../../QXlsx)
set(QXLSX_HEADERPATH ${QXLSX_PARENTPATH}/header)
set(QXLSX_SOURCEPATH ${QXLSX_PARENTPATH}/source)
# specify QXlsx subdirectory and where to build QXlsx
add_subdirectory(${QXLSX_PARENTPATH} ${QXLSX_PARENTPATH}/../build)

message("target is built in " ${CMAKE_BINARY_DIR})

#########################
# Console Application {{

add_executable(Chromatogram
    main.cpp 
    sat_calc.cpp
    sat_calc.h
    test.qrc
    mainwindow.cpp
    mainwindow.h
    mainwindow.ui
    )
target_link_libraries(Chromatogram
    Qt${QT_VERSION_MAJOR}::Gui
    Qt${QT_VERSION_MAJOR}::Widgets
    QXlsx::QXlsx
)

# Console Application }}
########################
