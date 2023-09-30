# CMakeLists.txt for Console Application

# Set minumum cmake version
cmake_minimum_required(VERSION 3.14)

# Set project name
project(Charts LANGUAGES CXX)

# Set Your C++ version
set(CMAKE_CXX_STANDARD 17)

find_package(QT NAMES Qt6 Qt5 COMPONENTS Core Gui REQUIRED)
find_package(Qt${QT_VERSION_MAJOR} COMPONENTS Core Gui REQUIRED)

# NOTE: Here you can change path to QXlsx sources

set(QXLSX_PARENTPATH ${CMAKE_CURRENT_SOURCE_DIR}/../../QXlsx)
set(QXLSX_HEADERPATH ${QXLSX_PARENTPATH}/header)
set(QXLSX_SOURCEPATH ${QXLSX_PARENTPATH}/source)
# specify QXlsx subdirectory and where to build QXlsx
add_subdirectory(${QXLSX_PARENTPATH} ${QXLSX_PARENTPATH}/../build)

message("target is built in " ${CMAKE_BINARY_DIR})

#########################
# Console Application {{

add_executable(Charts 
    main.cpp 
    chart.cpp 
    barchart.cpp
    chartextended.cpp
    chartlinefill.cpp
    )
target_link_libraries(Charts PRIVATE QXlsx::QXlsx)

# Console Application }}
########################
