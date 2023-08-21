# CMakeLists.txt for Console Application

# TODO: Set minumum cmake version 
cmake_minimum_required(VERSION 3.14)

# TODO: Set project name 
project(ChartLineFill LANGUAGES CXX)

# TODO: Set Your C++ version
set(CMAKE_CXX_STANDARD 11) # C++ 17

find_package(QT NAMES Qt6 Qt5 COMPONENTS Core Gui REQUIRED)
find_package(Qt${QT_VERSION_MAJOR} COMPONENTS Core Gui REQUIRED)

add_subdirectory(QXlsx)

#########################
# Console Application {{

add_executable(CombinedChart main.cpp)
target_link_libraries(ChartLineFill PRIVATE QXlsx::QXlsx)

# Console Application }}
########################
