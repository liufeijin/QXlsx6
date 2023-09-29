# CMakeLists.txt for Console Application

# Set minumum cmake version
cmake_minimum_required(VERSION 3.14)

# Set project name
project(Sheets LANGUAGES CXX)

# Set Your C++ version
set(CMAKE_CXX_STANDARD 17)

find_package(QT NAMES Qt6 Qt5 COMPONENTS Core Gui REQUIRED)
find_package(Qt${QT_VERSION_MAJOR} COMPONENTS Core Gui REQUIRED)

add_subdirectory(../../QXlsx ../../build/QXlsx)

message("target is built in " ${CMAKE_BINARY_DIR})

#########################
# Console Application {{

add_executable(Sheets 
    main.cpp 
    pageMargins.cpp 
    )
target_link_libraries(Sheets PRIVATE QXlsx::QXlsx)

# Console Application }}
########################
