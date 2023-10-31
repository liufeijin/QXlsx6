# CMakeLists.txt for QXlsx Library

cmake_minimum_required(VERSION 3.16)

project(QXlsx
    VERSION 1.4.6
    LANGUAGES CXX
)

set(CMAKE_INCLUDE_CURRENT_DIR ON)

set(CMAKE_AUTOMOC ON)

include(GNUInstallDirs)

if(NOT DEFINED QT_VERSION_MAJOR)
    find_package(QT NAMES Qt6 Qt5 COMPONENTS Core Gui REQUIRED)
endif()
find_package(Qt${QT_VERSION_MAJOR} COMPONENTS Core Gui REQUIRED)
set(EXPORT_NAME QXlsxQt${QT_VERSION_MAJOR})

# the standard required to build the lib
set(CMAKE_CXX_STANDARD 17 CACHE STRING "")
set(CMAKE_CXX_STANDARD_REQUIRED ON)

set(CMAKE_MODULE_PATH ${CMAKE_CURRENT_SOURCE_DIR}/cmake/modules)

if(NOT DEFINED ${QXLSX_PARENTPATH})
    set(QXLSX_PARENTPATH ${CMAKE_CURRENT_SOURCE_DIR}/..)
endif(NOT DEFINED ${QXLSX_PARENTPATH})

if(NOT DEFINED ${QXLSX_HEADERPATH})
    set(QXLSX_HEADERPATH ${QXLSX_PARENTPATH}/header)
endif(NOT DEFINED ${QXLSX_HEADERPATH})

if(NOT DEFINED ${QXLSX_SOURCEPATH})
    set(QXLSX_SOURCEPATH ${QXLSX_PARENTPATH}/source)
endif(NOT DEFINED ${QXLSX_SOURCEPATH})

message("Current paths of QXlsx: ")
message(${QXLSX_PARENTPATH})
message(${QXLSX_HEADERPATH})
message(${QXLSX_SOURCEPATH})
message("QXlsx is built in " ${CMAKE_CURRENT_BINARY_DIR})

# Due to historical reasons this value is kept off
option(BUILD_SHARED_LIBS "Build in shared lib mode" OFF)

set(SRC_FILES
    source/xlsxcellreference.cpp
    source/xlsxchart.cpp
    source/xlsxchartsheet.cpp
    source/xlsxcolor.cpp
    source/xlsxconditionalformatting.cpp
    source/xlsxcontenttypes.cpp
    source/xlsxdatavalidation.cpp
    source/xlsxdatetype.cpp
    source/xlsxdocpropsapp.cpp
    source/xlsxdocpropscore.cpp
    source/xlsxdocument.cpp
    source/xlsxdrawing.cpp
    source/xlsxdrawinganchor.cpp
    source/xlsxeffect.cpp
    source/xlsxfillformat.cpp
    source/xlsxfont.cpp
    source/xlsxformat.cpp
    source/xlsxlabel.cpp
    source/xlsxlayout.cpp
    source/xlsxlegend.cpp
    source/xlsxlineformat.cpp
    source/xlsxmain.cpp
    source/xlsxmarkerformat.cpp
    source/xlsxmediafile.cpp
    source/xlsxnumformatparser.cpp
    source/xlsxrelationships.cpp
    source/xlsxrichstring.cpp
    source/xlsxseries.cpp
    source/xlsxshapeformat.cpp
    source/xlsxsharedstrings.cpp
    source/xlsxsheetview.cpp
    source/xlsxsimpleooxmlfile.cpp
    source/xlsxstyles.cpp
    source/xlsxtext.cpp
    source/xlsxtheme.cpp
    source/xlsxtitle.cpp
    source/xlsxutility.cpp
    source/xlsxworkbook.cpp
    source/xlsxworksheet.cpp
    source/xlsxzipreader.cpp
    source/xlsxzipwriter.cpp
    source/xlsxabstractooxmlfile.cpp
    source/xlsxabstractsheet.cpp
    source/xlsxautofilter.cpp
    source/xlsxaxis.cpp
    source/xlsxcell.cpp
    source/xlsxcellformula.cpp
    source/xlsxcellrange.cpp
    header/xlsxconditionalformatting_p.h
    header/xlsxcontenttypes_p.h
    header/xlsxdatavalidation.h
    header/xlsxdatavalidation_p.h
    header/xlsxdatetype.h
    header/xlsxdocpropsapp_p.h
    header/xlsxdocpropscore_p.h
    header/xlsxdocument.h
    header/xlsxdocument_p.h
    header/xlsxdrawing_p.h
    header/xlsxdrawinganchor_p.h
    header/xlsxeffect.h
    header/xlsxfillformat.h
    header/xlsxfont.h
    header/xlsxformat.h
    header/xlsxformat_p.h
    header/xlsxglobal.h
    header/xlsxlabel.h
    header/xlsxlayout.h
    header/xlsxlegend.h
    header/xlsxlineformat.h
    header/xlsxmain.h
    header/xlsxmarkerformat.h
    header/xlsxmediafile_p.h
    header/xlsxnumformatparser_p.h
    header/xlsxrelationships_p.h
    header/xlsxrichstring.h
    header/xlsxrichstring_p.h
    header/xlsxseries.h
    header/xlsxshapeformat.h
    header/xlsxsharedstrings_p.h
    header/xlsxsheetview.h
    header/xlsxsimpleooxmlfile_p.h
    header/xlsxstyles_p.h
    header/xlsxtext.h
    header/xlsxtheme_p.h
    header/xlsxtitle.h
    header/xlsxutility_p.h
    header/xlsxworkbook.h
    header/xlsxworkbook_p.h
    header/xlsxworksheet.h
    header/xlsxworksheet_p.h
    header/xlsxzipreader_p.h
    header/xlsxzipwriter_p.h
    header/xlsxabstractooxmlfile.h
    header/xlsxabstractooxmlfile_p.h
    header/xlsxabstractsheet.h
    header/xlsxabstractsheet_p.h
    header/xlsxautofilter.h
    header/xlsxaxis.h
    header/xlsxcell.h
    header/xlsxcell_p.h
    header/xlsxcellformula.h
    header/xlsxcellformula_p.h
    header/xlsxcellrange.h
    header/xlsxcellreference.h
    header/xlsxchart.h
    header/xlsxchartsheet.h
    header/xlsxchartsheet_p.h
    header/xlsxcolor.h
    header/xlsxcolor_p.h
    header/xlsxconditionalformatting.h
    header/xlsxsheetprotection.h
    source/xlsxsheetprotection.cpp
)

set(QXLSX_PUBLIC_HEADERS
    header/xlsxautofilter.h
    header/xlsxsheetview.h
    header/xlsxlegend.h
    header/xlsxlayout.h
    header/xlsxseries.h
    header/xlsxeffect.h
    header/xlsxlabel.h
    header/xlsxfont.h
    header/xlsxtext.h
    header/xlsxtitle.h
    header/xlsxfillformat.h
    header/xlsxmain.h
    header/xlsxshapeformat.h
    header/xlsxaxis.h
    header/xlsxmarkerformat.h
    header/xlsxlineformat.h
    header/xlsxabstractooxmlfile.h
    header/xlsxabstractsheet.h
    header/xlsxabstractsheet_p.h
    header/xlsxcellformula.h
    header/xlsxcell.h
    header/xlsxcellrange.h
    header/xlsxcellreference.h
    header/xlsxchart.h
    header/xlsxchartsheet.h
    header/xlsxconditionalformatting.h
    header/xlsxdatavalidation.h
    header/xlsxdatetype.h
    header/xlsxdocument.h
    header/xlsxformat.h
    header/xlsxglobal.h
    header/xlsxrichstring.h
    header/xlsxworkbook.h
    header/xlsxworksheet.h
    header/xlsxsheetprotection.h
)

add_library(QXlsx
    ${SRC_FILES}
    ${QXLSX_PUBLIC_HEADERS}
)

add_library(QXlsx::QXlsx ALIAS QXlsx)

target_compile_definitions(QXlsx PRIVATE
    QT_NO_KEYWORDS
    QT_NO_URL_CAST_FROM_STRING
    QT_NO_CAST_FROM_BYTEARRAY
    QT_USE_QSTRINGBUILDER
    QT_NO_SIGNALS_SLOTS_KEYWORDS
    QT_USE_FAST_OPERATOR_PLUS
    QT_DISABLE_DEPRECATED_BEFORE=0x060200
)

if (NOT WIN32)
    # Strict iterators can't be used on Windows, they lead to a link error
    # when application code iterates over a QVector<QPoint> for instance, unless
    # Qt itself was also built with strict iterators.
    # See example at https://bugreports.qt.io/browse/AUTOSUITE-946
    target_compile_definitions(QXlsx PRIVATE QT_STRICT_ITERATORS)
endif()

target_compile_features(QXlsx INTERFACE cxx_std_17)

if (BUILD_SHARED_LIBS)
    target_compile_definitions(QXlsx PUBLIC QXlsx_SHAREDLIB)
endif()

target_link_libraries(${PROJECT_NAME}
   Qt${QT_VERSION_MAJOR}::Core
   Qt${QT_VERSION_MAJOR}::GuiPrivate
)

target_include_directories(QXlsx
PRIVATE
    ${QXLSX_HEADERPATH}
PUBLIC
    $<BUILD_INTERFACE:${CMAKE_CURRENT_SOURCE_DIR}/header>
    $<INSTALL_INTERFACE:include/QXlsxQt${QT_VERSION_MAJOR}>
)

set_target_properties(QXlsx PROPERTIES
    OUTPUT_NAME ${EXPORT_NAME}
    VERSION ${PROJECT_VERSION}
    SOVERSION 0.${PROJECT_VERSION}
    PUBLIC_HEADER "${QXLSX_PUBLIC_HEADERS}"
)

install(TARGETS QXlsx
    EXPORT ${EXPORT_NAME}Targets DESTINATION ${CMAKE_INSTALL_LIBDIR}
    RUNTIME DESTINATION ${CMAKE_INSTALL_BINDIR} COMPONENT runtime
    ARCHIVE DESTINATION ${CMAKE_INSTALL_LIBDIR} COMPONENT devel
    PUBLIC_HEADER DESTINATION ${CMAKE_INSTALL_INCLUDEDIR}/QXlsxQt${QT_VERSION_MAJOR} COMPONENT devel
)

install(EXPORT ${EXPORT_NAME}Targets
    DESTINATION ${CMAKE_INSTALL_LIBDIR}/cmake/${EXPORT_NAME}/
    FILE ${EXPORT_NAME}Targets.cmake
    NAMESPACE QXlsx::
    COMPONENT devel
)

configure_file(${CMAKE_CURRENT_SOURCE_DIR}/cmake/modules/qxlsx-config.cmake.in
    ${CMAKE_CURRENT_BINARY_DIR}/${EXPORT_NAME}Config.cmake
    @ONLY
)
configure_file(${CMAKE_CURRENT_SOURCE_DIR}/cmake/modules/qxlsx-config-version.cmake.in
    ${CMAKE_CURRENT_BINARY_DIR}/${EXPORT_NAME}ConfigVersion.cmake
    @ONLY
)
install(
    FILES
        ${CMAKE_CURRENT_BINARY_DIR}/${EXPORT_NAME}Config.cmake
        ${CMAKE_CURRENT_BINARY_DIR}/${EXPORT_NAME}ConfigVersion.cmake
    DESTINATION ${CMAKE_INSTALL_LIBDIR}/cmake/${EXPORT_NAME}/
)
include(CPackConfig)
