// documentproperty.cpp

#include <QtGlobal>
#include <QtCore>
#include <QDebug>

#include "xlsxdocument.h"

int documentproperty()
{
    QXlsx::Document xlsx;
    xlsx.write("A1", "View the properties through:");
    xlsx.write("A2", "Office Button -> Prepare -> Properties option in Excel");

    xlsx.setMetadata(QXlsx::Document::Metadata::Title, "This is an example spreadsheet");
    xlsx.setMetadata(QXlsx::Document::Metadata::Subject, "With document properties");
    xlsx.setMetadata(QXlsx::Document::Metadata::Creator, "Debao Zhang");
    xlsx.setMetadata(QXlsx::Document::Metadata::Company, "HMICN");
    xlsx.setMetadata(QXlsx::Document::Metadata::Category, "Example spreadsheets");
    xlsx.setMetadata(QXlsx::Document::Metadata::Keywords, "Sample, Example, Properties");
    xlsx.setMetadata(QXlsx::Document::Metadata::Description, "Created with QXlsx");

    xlsx.saveAs("documentproperty.xlsx");
    return 0;
}
