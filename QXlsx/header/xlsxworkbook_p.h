// xlsxworkbook_p.h

#ifndef XLSXWORKBOOK_P_H
#define XLSXWORKBOOK_P_H

#include <QtGlobal>
#include <QSharedPointer>
#include <QStringList>

#include "xlsxworkbook.h"
#include "xlsxabstractooxmlfile_p.h"
#include "xlsxtheme_p.h"
#include "xlsxsimpleooxmlfile_p.h"
#include "xlsxrelationships_p.h"

namespace QXlsx {

struct WorkbookView
{
//    default values to be used in new workbooks
//    xWindow = 240;
//    yWindow = 15;
//    windowWidth = 16095;
//    windowHeight = 9660;
//    activeTab = 0;
//    firstSheet = 0;

    std::optional<int> xWindow;
    std::optional<int> yWindow;
    std::optional<int> windowWidth;
    std::optional<int> windowHeight;
    std::optional<int> activeTab;
    std::optional<int> firstSheet;
    std::optional<bool> minimized; //default=false

    std::optional<AbstractSheet::Visibility> visibility;// default="visible"/>
    std::optional<bool> showHorizontalScroll;// default="true"/>
    std::optional<bool> showVerticalScroll;// default="true"/>
    std::optional<bool> showSheetTabs;// default="true"/>
    std::optional<int> tabRatio;// default="600"/>
    std::optional<bool> autoFilterDateGrouping;// default="true"/>

    ExtensionList extLst;
    void write(QXmlStreamWriter &writer) const;
    void read(QXmlStreamReader &reader);
};

struct WebPublishObject
{
    int id = 0;
    QString divId;
    QString sourceObject;
    QString destinationFile;
    QString title;
    std::optional<bool> autoRepublish;
};

class WorkbookPrivate : public AbstractOOXmlFilePrivate
{
    Q_DECLARE_PUBLIC(Workbook)
public:
    WorkbookPrivate(Workbook *q, Workbook::CreateFlag flag);

    QSharedPointer<SharedStrings> sharedStrings;
    QList<QSharedPointer<AbstractSheet> > sheets;
    QList<QSharedPointer<SimpleOOXmlFile> > externalLinks;
    QSharedPointer<Styles> styles;
    QSharedPointer<Theme> theme;
    QList<QWeakPointer<MediaFile> > mediaFiles;
    QList<QWeakPointer<Chart> > chartFiles;
    QList<DefinedName> definedNamesList;

    bool strings_to_numbers_enabled;
    bool strings_to_hyperlinks_enabled;
    bool html_to_richstring_enabled;

    QString defaultDateFormat;

    int table_count;

    // --- These members have no corresponding methods in Workbook
    // fileVersion
    QString appName{"xl"};
    QString lastEdited{"4"};
    QString lowestEdited{"4"};
    QString rupBuild{"4505"};
    QString appCodeName;

    // --- These members have no corresponding methods in Workbook
    // fileSharing
    std::optional<bool> readOnlyRecommended;
    QString userName;
    QString algorithmName;
    QString hashValue;
    QString saltValue;
    std::optional<int> spinCount;

    // workbookPr
    WorkbookProtection protection;

    // --- These members have no corresponding methods in Workbook
    // oleSize
    CellRange oleSize;

    // --- These members have no corresponding methods in Workbook
    // webPublishObjects
    QList<WebPublishObject> webPublishObjects;
    // webPublishing
    std::optional<bool> css;
    std::optional<bool> thicket;
    std::optional<bool> longFileNames;
    std::optional<bool> vml;
    std::optional<bool> allowPng;
    QString targetScreenSize;
    std::optional<int> dpi;
    QString characterSet;

    //Used to generate new sheet name and id
    int lastWorksheetIndex = 0;
    int lastChartsheetIndex = 0;
    int lastSheetId = 0;

    // workbookView
    mutable QList<WorkbookView> views;

    // --- Not all of these members have corresponding methods in Workbook
    // workBookPr
    std::optional<bool> date1904;
    std::optional<Workbook::ShowObjects> showObjects;
    std::optional<bool> showBorderUnselectedTables;
    std::optional<bool> filterPrivacy;
    std::optional<bool> promptedSolutions;
    std::optional<bool> showInkAnnotation;
    std::optional<bool> backupFile;
    std::optional<bool> saveExternalLinkValues;
    std::optional<Workbook::UpdateLinks> updateLinks;
    QString codeName;
    std::optional<bool> hidePivotFieldList;
    std::optional<bool> showPivotChartFilter;
    std::optional<bool> allowRefreshQuery;
    std::optional<bool> publishItems;
    std::optional<bool> checkCompatibility;
    std::optional<bool> autoCompressPictures;
    std::optional<bool> refreshAllConnections;
    std::optional<int> defaultThemeVersion;

    // --- Not all of these members have corresponding methods in Workbook
    // calcPr
    int calcId = 124519; //required, if missing we use 124519 (why?)
    std::optional<Workbook::CalculationMode> calcMode; //default="auto"
    std::optional<bool> fullCalcOnLoad; // default="false"/>
    std::optional<Workbook::ReferenceMode> refMode;// default="A1"/>
    std::optional<bool> iterate;// default="false"/>
    std::optional<int> iterateCount;// default="100"/>
    std::optional<double> iterateDelta;// default="0.001"/>
    std::optional<bool> fullPrecision;//default="true"/>
    std::optional<bool> calcCompleted;// default="true"/>
    std::optional<bool> calcOnSave;// default="true"/>
    std::optional<bool> concurrentCalc;// default="true"/>
    std::optional<int> concurrentManualCount;
    std::optional<bool> forceFullCalc;

    // --- These members have no corresponding methods in Workbook
    // functionGroups
    QStringList functionGroups;
    std::optional<int> builtInGroupCount;

    ExtensionList extLst;
};

}

#endif // XLSXWORKBOOK_P_H
