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

class WorkbookPrivate : public AbstractOOXmlFilePrivate
{
    Q_DECLARE_PUBLIC(Workbook)
public:
    WorkbookPrivate(Workbook *q, Workbook::CreateFlag flag);

    QSharedPointer<SharedStrings> sharedStrings;
    QList<QSharedPointer<AbstractSheet> > sheets;
    QList<QSharedPointer<SimpleOOXmlFile> > externalLinks;
    QStringList sheetNames;
    QSharedPointer<Styles> styles;
    QSharedPointer<Theme> theme;
    QList<QWeakPointer<MediaFile> > mediaFiles;
    QList<QWeakPointer<Chart> > chartFiles;
    QList<DefinedName> definedNamesList;

    bool strings_to_numbers_enabled;
    bool strings_to_hyperlinks_enabled;
    bool html_to_richstring_enabled;

    QString defaultDateFormat;

    int x_window;
    int y_window;
    int window_width;
    int window_height;

    int activesheetIndex;
    int firstsheet;
    int table_count;

    //Used to generate new sheet name and id
    int lastWorksheetIndex = 0;
    int lastChartsheetIndex = 0;
    int lastSheetId = 0;

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
};

}

#endif // XLSXWORKBOOK_P_H
