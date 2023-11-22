// xlsxabstractsheet_p/h

#ifndef XLSXABSTRACTSHEET_P_H
#define XLSXABSTRACTSHEET_P_H

#include <QString>
#include <QImage>

#include <memory>

#include "xlsxglobal.h"

#include "xlsxabstractsheet.h"
#include "xlsxabstractooxmlfile_p.h"
#include "xlsxdrawing_p.h"
#include "xlsxmediafile_p.h"
#include "xlsxsheetview.h"

namespace QXlsx {

struct SheetProperties
{
    Color tabColor;
    std::optional<bool> applyStyles;
    std::optional<bool> summaryBelow;
    std::optional<bool> summaryRight;
    std::optional<bool> showOutlineSymbols; // according to ECMA-376 this attribute
    //should be overridden with the view's showOutlineSymbols in case of conflict.
    //So I removed the methods of accessing this attribute.
    std::optional<bool> autoPageBreaks;
    std::optional<bool> fitToPage;
    std::optional<bool> syncHorizontal;
    std::optional<bool> syncVertical;
    CellReference syncRef;
    std::optional<bool> transitionEvaluation; // this attribute has no doc in ECMA-376,
    //so I ignored it as an unknown attribute.
    std::optional<bool> transitionEntry; // this attribute has no doc in ECMA-376,
    //so I ignored it as an unknown attribute.
    std::optional<bool> published;
    QString codeName;
    std::optional<bool> filterMode; //this attribute is set internally if autofiltering is applied.
    std::optional<bool> enableFormatConditionsCalculation;
    bool isValid() const;
    void read(QXmlStreamReader &reader);
    void write(QXmlStreamWriter &writer, const QLatin1String &name) const;
};

class AbstractSheetPrivate : public AbstractOOXmlFilePrivate
{
    Q_DECLARE_PUBLIC(AbstractSheet)
public:
    AbstractSheetPrivate(AbstractSheet *p, AbstractSheet::CreateFlag flag);
    ~AbstractSheetPrivate();

    Workbook *workbook;
    std::shared_ptr<Drawing> drawing;

    QString name;
    int id;
    AbstractSheet::Visibility sheetState;
    AbstractSheet::Type type;
    HeaderFooter headerFooter;
    PageMargins pageMargins;
    PageSetup pageSetup;
    QSharedPointer<MediaFile> pictureFile;
    SheetProtection sheetProtection;
    QList<SheetView> sheetViews;
    ExtensionList extLst;
    SheetProperties sheetProperties;
};

}

#endif // XLSXABSTRACTSHEET_P_H
