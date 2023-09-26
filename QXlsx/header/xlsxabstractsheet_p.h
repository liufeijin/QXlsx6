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
    std::optional<bool> showOutlineSymbols;
    std::optional<bool> autoPageBreaks;
    std::optional<bool> fitToPage;
    std::optional<bool> syncHorizontal;
    std::optional<bool> syncVertical;
    CellReference syncRef;
    std::optional<bool> transitionEvaluation;
    std::optional<bool> transitionEntry;
    std::optional<bool> published;
    QString codeName;
    std::optional<bool> filterMode;
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
    std::optional<SheetProtection> sheetProtection; //using optional allows adding default protection
    QList<SheetView> sheetViews;
    ExtensionList extLst;
    SheetProperties sheetProperties;
};

}

#endif // XLSXABSTRACTSHEET_P_H
