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

namespace QXlsx {

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

    ExtensionList extLst;
};

}

#endif // XLSXABSTRACTSHEET_P_H
