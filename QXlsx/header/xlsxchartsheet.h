// xlsxchartsheet.h

#ifndef XLSXCHARTSHEET_H
#define XLSXCHARTSHEET_H

#include <QtGlobal>
#include <QStringList>

#include "xlsxabstractsheet.h"

namespace QXlsx {

class Workbook;
class DocumentPrivate;
class ChartsheetPrivate;
class Chart;

/**
 * @brief The Chartsheet class represents a chartsheet in a workbook.
 *
 * Each chartsheet has a pointer to #chart() that can be used to edit the chart series,
 * axes, title etc.
 *
 */
class QXLSX_EXPORT Chartsheet : public AbstractSheet
{
    Q_DECLARE_PRIVATE(Chartsheet)

public:
    ~Chartsheet();
    Chart *chart();

private:
    friend class DocumentPrivate;
    friend class Workbook;

    Chartsheet(const QString &sheetName, int sheetId, Workbook *book, CreateFlag flag);
    Chartsheet *copy(const QString &distName, int distId) const override;

    void saveToXmlFile(QIODevice *device) const override;
    bool loadFromXmlFile(QIODevice *device) override;
};

}
#endif // XLSXCHARTSHEET_H
