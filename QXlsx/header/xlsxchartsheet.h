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
 * Each chartsheet has a pointer to #chart() that can be used to edit the chart
 * series, axes, title etc.
 *
 * Chartsheets inherit from AbstractSheet class all methods that allow to modify
 * sheet properties (code name, tab color etc.), sheet views, page and print
 * options, header and footer etc.
 *
 *
 *
 *
 *
 */
class QXLSX_EXPORT Chartsheet : public AbstractSheet
{
    Q_DECLARE_PRIVATE(Chartsheet)

public:
    ~Chartsheet();
    Chart *chart();

    /**
     * @brief returns whether the chartsheet is zoomed to fit the sheet view.
     * @return `true` if the chartsheet is zoomed.
     *
     * The default value is `false` according to ECMA-376, but to mimic the Excel behavior if
     * the parameter was not set directly, then on writing the document zoomToFit is written as `true`.
     */
    std::optional<bool> zoomToFit() const;
    /**
     * @brief sets whether the chartsheet is zoomed to fit the sheet view.
     * @param zoom
     *
     * The default value is `false` according to ECMA-376, but to mimic the Excel behavior if
     * the parameter is not set directly, then on writing the document zoomToFit is written as `true`.
     */
    void setZoomToFit(bool zoom);
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
