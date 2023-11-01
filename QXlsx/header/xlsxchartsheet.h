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
 * # Sheet Properties
 *
 * The following parameters define the behavior and look-and-feel of the chartsheet:
 *
 * - AbstractSheet::codeName(), AbstractSheet::setCodeName() manage the unique stable name of the sheet.
 * - AbstractSheet::isPublished(), AbstractSheet::setPublished() manage the publishing of the sheet.
 * - AbstractSheet::tabColor(), AbstractSheet::setTabColor() manage the sheet's tab color.
 *
 * # Sheet Views
 *
 * Each chartsheet can have 1 to infinity 'sheet views', that display a specific portion of
 * the chartsheet with specific view parameters.
 *
 * The following methods manage sheet views:
 *
 * - AbstractSheet::view(int index) returns a specific view.
 * - AbstractSheet::viewsCount() returns the count of views in the sheet.
 * - AbstractSheet::addView() adds a view.
 * - AbstractSheet::removeView(int index) removes the view.
 *
 * The following methods manage the parameters of the _last added_ view:
 *
 * - AbstractSheet::isSelected(), AbstractSheet::setSelected(bool select) manage the selection of the sheet tab.
 * - AbstractSheet::viewZoomScale(), AbstractSheet::setViewZoomScale() manage the scale of the current view.
 * - #zoomToFit(), #setZoomToFit() manage the automatic zooming of the chartsheet.
 *
 * The above-mentioned methods return the default values if the corresponding parameters were not set.
 * See SheetView documentation on the default values.
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
     * @return true if the chartsheet is zoomed.
     *
     * The default value is false according to ECMA-376, but to mimic the Excel behavior if
     * the parameter was not set directly, then on writing the document zoomToFit is written as true.
     */
    std::optional<bool> zoomToFit() const;
    /**
     * @brief sets whether the chartsheet is zoomed to fit the sheet view.
     * @param zoom
     *
     * The default value is false according to ECMA-376, but to mimic the Excel behavior if
     * the parameter is not set directly, then on writing the document zoomToFit is written as true.
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
