// xlsxworksheet.h

#ifndef XLSXWORKSHEET_H
#define XLSXWORKSHEET_H

#include <memory>
#include <QtGlobal>
#include <QObject>
#include <QStringList>
#include <QMap>
#include <QVariant>
#include <QPointF>
#include <QIODevice>
#include <QDateTime>
#include <QUrl>
#include <QImage>

#include "xlsxabstractsheet.h"
#include "xlsxcell.h"
#include "xlsxcellrange.h"
#include "xlsxcellreference.h"
#include "xlsxsheetview.h"

class WorksheetTest;

namespace QXlsx {

class DocumentPrivate;
class Workbook;
class Format;
class Drawing;
class DataValidation;
class ConditionalFormatting;
class CellRange;
class RichString;
class Relationships;
class Chart;



class WorksheetPrivate;
//TODO: Full documentation
/**
 * @brief The Worksheet class represents a worksheet in the workbook.
 *
 *
 * Each worksheet can have 1 to infinity 'sheet views', that display a specific portion of
 * the worksheet with specific view parameters.
 *
 * The following methods manage sheet views:
 *
 * - AbstractSheet::view(int index) returns a specific view.
 * - AbstractSheet::viewCount() returns the count of views in the sheet.
 * - AbstractSheet::addView() adds a view.
 * - AbstractSheet::removeView(int index) removes the view.
 *
 * The following methods manage the parameters of the _last added_ view:
 *
 * - #isWindowProtected(), #setWindowProtected(bool protect) manage view protection.
 * - #isFormulasVisible(), #setFormulasVisible(bool visible) manage formulas visibility.
 * - #isGridLinesVisible(), #setGridLinesVisible(bool visible) manage visibility of gridlines between sheet cells.
 * - #isRowColumnHeadersVisible(), #setRowColumnHeadersVisible(bool visible) manage headers visibility.
 * - #isZerosVisible(), #setZerosVisible(bool visible) manage the zero values appearance.
 * - #isRightToLeft(), #setRightToLeft(bool enable) manage the right-to-left appearance of the view.
 * - AbstractSheet::isSelected(), AbstractSheet::setSelected(bool select) manage the selection of the sheet tab.
 * - #isRulerVisible(), #setRulerVisible(bool visible) manage the ruler visibility.
 * - #isOutlineSymbolsVisible(), #setOutlineSymbolsVisible(bool visible) manage the outline symbols visibility.
 * - #isWhiteSpaceVisible(), #setWhiteSpaceVisible(bool visible) manage visibility of whitespaces in the sheet's cells.
 * - #selection(), #setSelection(), #selectedRanges(), #addSelection(), #removeSelection(), and #clearSelection() manage the selection on the worksheet.
 * - #activeCell(), #setActiveCell() manage the active cell on the worksheet.
 *
 * The above-mentioned methods return the default values if the corresponding parameters were not set.
 * See SheetView documentation on the default values.
 *
 */
class QXLSX_EXPORT Worksheet : public AbstractSheet
{
    Q_DECLARE_PRIVATE(Worksheet)

private:
    friend class DocumentPrivate;
    friend class Workbook;
    friend class ::WorksheetTest;
    Worksheet(const QString &sheetName, int sheetId, Workbook *book, CreateFlag flag);
    Worksheet *copy(const QString &distName, int distId) const override;

public:
    ~Worksheet();

public:
    bool write(const CellReference &row_column, const QVariant &value, const Format &format=Format());
    bool write(int row, int column, const QVariant &value, const Format &format=Format());

    QVariant read(const CellReference &row_column) const;
    QVariant read(int row, int column) const;

    bool writeString(const CellReference &row_column, const QString &value, const Format &format=Format());
    bool writeString(int row, int column, const QString &value, const Format &format=Format());
    bool writeString(const CellReference &row_column, const RichString &value, const Format &format=Format());
    bool writeString(int row, int column, const RichString &value, const Format &format=Format());

    bool writeInlineString(const CellReference &row_column, const QString &value, const Format &format=Format());
    bool writeInlineString(int row, int column, const QString &value, const Format &format=Format());

    bool writeNumeric(const CellReference &row_column, double value, const Format &format=Format());
    bool writeNumeric(int row, int column, double value, const Format &format=Format());

    bool writeFormula(const CellReference &row_column, const CellFormula &formula, const Format &format=Format(), double result=0);
    bool writeFormula(int row, int column, const CellFormula &formula, const Format &format=Format(), double result=0);

    bool writeBlank(const CellReference &row_column, const Format &format=Format());
    bool writeBlank(int row, int column, const Format &format=Format());

    bool writeBool(const CellReference &row_column, bool value, const Format &format=Format());
    bool writeBool(int row, int column, bool value, const Format &format=Format());

    bool writeDateTime(const CellReference &row_column, const QDateTime& dt, const Format &format=Format());
    bool writeDateTime(int row, int column, const QDateTime& dt, const Format &format=Format());

    // dev67
    bool writeDate(const CellReference &row_column, const QDate& dt, const Format &format=Format());
    bool writeDate(int row, int column, const QDate& dt, const Format &format=Format());

    bool writeTime(const CellReference &row_column, const QTime& t, const Format &format=Format());
    bool writeTime(int row, int column, const QTime& t, const Format &format=Format());

    bool writeHyperlink(const CellReference &row_column, const QUrl &url, const Format &format=Format(), const QString &display=QString(), const QString &tip=QString());
    bool writeHyperlink(int row, int column, const QUrl &url, const Format &format=Format(), const QString &display=QString(), const QString &tip=QString());

    bool addDataValidation(const DataValidation &validation);
    bool addConditionalFormatting(const ConditionalFormatting &cf);
    /**
     * @brief returns cell by its reference.
     * @param row_column reference to the cell.
     * @return valid pointer to the cell if the cell was found, nullptr otherwise.
     */
    Cell *cell(const CellReference &row_column) const;
    /**
     * @overload
     * @brief returns cell by its row and column number.
     * @param row 1-based cell row number.
     * @param column 1-based cell column number.
     * @return valid pointer to the cell if the cell was found, nullptr otherwise.
     */
    Cell *cell(int row, int column) const;

    /**
     * @brief Inserts an @a image at the position @a row, @a column.
     * @param row the 1-based row index of the image top left corner.
     * @param column the 1-based column index of the image top left corner.
     * @param image Image to be inserted.
     * @return image index (zero-based) on success, -1 otherwise.
     */
    int insertImage(int row, int column, const QImage &image);
    /**
     * @brief returns an image.
     * @param index zero-based image index (0 to #imageCount()-1).
     * @return non-null QImage if image was found and read.
     */
    QImage image(int index) const;
    /**
     * @overload
     * @brief returns image.
     * @param row the 1-based row index of the image top left corner.
     * @param column the 1-based column index of the image top left corner.
     * @return non-null QImage if image was found and read.
     */
    QImage image(int row, int column) const;
    /**
     * @brief returns the count of images on this worksheet.
     */
    int imageCount() const;
    /**
     * @brief removes image from the worksheet
     * @param row the 1-based row index of the image top left corner.
     * @param column the 1-based column index of the image top left corner.
     * @return true if image was found and removed, false otherwise.
     */
    bool removeImage(int row, int column);
    /**
     * @brief removes image from the worksheet
     * @param index zero-based index of the image (0 to #imageCount()-1).
     * @return true if image was found and removed, false otherwise.
     */
    bool removeImage(int index);

    /**
     * @brief creates a new chart and places it inside the current worksheet.
     * @param row the 1-based row index of the chart top left corner.
     * @param column the 1-based column index of the chart top left corner.
     * @param size the chart size in pixels at 96 dpi.
     * @return pointer to the new chart or nullptr if no chart was created.
     */
    Chart *insertChart(int row, int column, const QSize &size);
    /**
     * @overload
     * @brief creates a new chart and places it inside the current worksheet.
     * @param row the 1-based row index of the chart top left corner.
     * @param column the 1-based column index of the chart top left corner.
     * @param width width of a chart specified as a Coordinate object. You can use it to set
     * width in pixels, points, millimeters, EMU etc. See Coordinate for help.
     * @param height height of a chart specified as a Coordinate object. You can use it to set
     * height in pixels, points, millimeters, EMU etc. See Coordinate for help.
     * @return pointer to the new chart or nullptr if no chart was created.
     */
    Chart *insertChart(int row, int column, Coordinate width, Coordinate height);
    /**
     * @brief returns chart
     * @param index zero-based index of the chart (0 to #chartCount()-1)
     * @return valid pointer to the chart if the chart was found, nullptr otherwise.
     */
    Chart *chart(int index) const;
    /**
     * @overload
     * @brief returns chart.
     * @param row the 1-based row index of the chart top left corner.
     * @param column the 1-based column index of the chart top left corner.
     * @return valid pointer to the chart if the chart was found, nullptr otherwise.
     */
    Chart *chart(int row, int column) const;
    /**
     * @brief removes chart from the worksheet.
     * @param row the 1-based row index of the chart top left corner.
     * @param column the 1-based column index of the chart top left corner.
     * @return true if the chart was found and successfully removed, false otherwise.
     */
    bool removeChart(int row, int column);
    /**
     * @overload
     * @brief removes chart from the worksheet.
     * @param index zero-based index of the chart.
     * @return true if the chart was found and successfully removed, false otherwise.
     */
    bool removeChart(int index);
    /**
     * @brief returns count of charts on the worksheet.
     */
    int chartCount() const;
    /**
    @brief Merges a @a range of cells.

    The first cell will retain its data, other cells will be cleared.
    All cells will get the same @a format if a valid @a format is given.
    Returns true on success.
    */
    bool mergeCells(const CellRange &range, const Format &format=Format());
    bool unmergeCells(const CellRange &range);
    QList<CellRange> mergedCells() const;

    bool setColumnWidth(const CellRange& range, double width);
    bool setColumnFormat(const CellRange& range, const Format &format);
    bool setColumnHidden(const CellRange& range, bool hidden);
    bool setColumnWidth(int colFirst, int colLast, double width);
    bool setColumnFormat(int colFirst, int colLast, const Format &format);
    bool setColumnHidden(int colFirst, int colLast, bool hidden);

    double columnWidth(int column) const;
    Format columnFormat(int column) const;
    bool isColumnHidden(int column) const;

    /**
     * @brief sets rows heights for rows in the range from rowFirst to rowLast.
     * @param rowFirst the first row index (1-based).
     * @param rowLast the last row index (1-based).
     * @param height row height in points.
     * @return true on success.
     */
    bool setRowHeight(int rowFirst,int rowLast, double height);
    bool setRowFormat(int rowFirst,int rowLast, const Format &format);
    bool setRowHidden(int rowFirst,int rowLast, bool hidden);

    double rowHeight(int row) const;
    Format rowFormat(int row) const;
    bool isRowHidden(int row) const;

    bool groupRows(int rowFirst, int rowLast, bool collapsed = true);
    /**
     * @brief groups columns from @a colFirst to @a colLast and collapse them as specified.
     * @param colFirst 1-based index of the first column.
     * @param colLast 1-based index of the last column.
     * @param collapsed if true, columns will be collapsed.
     * @return true if the maximum level of grouping is not exceeded and the grouping is successful,
     * false otherwise.
     *
     * The "group/ungroup" button will appear over the column with index colLast+1.
     */
    bool groupColumns(int colFirst, int colLast, bool collapsed = true);
    /**
     * @overload
     * @brief groups columns from @a range and collapse them as specified.
     * @param range range of columns to group.
     * @param collapsed
     * @return true if the maximum level of grouping is not exceeded and the grouping is successful,
     * false otherwise.
     *
     * The "group/ungroup" button will appear over the column with index colLast+1.
     */
    bool groupColumns(const CellRange &range, bool collapsed = true);
    CellRange dimension() const;
    /**
     * @brief returns whether the panes in the window are locked due to workbook protection.
     *
     * The default value is false.
     * @note The worksheet can have more than one view. This method returns the
     * property of the last one. To get the specific view use #view() method.
     */
    bool isWindowProtected() const;
    /**
     * @brief sets whether the panes in the window are locked due to workbook protection.
     *
     * @param protect Sets protected to the view. The default value is false.
     * @note The worksheet can have more than one view. This method sets the
     * property of the last one. If no sheet views were added, this method
     * adds the default one. To get the specific view use #view() method.
     */
    void setWindowProtected(bool protect);
    //TODO: doc
    bool isFormulasVisible() const;
    void setFormulasVisible(bool visible);//TODO: doc
    bool isGridLinesVisible() const;//TODO: doc
    void setGridLinesVisible(bool visible);//TODO: doc
    bool isRowColumnHeadersVisible() const;//TODO: doc
    void setRowColumnHeadersVisible(bool visible);//TODO: doc
    /**
     * @brief returns whether the window should show 0 (zero) in cells
     * containing zero value.
     * @return true if zeroes are displayed as is, false if cells with zero value
     * appear blank.
     *
     * The default value is true.
     */
    bool isZerosVisible() const;
    /**
     * @brief sets whether the window should show 0 (zero) in cells
     * containing zero value.
     * @param visible If true, cells containing zero value shall display 0. If false,
     * cells containing zero value shall appear blank.
     *
     * If not set, the default value is true.
     */
    void setZerosVisible(bool visible);
    /**
     * @brief returns whether the sheet is in 'right to left' display mode.
     *
     * When in this mode, Column A is on the far right, Column B is one column
     * left of Column A, and so on. Also, information in cells is displayed in
     * the Right to Left format.
     *
     * @return true if the sheet is in 'right to left' display mode.
     *
     * The default value is false.
     */
    bool isRightToLeft() const;
    /**
     * @brief sets whether the sheet is in 'right to left' display mode.
     *
     * @param enable If true, Column A is on the far right, Column B is one column
     * left of Column A, and so on. Also, information in cells is displayed in
     * the Right to Left format.
     *
     * If not set, the default value is false.
     */
    void setRightToLeft(bool enable);

    bool isRulerVisible() const;//TODO: doc
    void setRulerVisible(bool visible);//TODO: doc
    bool isOutlineSymbolsVisible() const;//TODO: doc
    void setOutlineSymbolsVisible(bool visible);//TODO: doc
    bool isWhiteSpaceVisible() const;//TODO: doc
    void setWhiteSpaceVisible(bool visible);//TODO: doc
    bool isDefaultGridColorUsed() const;
    void setDefaultGridColorUsed(bool value);
    SheetView::Type viewType() const;
    void setViewType(SheetView::Type type);
    /**
     * @brief returns the location of the last added view's top left cell.
     * @return Valid location if it was set, invalid one otherwise.
     */
    CellReference viewTopLeftCell() const;
    void setViewTopLeftCell(const CellReference &ref);
    int viewColorIndex() const;
    void setViewColorIndex(int index);


    /**
     * @brief returns the last defined sheet view's active cell.
     * @return copy of CellReference object.
     */
    CellReference activeCell() const;
    /**
     * @brief sets active cell to the last added sheet view.
     * @param activeCell
     */
    void setActiveCell(const CellReference &activeCell);
    /**
     * @brief returns a list of cell ranges selected in the last added sheet view.
     * @return
     */
    QList<CellRange> selectedRanges() const;
    /**
     * @brief adds @a range to the list of cell ranges selected in the last added sheet view.
     * @param range valid cell range.
     * @return true if @a range was successfully added, false if @a range is invalid or already
     * present in the selection.
     */
    bool addSelection(const CellRange &range);
    /**
     * @brief removes range from the list of cell ranges selected in the last added sheet view.
     * @param range cell range to remove.
     * @return true if @a range was found and removed, false otherwise.
     * @note This method does not check @a range for intersections with selection, it simply
     * searches the selection for the whole range and if found, removes it.
     */
    bool removeSelection(const CellRange &range);
    /**
     * @brief removes all selection from the last added sheet view.
     */
    void clearSelection();
    /**
     * @brief returns selection parameters of the last added sheet view.
     * @return copy of Selection object.
     */
    Selection selection() const;
    /**
     * @brief returns selection parameters of the last added sheet view.
     * @return reference to the Selection object.
     */
    Selection &selection();
    /**
     * @brief sets selection parameters of the last added sheet view.
     * @param selection the Selection object.
     */
    void setSelection(const Selection &selection);


    //TODO: test methods to set, add, remove and clear selections.

    /**
     * @brief sets the worksheet's print scale in percents.
     * @param scale value from 10 to 400. The value of 100 equals 100% scaling.
     */
    void setPrintScale(int scale);
    /**
     * @brief returns the worksheet's print scale in percents.
     * @return value from 10 to 400. If no print scale was specified, this method
     * returns the default value of 100.
     */
    int printScale() const;
    /**
     * @brief returns the order of printing the worksheet pages.
     * @return
     */
    PageSetup::PageOrder pageOrder() const;
    /**
     * @brief sets the order of printing the worksheet pages.
     * @param pageOrder
     */
    void setPageOrder(PageSetup::PageOrder pageOrder);

    /**
     * @brief autosizes columns widths for all rows in the worksheet.
     * @param firstColumn 1-based index of the first column to autosize.
     * @param lastColumn 1-based index of the last column to autosize.
     * @return true on success.
     */
    bool autosizeColumnWidths(int firstColumn, int lastColumn);
    /**
     * @overload
     * @brief autosizes columns widths for columns specified by range.
     * @param range valid CellRange that specifies the columns range to autosize and
     * the rows range to estimate the maximum column width.
     * @return true on success.
     */
    bool autosizeColumnWidths(const CellRange &range);

private:
    QMap<int, int> getMaximumColumnWidths(int firstRow = 1, int lastRow = INT_MAX);
    void saveToXmlFile(QIODevice *device) const override;
    bool loadFromXmlFile(QIODevice *device) override;
};

}
#endif // XLSXWORKSHEET_H
