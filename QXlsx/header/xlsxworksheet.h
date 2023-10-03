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
#include "xlsxautofilter.h"
#include "xlsxdatavalidation.h"

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
 * # Sheet Properties
 *
 * The following parameters define the behavior and look-and-feel of the worksheet:
 *
 * - AbstractSheet::codeName(), AbstractSheet::setCodeName() manage the unique stable name of the sheet.
 * - #isFormatConditionsCalculationEnabled(), #setFormatConditionsCalculationEnabled()
 *   manage the recalculation of the conditional formatting.
 * - AbstractSheet::isPublished(), AbstractSheet::setPublished() manage the publishing of the sheet.
 * - #syncHorizontal(), #setSyncHorizontal(), #syncVertical(), #setSyncVertical(),
 *   #syncRef(), #setSyncRef() manage the anchor parameters of the sheet.
 * - AbstractSheet::tabColor(), AbstractSheet::setTabColor() manage the sheet's tab color.
 * - #showAutoPageBreaks(), #setShowAutoPageBreaks() manage the auto page breaks visibility.
 * - #thickBottom(), #setThickBottom(), #thickTop(), #setThickTop() manage default thickness of rows borders.
 * - #rowsHiddenByDefault(), #setRowsHiddenByDefault() manage the default visibility of rows.
 *
 * # Sheet Views
 *
 * Each worksheet can have 1 to infinity 'sheet views', that display a specific portion of
 * the worksheet with specific view parameters.
 *
 * There's always at least one default sheet view.
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
 * - #isDefaultGridColorUsed(), #setDefaultGridColorUsed(bool value) manage using the default/custom grid lines color.
 * - #viewType(), #setViewType(SheetView::Type type) manage the type of the sheet view.
 * - #viewTopLeftCell(), #setViewTopLeftCell(const CellReference &ref) manage the reference to the view's top left cell.
 * - #viewColorIndex(), #setViewColorIndex(int index) manage the sheet tab's color in the current view.
 * - #selection(), #setSelection(), #selectedRanges(), #addSelection(), #removeSelection(), and #clearSelection() manage the selection on the worksheet.
 * - #activeCell(), #setActiveCell() manage the active cell on the worksheet.
 *
 * The above-mentioned methods return the default values if the corresponding parameters were not set.
 * See SheetView and Selection documentation on the default values.
 *
 * # Sheet Printing Parameters
 *
 *
 * # Rows and columns
 *
 * Row heights are measured in points. The row height can be set separately for
 * each row with #setRowHeight() or for all rows in the sheet with
 * #setDefaultRowHeight(), and tested with #rowHeight().
 * The default row height is 14.4 pt (24 px at 96 dpi).
 *
 * Column widths are measured in widths of the widest digit of the normal style's font.
 * The column width can be set separately for each column with #setColumnWidth() or
 * for all columns in the sheet with #setDefaultColumnWidth(), and tested with
 * #columnWidth(). The default column width is 8 5/7 widths of the widest digit
 * of the normal style's font (Calibri 11 pt).
 *
 * Rows and columns can be hidden with #setRowHidden(), #setColumnHidden(). Also if only
 * some rows need to be visible, then #setRowsHiddenByDefault() can come in handy.
 * Rows and columns visibility can be tested with #isRowHidden() and #isColumnHidden().
 *
 * # Data manipulation
 *
 *
 * # Data validation
 *
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

    bool setFormat(const CellRange &range, const Format &format);
    bool setFormat(int row, int column, const Format &format);
    //TODO: bool clearFormat()

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


    /// Data Validation methods
    /**
     * @brief adds data validation in the sheet.
     * @param validation A valid DataValidation object.
     * @return true if @a validation is valid and contains valid range(s), false otherwise.
     */
    bool addDataValidation(const DataValidation &validation);
    /**
     * @overload
     * @brief adds data validation in the sheet.
     *
     * This is a low-level method that allows to create a validation from scratch.
     *
     * @param range cells range to apply data validation to.
     * @param type type of data validation
     * @param formula1 the first validation criterion
     * @param predicate operation to combine two validation criteria
     * @param formula2 the second validation criterion.
     * @param strict If true, then error message will be shown if the input cell value
     * is not allowable.
     * @return true if validation was successfully added, false otherwise.
     *
     * You can use overloaded methods to add a specific type of validation.
     */
    bool addDataValidation(const CellRange &range, DataValidation::Type type, const QString &formula1,
                           std::optional<DataValidation::Predicate> predicate = std::nullopt,
                           const QString &formula2 = QString(), bool strict = true);
    /**
     * @overload
     * @brief adds the data validation that allows only a list of specified values.
     * @param range cells range to apply data validation to.
     * @param allowableValues cells range that contains all allowable values.
     * @param strict If true, then error message will be shown if the input cell value
     * is not allowable.
     * @return true if validation was successfully added, false otherwise.
     * @note If you need to customize validation params, f.e. prompt and error messages,
     * use DataValidation constructor.
     */
    bool addDataValidation(const CellRange &range, const CellRange &allowableValues, bool strict = true);
    /**
     * @overload
     * @brief adds the data validation that validates time values.
     * @param range cells range to apply data validation to.
     * @param time1 first time value.
     * @param predicate operation to combine two validation criteria.
     * @param time2 second time value.
     * @param strict If true, then error message will be shown if the input cell value
     * is not allowable.
     * @return true if validation was successfully added, false otherwise.
     * @note If you need to customize validation params, f.e. prompt and error messages,
     * use DataValidation constructor.
     */
    bool addDataValidation(const CellRange &range, const QTime &time1,
                           std::optional<DataValidation::Predicate> predicate = std::nullopt,
                           const QTime &time2 = QTime(), bool strict = true);
    /**
     * @overload
     * @brief adds the data validation that validates date values.
     * @param range cells range to apply data validation to.
     * @param date1 first date value.
     * @param predicate operation to combine two validation criteria.
     * @param date2 second date value.
     * @param strict If true, then error message will be shown if the input cell value
     * is not allowable.
     * @return true if validation was successfully added, false otherwise.
     * @note If you need to customize validation params, f.e. prompt and error messages,
     * use DataValidation constructor.
     */
    bool addDataValidation(const CellRange &range, const QDate &date1,
                           std::optional<DataValidation::Predicate> predicate = std::nullopt,
                           const QDate &date2 = QDate(), bool strict = true);
    /**
     * @overload
     * @brief adds data validation that validates text length.
     * @param range cells range to apply data validation to.
     * @param len1 First text length
     * @param predicate operation to combine two validation criteria.
     * @param len2 Second text length.
     * @param strict If true, then error message will be shown if the input cell value
     * is not allowable.
     * @return true if validation was successfully added, false otherwise.
     * @note If you need to customize validation params, f.e. prompt and error messages,
     * use DataValidation constructor.
     */
    bool addDataValidation(const CellRange &range, int len1,
                           std::optional<DataValidation::Predicate> predicate = std::nullopt,
                           std::optional<int> len2 = std::nullopt, bool strict = true);
    /**
     * @brief returns whether all input prompts from the worksheet are disabled.
     * @return if true, then all data validation prompts will not be shown on the worksheet.
     *
     * The default value is false.
     */
    bool dataValidationPromptsDisabled() const;
    /**
     * @brief sets whether all input prompts from the worksheet are disabled.
     * @param disabled if true, then all data validation prompts will not be shown on the worksheet.
     *
     * If not set, false is assumed.
     */
    void setDataValidationPromptsDisabled(bool disabled);




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
     * @brief sets new image content to the image specified by @a index.
     * @param index zero-based index of the image (0 to #imagecount()-1).
     * @param fileName file name where to load new content from.
     * @param keepSize if true, then new image will be resized to the old image size.
     * If false, then new image will have its size.
     * @return true if the image was found and file was loaded, false otherwise.
     */
    bool changeImage(int index, const QString &fileName, bool keepSize = true);

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
    /**
     * @brief Un-merges a @a range of cells.
     * @param range the range to unmerge.
     * @return true if @a range was successfully unmerged.
     * @note This method does not check the merged cells range for intersections with
     * other merged cells.
     */
    bool unmergeCells(const CellRange &range);
    /**
     * @brief returns the list of merged ranges.
     * @return non-empty list of merged ranges if any of cells were merged.
     */
    QList<CellRange> mergedCells() const;

    bool setColumnWidth(const CellRange& range, double width);
    bool setColumnFormat(const CellRange& range, const Format &format);
    bool setColumnHidden(const CellRange& range, bool hidden);
    bool setColumnWidth(int colFirst, int colLast, double width);
    bool setColumnFormat(int colFirst, int colLast, const Format &format);
    bool setColumnHidden(int colFirst, int colLast, bool hidden);
    /**
     * @brief returns the width of @column in characters.
     * @param column column index (1-based).
     * @return the width of @column in characters of a standard font.
     */
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
    /**
     * @brief sets @a hidden to rows from @a rowFirst to @a rowLast.
     * @param rowFirst The first row index (1-based) of the range to set @a hidden.
     * @param rowLast The last row index (1-based) of the range to set @a hidden.
     * @param hidden visibility of rows.
     * @return true if @a rowFirst and @a rowLast are valid and the visibility is successfully changed.
     */
    bool setRowHidden(int rowFirst, int rowLast, bool hidden);
    /*!
     Returns height of \a row in points.
    */
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
    /**
     * @brief returns the dimension (extent) of the worksheet.
     * @return dimension (extent) of the worksheet.
     */
    CellRange dimension() const;

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




    // Sheet Parameters

    /**
     * @brief returns whether the conditional formatting calculations shall be
     * evaluated. If set to false, then the min/max values of color scales or
     * databars or threshold values in Top N rules shall not be updated.
     *
     * This is useful when conditional formats are being set programmatically at
     * runtime, recalculation of the conditional formatting does not need to be
     * done until the program execution has finished setting all the conditional
     * formatting properties.
     *
     * If not set, the default value is true.
     *
     * @return true if the recalculation of the conditional formatting is on.
     */
    bool isFormatConditionsCalculationEnabled() const;
    /**
     * @brief sets whether the conditional formatting calculations shall be
     * evaluated.
     * @param enabled If false, then the min/max values of color scales or
     * databars or threshold values in Top N rules shall not be updated.
     *
     * This is useful when conditional formats are being set programmatically at
     * runtime, recalculation of the conditional formatting does not need to be
     * done until the program execution has finished setting all the conditional
     * formatting properties.
     *
     * The default value is true.
     */
    void setFormatConditionsCalculationEnabled(bool enabled);
    /**
     * @brief returns whether the worksheet is horizontally synced to the #topLeftAnchor().
     * @return If true, and scroll location is missing from the window properties,
     * the window view shall be scrolled to the #topLeftAnchor() column.
     *
     * If not set, the default value is false.
     */
    bool isSyncedHorizontal() const;
    /**
     * @brief sets whether the worksheet is horizontally synced to the #topLeftAnchor().
     * @param sync If true, and scroll location is missing from the window properties,
     * the window view shall be scrolled to the #topLeftAnchor() column.
     *
     * The default value is false.
     */
    void setSyncedHorizontal(bool sync);
    /**
     * @brief returns whether the worksheet is vertically synced to the #topLeftAnchor().
     * @return If true, and scroll location is missing from the window properties,
     * the window view shall be scrolled to the #topLeftAnchor() row.
     *
     * If not set, the default value is false.
     */
    bool isSyncedVertical() const;
    /**
     * @brief sets whether the worksheet is vertically synced to the #topLeftAnchor().
     * @param sync If true, and scroll location is missing from the window properties,
     * the window view shall be scrolled to the #topLeftAnchor() row.
     *
     * The default value is false.
     */
    void setSyncedVertical(bool sync);
    /**
     * @brief returns the anchor cell the window view shall be scrolled to according
     * to the #isSyncedHorizontal() and #isSyncedVertical() parameters.
     * @return valid CellReference if the anchor cell was set.
     */
    CellReference topLeftAnchor() const;
    /**
     * @brief sets the anchor cell the window view shall be scrolled to according
     * to the #isSyncedHorizontal() and #isSyncedVertical() parameters.
     * @param ref If not valid, clears the anchor cell.
     */
    void setTopLeftAnchor(const CellReference &ref);
    /**
     * @brief returns  whether the sheet displays Automatic Page Breaks.
     * @return automatic page breaks visibility.
     *
     * If not set, the defalult value is true.
     */
    bool showAutoPageBreaks() const;
    /**
     * @brief sets whether the sheet displays Automatic Page Breaks.
     * @param show automatic page breaks visibility.
     *
     * The default value is true.
     */
    void setShowAutoPageBreaks(bool show);
    /**
     * @brief returns whether the Fit to Page print option is enabled.
     * @return the Fit to Page print option.
     *
     * If not set, the default value is false.
     */
    bool fitToPage() const;
    /**
     * @brief sets the Fit to Page print option.
     * @param value If true, then the sheet's contents will be scaled to fit the page.
     *
     * The default value is false.
     */
    void setFitToPage(bool value);

    // Sheet Format Parameters

    /**
     * @brief returns whether rows have thick bottom borders by default.
     * @return true if rows have thick bottom borders by default.
     *
     * The default value is false.
     */
    bool thickBottomBorder() const;
    /**
     * @brief sets whether rows have thick bottom borders by default.
     * @param thick if true, then rows will have thick bottom borders by default.
     *
     * The default value is false.
     */
    void setThickBottomBorder(bool thick);
    /**
     * @brief returns whether rows have thick top borders by default.
     * @return true if rows have thick top borders by default.
     *
     * The default value is false.
     */
    bool thickTopBorder() const;
    /**
     * @brief sets whether rows have thick top borders by default.
     * @param thick if true, then rows will have thick top borders by default.
     *
     * The default value is false.
     */
    void setThickTopBorder(bool thick);

    /**
     * @brief returns whether rows are hidden by default.
     * @return true if all rows are hidden by default.
     *
     * The default value is false.
     *
     * This parameter is useful if it is much shorter to only write out the rows
     * that are not hidden with #setRowHidden().
     */
    bool rowsHiddenByDefault() const;
    /**
     * @brief sets whether rows are hidden by default.
     * @param hidden if true, then all rows are hidden by default.
     *
     * The default value is false.
     *
     * This parameter is useful if it is much shorter to only write out the rows
     * that are not hidden with #setRowHidden().
     */
    void setRowsHiddenByDefault(bool hidden);
    /**
     * @brief returns the default row height set in the worksheet.
     * @return the default row height in points.
     *
     * The default value is 14.4 (pt).
     */
    double defaultRowHeight() const;
    /**
     * @brief sets the default row height.
     * @param height the default row height in points.
     *
     * The default value is 14.4 (pt).
     *
     * This parameter is useful if most of the rows in a sheet have the same
     * height, but that height isn't the default height (14.4 pt)
     */
    void setDefaultRowHeight(double height);
    /**
     * @brief returns the default column width.
     * @return the default column width in characters of the widest digit of the normal style's font.
     *
     * The default value is 8 5/7.
     */
    double defaultColumnWidth() const;
    /**
     * @brief sets the default column width.
     * @param width value measured in characters of the widest digit of the normal style's font.
     *
     * The default value is 8 5/7.
     */
    void setDefaultColumnWidth(double width);

    // Sheet View parameters

    /**
     * @brief returns whether the panes in the window are locked due to workbook protection.
     *
     * The default value is false.
     * @note Worksheet can have more than one view. This method returns the
     * property of the last one. To get the specific view use #view() method.
     */
    bool isWindowProtected() const;
    /**
     * @brief sets whether the panes in the window are locked due to workbook protection.
     *
     * @param protect Sets protected to the view. The default value is false.
     * @note Worksheet can have more than one view. This method sets the
     * property of the last one. If no sheet views were added, this method
     * adds the default one. To get the specific view use #view() method.
     */
    void setWindowProtected(bool protect);
    bool isFormulasVisible() const; //TODO: doc
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
    bool isDefaultGridColorUsed() const;//TODO: doc
    void setDefaultGridColorUsed(bool value);//TODO: doc
    SheetView::Type viewType() const;//TODO: doc
    void setViewType(SheetView::Type type);//TODO: doc
    /**
     * @brief returns the location of the last added view's top left cell.
     * @return Valid location if it was set, invalid one otherwise.
     */
    CellReference viewTopLeftCell() const;
    void setViewTopLeftCell(const CellReference &ref);//TODO: doc
    int viewColorIndex() const;//TODO: doc
    void setViewColorIndex(int index);//TODO: doc
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

    // Print and page parameters

    //TODO: add methods for print and page parameters

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


    // Autofilter parameters

    AutoFilter &autofilter();
    void clearAutofilter();

private:
    QMap<int, double> getMaximumColumnWidths(int firstRow = 1, int lastRow = INT_MAX);
    void saveToXmlFile(QIODevice *device) const override;
    bool loadFromXmlFile(QIODevice *device) override;
};

}
#endif // XLSXWORKSHEET_H
