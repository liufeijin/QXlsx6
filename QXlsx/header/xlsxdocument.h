// xlsxdocument.h

#ifndef QXLSX_XLSXDOCUMENT_H
#define QXLSX_XLSXDOCUMENT_H

#include <QtGlobal>
#include <QObject>
#include <QVariant>
#include <QIODevice>
#include <QImage>

#include "xlsxglobal.h"
#include "xlsxformat.h"
#include "xlsxworksheet.h"

namespace QXlsx {

class Workbook;
class Cell;
class CellRange;
class DataValidation;
class ConditionalFormatting;
class Chart;
class CellReference;
class DocumentPrivate;
class DefinedName;

class QXLSX_EXPORT Document : public QObject
{
    Q_OBJECT
    Q_DECLARE_PRIVATE(Document) // D-Pointer. Qt classes have a Q_DECLARE_PRIVATE
                                // macro in the public class. The macro reads: qglobal.h
public:
    explicit Document(QObject *parent = nullptr);
    Document(const QString& xlsxName, QObject* parent = nullptr);
    Document(QIODevice* device, QObject* parent = nullptr);
    ~Document();

    bool write(const CellReference &cell, const QVariant &value, const Format &format=Format());
    bool write(int row, int col, const QVariant &value, const Format &format=Format());
    
    QVariant read(const CellReference &cell) const;
    QVariant read(int row, int col) const;
    
    /*!
     * @brief Inserts an \a image to current active worksheet at the position \a row, \a column.
     * @param row 1-based row number where the image is anchored.
     * @param col 1-based column number where the image is anchored.
     * @param image the image to insert.
     * @return image index (zero-based) on success, -1 otherwise.
     */
    int insertImage(int row, int col, const QImage &image);
    /**
     * @brief returns image from the current active worksheet.
     * @param imageIndex zero-based image index (0 to #imageCount()-1).
     */
    QImage image(int imageIndex);
    /**
     * @overload
     * @brief returns image from the current active worksheet.
     * @param row 1-based row number where the image is anchored.
     * @param col 1-based column number where the image is anchored.
     */
    QImage image(int row, int col);
    /**
     * @brief returns the count of images on the current active worksheet.
     */
    int imageCount();
    /**
     * @brief sets new content to the image with (zero-based) @a index.
     * @param index zero-based index of image (0 to #imageCount()-1).
     * @param fileName the name of the file from which new content will be loaded.
     * @param keepSize if true, then new image will be resized to the old image size.
     * If false, then new image will have its size.
     * @return true if the image was found and new content was loaded, false otherwise.
     */
    bool changeImage(int index, const QString &fileName, bool keepSize = true);

    /**
     * @brief removes image from the current active worksheet
     * @param row 1-based row number where the image is anchored.
     * @param column 1-based column number where the image is anchored.
     * @return true if image was found and removed, false otherwise.
     */
    bool removeImage(int row, int column);
    /**
     * @overload
     * @brief removes image from the current active worksheet
     * @param index zero-based index of the image (0 to #imageCount()-1).
     * @return true if image was found and removed, false otherwise.
     */
    bool removeImage(int index);

    /**
     * @brief sets the current sheet's background image.
     * @param image The image will be written as a PNG image.
     */
    void setBackgroundImage(const QImage &image);
    /**
     * @brief sets the current sheet's background image from a file.
     * @overload
     * @param fileName the file name.
     * @note If the image is a PNG or JPG image, it will be written as a PNG or JPG respectively.
     * All other pictures will be converted to PNG.
     */
    void setBackgroundImage(const QString &fileName);
    /**
     * @brief returns the current sheet's background image.
     * @return Valid QImage if the background image was set, null QImage otherwise.
     */
    QImage backgroundImage() const;
    /**
     * @brief removes the current active sheet background image.
     * @return true if the background image was removed.
     */
    bool removeBackgroundImage();

    /**
     * @brief creates a new chart and places it inside the current active worksheet.
     * @param row the 1-based row index of the chart top left corner.
     * @param column the 1-based column index of the chart top left corner.
     * @param size the chart size in pixels at 96 dpi.
     * @return pointer to the new chart or nullptr if no chart was created.
     */
    Chart *insertChart(int row, int column, const QSize &size);
    /**
     * @brief creates a new chart and places it inside the current active worksheet.
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
     * @brief returns chart from the current active worksheet.
     * @param index zero-based index of the chart (0 to #chartCount()-1)
     * @return valid pointer to the chart if the chart was found, nullptr otherwise.
     */
    Chart *chart(int index) const;
    /**
     * @overload
     * @brief returns chart from the current active worksheet.
     * @param row the 1-based row index of the chart top left corner.
     * @param column the 1-based column index of the chart top left corner.
     * @return valid pointer to the chart if the chart was found, nullptr otherwise.
     */
    Chart *chart(int row, int column) const;
    /**
     * @brief removes chart from the current active worksheet.
     * @param row the 1-based row index of the chart top left corner.
     * @param column the 1-based column index of the chart top left corner.
     * @return true if the chart was found and successfully removed, false otherwise.
     */
    bool removeChart(int row, int column);
    /**
     * @overload
     * @brief removes chart from the current active worksheet.
     * @param index zero-based index of the chart.
     * @return true if the chart was found and successfully removed, false otherwise.
     */
    bool removeChart(int index);
    /**
     * @brief returns count of charts on the current active worksheet or 0 if there's no current active worksheet.
     */
    int chartCount() const;
    /**
    @brief Merges a @a range of cells in the current active worksheet.

    The first cell will retain its data, other cells will be cleared.
    All cells will get the same @a format if a valid @a format is given.
    Returns true on success.
    */
    bool mergeCells(const CellRange &range, const Format &format=Format());
    bool unmergeCells(const CellRange &range);

    bool setColumnWidth(const CellRange &range, double width);
    bool setColumnFormat(const CellRange &range, const Format &format);
    bool setColumnHidden(const CellRange &range, bool hidden);
    bool setColumnWidth(int column, double width);
    bool setColumnFormat(int column, const Format &format);
    bool setColumnHidden(int column, bool hidden);
    bool setColumnWidth(int colFirst, int colLast, double width);
    bool setColumnFormat(int colFirst, int colLast, const Format &format);
    bool setColumnHidden(int colFirst, int colLast, bool hidden);
    
    double columnWidth(int column);
    Format columnFormat(int column);
    bool isColumnHidden(int column);

    bool setRowHeight(int row, double height);
    bool setRowFormat(int row, const Format &format);
    bool setRowHidden(int row, bool hidden);
    bool setRowHeight(int rowFirst, int rowLast, double height);
    bool setRowFormat(int rowFirst, int rowLast, const Format &format);
    bool setRowHidden(int rowFirst, int rowLast, bool hidden);

    double rowHeight(int row);
    Format rowFormat(int row);
    bool isRowHidden(int row);

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
    
    bool addDataValidation(const DataValidation &validation);
    bool addConditionalFormatting(const ConditionalFormatting &cf);

    Cell *cellAt(const CellReference &cell) const;
    Cell *cellAt(int row, int col) const;

    /**
     * @brief Adds a defined name in the workbook.
     *
     * Defined names are descriptive names to represent cells, ranges of cells, formulas, or
     * constant values. Defined names can be used to represent a range on any worksheet.
     * @param name The defined name
     * @param formula The value, the cell or range that the defined name refers to.
     * @param scope The name of a worksheet which @a formula refers to, or empty which means global scope.
     * @return pointer to the added defined name. It may be null if @a name or @a formula is empty or there
     * is already a definedName with @a name.
     * @note Even if you specify the sheet name in @a scope, Excel will treat @a formula without
     * the sheet name part as invalid. LibreOffice is OK with that. Example:
     * @code
     * xlsx.addDefinedName("MyCol_3", "=Sheet1!$C$1:$C$10", "Sheet1"); //The only way for Excel
     * //vs
     * xlsx.addDefinedName("MyCol_3", "=$C$1:$C$10", "Sheet1"); //OK for LibreOffice, error for Excel
     * @endcode
     */
    DefinedName *addDefinedName(const QString &name,
                                const QString &formula,
                                const QString &scope = QString());
    /**
     * @brief removes a defined name from the workbook.
     *
     * Defined names are descriptive names to represent cells, ranges of cells, formulas, or
     * constant values. Defined names can be used in formulas to represent a range on any worksheet.
     *
     * @param name The defined name
     * @return true if @a name was found and successfully removed, false otherwise.
     */
    bool removeDefinedName(const QString &name);
    /**
     * @overload
     * @brief removes a defined name from the workbook.
     *
     * Defined names are descriptive names to represent cells, ranges of cells, formulas, or
     * constant values. Defined names can be used in formulas to represent a range on any worksheet.
     *
     * @param name The defined name
     * @return true if @a name was found and successfully removed, false otherwise.
     */
    bool removeDefinedName(DefinedName *name);
    /**
     * @brief returns whether the workbook has a specific defined name.
     *
     * Defined names are descriptive names to represent cells, ranges of cells, formulas, or
     * constant values. Defined names can be used in formulas to represent a range on any worksheet.
     * @param name The defined name
     * @return true if the workbook has a defined @a name.
     */
    bool hasDefinedName(const QString &name) const;
    /**
     * @brief returns a list of defined names.
     * @return a list of defined names.
     */
    QStringList definedNames() const;
    /**
     * @brief returns a defined name by its name.
     * @param name A name to search.
     * @return A pointer to the defined name. May be null if there's no such defined name in the workbook.
     */
    DefinedName *definedName(const QString &name);




    CellRange dimension() const;

    QString documentProperty(const QString &name) const;
    /**
        @brief Set the document properties such as Title, Author etc.

        The method can be used to set the document properties of the Excel
        file created by Qt Xlsx. These properties are visible when you use the
        Office Button -> Prepare -> Properties option in Excel and are also
        available to external applications that read or index windows files.

        The property @a name that can be set are:

        - title
        - subject
        - creator
        - manager
        - company
        - category
        - keywords
        - description
        - status
    */
    void setDocumentProperty(const QString &name, const QString &property);
    QStringList documentPropertyNames() const;

    QStringList sheetNames() const;
    bool addSheet(const QString &name = QString(),
                  AbstractSheet::Type type = AbstractSheet::Type::Worksheet);
    Worksheet *addWorksheet(const QString &name = QString());
    bool insertSheet(int index, const QString &name = QString(),
                     AbstractSheet::Type type = AbstractSheet::Type::Worksheet);
    bool selectSheet(const QString &name);
    bool selectSheet(int index);
    bool renameSheet(const QString &oldName, const QString &newName);
    bool copySheet(const QString &srcName, const QString &distName = QString());
    bool moveSheet(const QString &srcName, int distIndex);
    bool deleteSheet(const QString &name);

    Workbook *workbook() const;
    AbstractSheet *sheet(const QString &sheetName) const;
    AbstractSheet *currentSheet() const;
    Worksheet *currentWorksheet() const;

    bool save() const;
    bool saveAs(const QString &xlsXname) const;
    bool saveAs(QIODevice *device) const;

    // copy style from one xlsx file to other
    static bool copyStyle(const QString &from, const QString &to);
    /**
     * @brief returns whether the document was successfully loaded.
     */
    bool isLoaded() const;

    bool autosizeColumnWidth(const CellRange &range);
    bool autosizeColumnWidth(int column);
    bool autosizeColumnWidth(int firstColumn, int lastColumn);
    bool autosizeColumnWidth();

private:
    Q_DISABLE_COPY(Document) // Disables the use of copy constructors and
                             // assignment operators for the given Class.
    DocumentPrivate* const d_ptr;
};

}

#endif // QXLSX_XLSXDOCUMENT_H
