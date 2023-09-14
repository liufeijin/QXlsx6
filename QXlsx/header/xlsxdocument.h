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
     * @param row zero-based row number where the image is anchored.
     * @param col zero-based column number where the image is anchored.
     * @return image index (zero-based) on success, -1 otherwise.
     */
    int insertImage(int row, int col, const QImage &image);
    /**
     * @brief returns image from the current active worksheet.
     * @param imageIndex zero-based image index (0 to #getImageCount()-1).
     */
    QImage image(int imageIndex);
    /**
     * @overload
     * @brief returns image from the current active worksheet.
     * @param row zero-based row number where the image is anchored.
     * @param col zero-based column number where the image is anchored.
     */
    QImage image(int row, int col);
    /**
     * @brief returns the count of images on the current active worksheet.
     */
    int imageCount();
    /**
     * @brief sets new content to the image with (zero-based) #index
     * @param index zero-base index of image in the list of all images in the workbook.
     * @param newfile the name of the file from which new content will be loaded.
     * @return true if the image was found and new content was loaded, false otherwise.
     */
    bool changeImage(int index, const QString &fileName); // TODO: remove, as this method uses index differently

    /**
     * @brief removes image from the current active worksheet
     * @param row zero-based row number where the image is anchored.
     * @param column zero-based column number where the image is anchored.
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
    
    Chart *insertChart(int row, int col, const QSize &size);
    
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

    bool defineName(const QString &name, const QString &formula,
                    const QString &comment=QString(), const QString &scope=QString());

    CellRange dimension() const;

    QString documentProperty(const QString &name) const;
    void setDocumentProperty(const QString &name, const QString &property);
    QStringList documentPropertyNames() const;

    QStringList sheetNames() const;
    bool addSheet(const QString &name = QString(),
                  AbstractSheet::Type type = AbstractSheet::Type::Worksheet);
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
