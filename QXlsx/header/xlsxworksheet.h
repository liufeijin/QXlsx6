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

    Cell *cell(const CellReference &row_column) const;
    Cell *cell(int row, int column) const;

    /*!
     * @brief Inserts an \a image at the position \a row, \a column.
     * @return image index (zero-based) on success, -1 otherwise.
     */
    int insertImage(int row, int column, const QImage &image);
    /**
     * @brief returns an image.
     * @param imageIndex zero-based image index (0 to #getImageCount()-1).
     * @return non-null QImage if image was found and read.
     */
    QImage image(int imageIndex);
    /**
     * @brief returns image.
     * @param row worksheet row number (starting from 1) where the image is anchored.
     * @param col worksheet column number (starting from 1) where the image is anchored.
     * @return non-null QImage if image was found and read.
     */
    QImage image(int row, int column);
    /**
     * @brief returns the count of images on this worksheet.
     */
    int imageCount();
    //TODO: removeImage(int row, int column), removeImage(int index)

    /**
     * @brief creates a new chart and places it inside the current worksheet.
     * @param row the row index of the chart top left corner.
     * @param column the column index of the chart top left corner.
     * @param size the chart size in pixels.
     * @return pointer to the new chart or nullptr if no chart was created.
     */
    Chart *insertChart(int row, int column, const QSize &size);
    //TODO: removeChart(int row, int column), chart(int row, int column)
    //TODO: removeChart(int index), chart(int index)

    bool mergeCells(const CellRange &range, const Format &format=Format());
    bool unmergeCells(const CellRange &range);
    QList<CellRange> mergedCells() const;

    bool setColumnWidth(const CellRange& range, double width);
    bool setColumnFormat(const CellRange& range, const Format &format);
    bool setColumnHidden(const CellRange& range, bool hidden);
    bool setColumnWidth(int colFirst, int colLast, double width);
    bool setColumnFormat(int colFirst, int colLast, const Format &format);
    bool setColumnHidden(int colFirst, int colLast, bool hidden);

    double columnWidth(int column);
    Format columnFormat(int column);
    bool isColumnHidden(int column);

    bool setRowHeight(int rowFirst,int rowLast, double height);
    bool setRowFormat(int rowFirst,int rowLast, const Format &format);
    bool setRowHidden(int rowFirst,int rowLast, bool hidden);

    double rowHeight(int row);
    Format rowFormat(int row);
    bool isRowHidden(int row);

    bool groupRows(int rowFirst, int rowLast, bool collapsed = true);
    bool groupColumns(int colFirst, int colLast, bool collapsed = true);
    bool groupColumns(const CellRange &range, bool collapsed = true);
    CellRange dimension() const;

    bool isWindowProtected() const;
    void setWindowProtected(bool protect);
    bool isFormulasVisible() const;
    void setFormulasVisible(bool visible);
    bool isGridLinesVisible() const;
    void setGridLinesVisible(bool visible);
    bool isRowColumnHeadersVisible() const;
    void setRowColumnHeadersVisible(bool visible);
    bool isZerosVisible() const;
    void setZerosVisible(bool visible);
    bool isRightToLeft() const;
    void setRightToLeft(bool enable);
    bool isSelected() const;
    void setSelected(bool select);
    bool isRulerVisible() const;
    void setRulerVisible(bool visible);
    bool isOutlineSymbolsVisible() const;
    void setOutlineSymbolsVisible(bool visible);
    bool isWhiteSpaceVisible() const;
    void setWhiteSpaceVisible(bool visible);

    QMap<CellReference, std::shared_ptr<Cell> > getFullCells(int* maxRow, int* maxCol);

private:
    void saveToXmlFile(QIODevice *device) const override;
    bool loadFromXmlFile(QIODevice *device) override;
};

}
#endif // XLSXWORKSHEET_H
