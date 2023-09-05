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
    
    int insertImage(int row, int col, const QImage &image);
    bool getImage(int imageIndex, QImage& img);
    bool getImage(int row, int col, QImage& img);
    uint getImageCount();
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

    bool isLoadPackage() const; 
    bool load() const; // equals to isLoadPackage()

    bool changeimage(int filenoinmidea,QString newfile); // add by liufeijin20181025

    bool autosizeColumnWidth(const CellRange &range);
    bool autosizeColumnWidth(int column);
    bool autosizeColumnWidth(int colFirst, int colLast);
    bool autosizeColumnWidth(void);

private:
    QMap<int, int> getMaximalColumnWidth(int firstRow=1, int lastRow=INT_MAX);


private:
    Q_DISABLE_COPY(Document) // Disables the use of copy constructors and
                             // assignment operators for the given Class.
    DocumentPrivate* const d_ptr;
};

}

#endif // QXLSX_XLSXDOCUMENT_H
