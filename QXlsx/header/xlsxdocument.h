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
class ConditionalFormatting;
class Chart;
class CellReference;
class DocumentPrivate;
class DefinedName;
class Chartsheet;

/**
 * @brief The Document class provides API to handle the contents of .xlsx files.
 *
 * This class is used as an entry point to the xlsx document. It has methods to
 * #save() and #load() xlsx file, and to some extent to manipulate sheets and
 * their contents (internally through the #workbook() pointer).
 *
 * See [Demo example](../../examples/Demo/demo.cpp).
 *
 * Each document has a pointer to #workbook(). The Workbook class presents methods
 * to add, delete and get sheets, to manupulate the workbook properties etc.
 *
 */
class QXLSX_EXPORT Document : public QObject
{
    Q_OBJECT
    Q_DECLARE_PRIVATE(Document) // D-Pointer. Qt classes have a Q_DECLARE_PRIVATE
                                // macro in the public class. The macro reads: qglobal.h
public:
    /**
     * @brief The Metadata enum specifies the document metadata that the user can change.
     */
    enum class Metadata
    {
        Category, /**< Core string property */
        ContentStatus, /**< Core string property */
        Created, /**< Core date/time property. By default QXlsx writes the current date/time */
        Creator, /**< Core string property. By default QXlsx writes "QXlsx Library" */
        Description, /**< Core string property */
        Identifier, /**< Core string property */
        Keywords, /**< Core string property */
        Language, /**< Core string property */
        LastModifiedBy, /**< Core string property. By default QXlsx writes "QXlsx Library" */
        LastPrinted, /**< Core date/time property. */
        Revision, /**< Core string property */
        Subject, /**< Core string property */
        Title, /**< Core string property */
        Version, /**< Core string property */
        Application, /**< Extended string property. By default QXlsx writes "Microsoft Excel" */
        AppVersion, /**< Extended string property. By default QXlsx writes "12.0000" */
        Company, /**< Extended string property */
        Manager /**< Extended string property */
    };
    /**
     * @brief creates an empty Document.
     * @param parent
     */
    explicit Document(QObject *parent = nullptr);
    /**
     * @overload
     * @brief creates Document and optionally reads @a name file.
     * @param name File name to read.
     * @param loadImmediately if `true`, then the document will be loaded on
     * creation. If `false`, use #load() method to read @a name.
     * @param parent
     */
    Document(const QString& name, bool loadImmediately = true, QObject* parent = nullptr);
    /**
     * @overload
     * @brief creates Document and reads from @a device.
     * @param device pointer to a readable device to read from.
     * @param parent this object parent.
     */
    Document(QIODevice* device, QObject* parent = nullptr);
    ~Document();

    /**
     * @brief saves the current document. If no name was specified with #saveAs(),
     * saves under the default-constructed name ("Book1.xlsx").
     * @return `true` on success.
     */
    bool save() const;
    /**
     * @brief saves the current document.
     * @param name The document name. If the file @a name already exists,
     * it will be overwritten.
     * @return `true` on success.
     */
    bool saveAs(const QString &name) const;
    /**
     * @brief writes the current document to the @a device.
     * @param device the pointer to the (writable) device.
     * @return `true` on success.
     */
    bool saveAs(QIODevice *device) const;

    // copy style from one xlsx file to other
    //    static bool copyStyle(const QString &from, const QString &to);

    /**
     * @brief returns whether the document was successfully loaded.
     *
     * The actual loading occurs in the document constructor.
     */
    bool isLoaded() const;
    /**
     * @brief loads the document contents.
     * @return  `true` on success.
     */
    bool load();


    // TODO: remove in future versions
    /**
     * @brief writes @a value to the active worksheet at @a cell and formats @a
     * cell with @a format.
     * @param cell the cell to write to.
     * @param value data to write.
     * @param format format to apply.
     * @return `true` if writing was successful.
     *
     * Equivalent to
     * ```cpp
     * if (auto worksheet = doc.activeWorksheet())
     *     worksheet->write(cell, value, format);
     * ```
     */
    bool write(const CellReference &cell, const QVariant &value, const Format &format=Format());
    // TODO: remove in future versions
    /**
     * @brief writes @a value to the active worksheet at (@a row, @a column) and
     * formats it with @a format.
     * @param row the row number of the cell (starting from 1).
     * @param column the column number of the cell (starting from 1).
     * @param value data to write.
     * @param format format to apply.
     * @return `true` if writing was successful.
     *
     * Equivalent to
     * ```cpp
     * if (auto worksheet = doc.activeWorksheet())
     *     worksheet->write(row, column, value, format);
     * ```
     */
    bool write(int row, int column, const QVariant &value, const Format &format=Format());
    // TODO: remove in future versions
    /**
     * @brief Reads the cell data from the active worksheet and returns it as a
     * QVariant object.
     * @param cell The cell to read from.
     * @return QVariant object, that may contain a number, a string, a date/time
     * etc.
     * @note If @a cell contains a shared/array formula, and this cell is not a
     * 'master' formula cell, this method recalculates the formula references to
     * the ones related to @a cell.
     *
     * Equivalent to
     * ```cpp
     * if (auto worksheet = doc.activeWorksheet())
     *     return worksheet->read(cell);
     * ```
     *
     * @note Data returned may differ from data stored in the document. Compare:
     * ```cpp
     * doc.write( "A3", QVariant(QDate(2019, 10, 9)) );
     * qDebug() << doc.read(3, 1).type() << doc.read(3, 1); // QVariant::QDate QVariant(QDate, QDate("2019-10-09"))
     * qDebug() << doc.activeWorksheet()->cell(3, 1)->value().type()
     *          << doc.activeWorksheet()->cell(3, 1)->value(); // QVariant::double QVariant(double, 43747)
     * ```
     *
     */
    QVariant read(const CellReference &cell) const;
    // TODO: remove in future versions
    /**
     * @overload
     * @brief Reads the cell data from the active worksheet and returns it as a
     * QVariant object.
     * @param row The cell row (starting from 1).
     * @param column The cell column (starting from 1).
     * @return QVariant object, that may contain a number, a string, a date/time
     * etc.
     * @note If the cell contains a shared/array formula, and this cell is not a
     * 'master' formula cell, this method recalculates the formula references to
     * the ones related to @a cell.
     *
     * Equivalent to
     * ```cpp
     * if (auto worksheet = doc.activeWorksheet())
     *     return worksheet->read(row, column);
     * ```
     *
     * @note Data returned may differ from data stored in the document. Compare:
     * ```cpp
     * doc.write( "A3", QVariant(QDate(2019, 10, 9)) );
     * qDebug() << doc.read(3, 1); // QVariant(QDate, QDate("2019-10-09"))
     * qDebug() << doc.activeWorksheet()->cell(3, 1)->value(); // QVariant(double, 43747)
     * //dates are stored as numeric values. See Workbook::date1904() docs for more help.
     * ```
     */
    QVariant read(int row, int column) const;


    /// Document metadata properties

    /**
     * @brief returns the metadata property associated with the document.
     * @param property The metadata to retrieve.
     * @return Valid QVariant object if @a property was set.
     */
    QVariant metadata(Metadata property) const;
    /**
     * @brief sets the metadata property associated with the document.
     * @param property The metadata to set.
     * @param value The metadata value.
     */
    void setMetadata(Metadata property, const QVariant &value);
    /**
     * @brief returns whether the document has metadata @a property associated with it.
     * @param property The metadata to test.
     * @return `true` if @a property was set.
     * @note The following metadata properties get the default values if they were not
     * set by the moment the document is being written:
     * - Metadata::Created (gets the current date/time)
     * - Metadata::Creator (gets 'QXlsx library')
     * - Metadata::LastModifiedBy (gets 'QXlsx library')
     * - Metadata::Application (gets 'Microsoft Excel')
     * - Metadata::AppVersion (gets '12.0000')
     */
    bool hasMetadata(Metadata property) const;
    /**
     * @brief returns the map of all metadata associated with the document.
     */
    QMap<Metadata, QVariant> allMetadata() const;

    /**
     * @brief returns the list of sheets names.
     */
    QStringList sheetNames() const;
    /**
     * @brief adds new sheet to a workbook and makes it the current active sheet.
     * @param name sheet name. If empty, the default name will be constructed
     * ('Sheet#' for worksheets, 'Chart#' for chartsheets, where # is the sequential
     * sheet number).
     * @param type The optional new sheet type.
     * @return `true` if new sheet was successfully added.
     */
    bool addSheet(const QString &name = QString(),
                  AbstractSheet::Type type = AbstractSheet::Type::Worksheet);
    /**
     * @brief adds new chartsheet to a workbook and makes it the current active sheet.
     * @param name sheet name. If empty, the default name will be constructed
     * ('Chart#', where # is the sequential sheet number).
     * @return pointer to the chartsheet if new chartsheet was successfully added,
     * `nullptr` otherwise.
     */
    Chartsheet *addChartsheet(const QString &name = QString());
    /**
     * @brief adds new worksheet to a workbook and makes it the current active sheet.
     * @param name sheet name. If empty, the default name will be constructed
     * ('Sheet#', where # is the sequential sheet number).
     * @return pointer to the worksheet if new worksheet was successfully added,
     * `nullptr` otherwise.
     */
    Worksheet *addWorksheet(const QString &name = QString());
    /**
     * @brief inserts new sheet at @a index and makes it the current active sheet.
     * @param index index of the new sheet (0 to #sheetsCount()).
     * @param name sheet name. If empty, the default name will be constructed
     * ('Chart#', where # is the sequential sheet number).
     * @param type The optional sheet type.
     * @return `true` if new sheet was successfully inserted.
     */
    bool insertSheet(int index, const QString &name = QString(),
                     AbstractSheet::Type type = AbstractSheet::Type::Worksheet);
    /**
     * @brief returns the sheets count (both worksheets and chartsheets).
     *
     * Equivalent to `workbook()->sheetsCount()`.
     */
    int sheetsCount() const;
    /**
     * @brief sets the active (current) sheet.
     * @param name the name of the sheet.
     * @return `true` if the sheet with @a name was found and set active.
     */
    bool setActiveSheet(const QString &name);
    /**
     * @brief sets the active (current) sheet.
     * @param index the sheet index (0 to #sheetsCount()-1).
     * @return `true` if the sheet with @a index was found and set active.
     */
    bool setActiveSheet(int index);
    /**
     * @brief renames the sheet.
     * @param oldName the old name.
     * @param newName the new name.
     * @return `true` on success.
     */
    bool renameSheet(const QString &oldName, const QString &newName);
    /**
     * @brief copies the sheet named @a srcName to the new sheet with name @a dstName.
     * @param srcName the source sheet name.
     * @param dstName the destination sheet name.
     * @return `true` on success.
     */
    bool copySheet(const QString &srcName, const QString &dstName = QString());
    /**
     * @brief moves the sheet named @a sheetName to @a dstIndex.
     * @param sheetName the sheet to move.
     * @param dstIndex the new index.
     * @return `true` on success.
     */
    bool moveSheet(const QString &sheetName, int dstIndex);
    /**
     * @brief moves the sheet from @a srcIndex to @a dstIndex.
     * @param srcIndex the old index.
     * @param dstIndex the new index.
     * @return `true` on success.
     */
    bool moveSheet(int srcIndex, int dstIndex);
    /**
     * @brief deletes the sheet @a name.
     * @param name The name of the sheet to delete.
     * @return `true` on success.
     */
    bool deleteSheet(const QString &name);
    /**
     * @brief deletes the sheet at @a index.
     * @param index the sheet index (from 0 to #sheetsCount()-1).
     * @return `true` on success.
     */
    bool deleteSheet(int index);
    /**
     * @brief returns the workbook.
     * @return pointer to the workbook.
     */
    Workbook *workbook() const;
    /**
     * @brief returns the sheet by its @a sheetName.
     * @param sheetName a string that contains the sheet name.
     * @return pointer to the sheet or `nullptr` if there is no sheet with
     * @a sheetName.
     */
    AbstractSheet *sheet(const QString &sheetName) const;
    /**
     * @brief returns the active (current) sheet.
     * @return pointer to the active sheet or `nullptr` if there is no sheets in
     * the workbook.
     */
    AbstractSheet *activeSheet() const;
    /**
     * @brief returns the active (current) worksheet.
     * @return pointer to the active worksheet or `nullptr` if there is no
     * worksheets or the active sheet is a chartsheet.
     */
    Worksheet *activeWorksheet() const;
    /**
     * @brief returns the active (current) chartsheet.
     * @return pointer to the active chartsheet or `nullptr` if there is no
     * chartsheets or the active sheet is a worksheet.
     */
    Chartsheet *activeChartsheet() const;
private:
    Q_DISABLE_COPY(Document) // Disables the use of copy constructors and
                             // assignment operators for the given Class.
    DocumentPrivate* const d_ptr;
};

}

#endif // QXLSX_XLSXDOCUMENT_H
