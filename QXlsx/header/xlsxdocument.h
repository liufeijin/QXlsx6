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
 * @brief The Document class provides API that is used to handle the contents of .xlsx files.
 */
class QXLSX_EXPORT Document : public QObject
{
    Q_OBJECT
    Q_DECLARE_PRIVATE(Document) // D-Pointer. Qt classes have a Q_DECLARE_PRIVATE
                                // macro in the public class. The macro reads: qglobal.h
public:
    /**
     * @brief creates an empty Document.
     * @param parent
     */
    explicit Document(QObject *parent = nullptr);
    /**
     * @overload
     * @brief creates Document and reads @a xlsxName file.
     * @param xlsxName File name to read.
     * @param parent
     */
    Document(const QString& xlsxName, QObject* parent = nullptr);
    /**
     * @overload
     * @brief creates Document and reads from @a device.
     * @param device pointer to a device to read from.
     * @param parent
     */
    Document(QIODevice* device, QObject* parent = nullptr);
    ~Document();

    // TODO: remove in future versions
    bool write(const CellReference &cell, const QVariant &value, const Format &format=Format());
    // TODO: remove in future versions
    bool write(int row, int col, const QVariant &value, const Format &format=Format());
    // TODO: remove in future versions
    QVariant read(const CellReference &cell) const;
    // TODO: remove in future versions
    QVariant read(int row, int col) const;

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
    bool addDefinedName(const QString &name,
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
    /**
     * @brief adds new sheet to a workbook and makes it the current active sheet.
     * @param name sheet name. If empty, the default name will be constructed
     * (SheetN for worksheets, ChartN for chartsheets, where N is the sheet number).
     * @param type The new sheet type.
     * @return true if new sheet was successfully added.
     */
    bool addSheet(const QString &name = QString(),
                  AbstractSheet::Type type = AbstractSheet::Type::Worksheet);
    Chartsheet *addChartsheet(const QString &name = QString());
    Worksheet *addWorksheet(const QString &name = QString());
    bool insertSheet(int index, const QString &name = QString(),
                     AbstractSheet::Type type = AbstractSheet::Type::Worksheet);

    bool setActiveSheet(const QString &name);
    bool setActiveSheet(int index);
    bool renameSheet(const QString &oldName, const QString &newName);
    bool copySheet(const QString &srcName, const QString &distName = QString());
    bool moveSheet(const QString &srcName, int distIndex);
    bool deleteSheet(const QString &name);

    Workbook *workbook() const;
    AbstractSheet *sheet(const QString &sheetName) const;
    AbstractSheet *activeSheet() const;
    Worksheet *activeWorksheet() const;

    bool save() const;
    bool saveAs(const QString &xlsXname) const;
    bool saveAs(QIODevice *device) const;

    // copy style from one xlsx file to other
    static bool copyStyle(const QString &from, const QString &to);
    /**
     * @brief returns whether the document was successfully loaded.
     */
    bool isLoaded() const;

private:
    Q_DISABLE_COPY(Document) // Disables the use of copy constructors and
                             // assignment operators for the given Class.
    DocumentPrivate* const d_ptr;
};

}

#endif // QXLSX_XLSXDOCUMENT_H
