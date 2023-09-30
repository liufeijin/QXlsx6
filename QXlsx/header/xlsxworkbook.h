// xlsxworkbook.h

#ifndef XLSXWORKBOOK_H
#define XLSXWORKBOOK_H

#include <QtGlobal>
#include <QList>
#include <QImage>
#include <QSharedPointer>
#include <QIODevice>

#include <memory>

#include "xlsxglobal.h"
#include "xlsxabstractooxmlfile.h"
#include "xlsxabstractsheet.h"

namespace QXlsx {

class SharedStrings;
class Styles;
class Drawing;
class Document;
class Theme;
class Relationships;
class DocumentPrivate;
class MediaFile;
class Chart;
class Chartsheet;
class Worksheet;
class WorkbookPrivate;

struct QXLSX_EXPORT DefinedName
{
    //TODO: convert to shallow-copyable and move to a separate header.
    DefinedName()
    {}
    DefinedName(const QString &name, const QString &formula, int sheetId=-1)
        :name(name), formula(formula), sheetId(sheetId)
    {
    }
    QString name; /**< The required (non-empty) defined name. */
    QString formula; /**< Formula to be represented with this name. */
    QString comment; /**< A comment to this name. */
    QString description; /**< A description of this name. */
    std::optional<bool> hidden; /**< Whether this defined name is hidden
from the program interface. If true, then the defined name will not be shown in the defined names
list (in Excel). The default value is false.*/
    std::optional<bool> workbookParameter; /**<  default = false */

    int sheetId = -1; //using internal sheetId, instead of the localSheetId(order in the workbook)
};

class QXLSX_EXPORT Workbook : public AbstractOOXmlFile
{
    Q_DECLARE_PRIVATE(Workbook)
public:
    ~Workbook();

    int sheetCount() const;
    AbstractSheet *sheet(int index) const;

    AbstractSheet *addSheet(const QString &name = QString(), AbstractSheet::Type type = AbstractSheet::Type::Worksheet);
    AbstractSheet *insertSheet(int index, const QString &name = QString(), AbstractSheet::Type type = AbstractSheet::Type::Worksheet);
    bool renameSheet(int index, const QString &name);
    bool renameSheet(const QString &oldName, const QString &newName);
    bool deleteSheet(int index);
    bool copySheet(int index, const QString &newName=QString());
    bool moveSheet(int srcIndex, int distIndex);

    AbstractSheet *activeSheet() const;
    bool setActiveSheet(int index);


    // Defined names

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


    bool isDate1904() const;
    /*!
    Excel for Windows uses a default epoch of 1900 and Excel
    for Mac uses an epoch of 1904. However, Excel on either
    platform will convert automatically between one system
    and the other. Qt Xlsx stores dates in the 1900 format
    by default.

    \note This function should be called before any date/time
    has been written.
    */
    void setDate1904(bool date1904);
    bool isStringsToNumbersEnabled() const;
    /**
    Enables the worksheet.write() method to convert strings
    to numbers, where possible, using float() in order to avoid
    an Excel warning about "Numbers Stored as Text".

    The default is false
    */
    void setStringsToNumbersEnabled(bool enable=true);
    bool isStringsToHyperlinksEnabled() const;
    void setStringsToHyperlinksEnabled(bool enable=true);
    bool isHtmlToRichStringEnabled() const;
    void setHtmlToRichStringEnabled(bool enable=true);
    QString defaultDateFormat() const;
    void setDefaultDateFormat(const QString &format);

    //internal used member
    bool addMediaFile(QSharedPointer<MediaFile> media, bool force=false);
    void removeMediaFile(QSharedPointer<MediaFile> media);
    QList<QWeakPointer<MediaFile> > mediaFiles() const;
    void addChartFile(const QSharedPointer<Chart> &chartFile);
    void removeChartFile(const QSharedPointer<Chart> &chart);
    QList<QWeakPointer<Chart> > chartFiles() const;

private:
    friend class Worksheet;
    friend class Chartsheet;
    friend class WorksheetPrivate;
    friend class Document;
    friend class DocumentPrivate;

    Workbook(Workbook::CreateFlag flag);

    void saveToXmlFile(QIODevice *device) const override;
    bool loadFromXmlFile(QIODevice *device) override;

    SharedStrings *sharedStrings() const;
    Styles *styles();
    Theme *theme();
    QList<QImage> images();
    QList<Drawing *> drawings();
    QList<QSharedPointer<AbstractSheet> > getSheetsByType(AbstractSheet::Type type) const;
    QStringList worksheetNames() const;
    AbstractSheet *addSheet(const QString &name, int sheetId, AbstractSheet::Type type);
};

}

#endif // XLSXWORKBOOK_H
