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
#include "xlsxutility_p.h"

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

/**
 * @brief The DefinedName struct represents properties of a defined name attached
 * to a workbook.
 *
 * Defined names are descriptive names to represent cells, ranges of cells, formulas, or
 * constant values. Defined names can be used to represent a range on any worksheet.
 *
 *
 */
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
from the program interface. If true, then the defined name will not be shown in
the defined names list (in Excel). The default value is false.*/
    std::optional<bool> workbookParameter; /**<  default = false */

    int sheetId = -1; //using internal sheetId, instead of the localSheetId(order in the workbook)
};


/**
 * @brief The Workbook class represents a collection of sheets and provides methods
 * to add, remove, copy sheets, to manupulate the workbook properties etc.
 */
class QXLSX_EXPORT Workbook : public AbstractOOXmlFile
{
    Q_DECLARE_PRIVATE(Workbook)
public:
    /**
     * @brief The ShowObjects enum specifies  how the application displays objects
     * in this workbook. Objects might include charts, images, and other object
     * data that the application supports.
     */
    enum class ShowObjects {
        All, /**< All objects are shown in the workbook (default). */
        Placeholders, /**< The application show placeholders for objects in the workbook. */
        None /**< All objects are hidden in the workbook. */
    };
    /**
     * @brief The UpdateLinks enum specifies when the application updates links
     * to other workbooks when the workbook is opened.
     */
    enum class UpdateLinks {
        UserSet, /**< The end-user specifies whether they receive an alert to update
links to other workbooks when the workbook is opened (default).*/
        Never, /**< Links to other workbooks are never updated when the workbook is opened. The
application will not display an alert in the user interface.*/
        Always /**< Links to other workbooks are always updated when the workbook is opened. The
application will not display an alert in the user interface.*/
    };
    /**
     * @brief The CalculationMode enum specifies when the application should
     * calculate formulas in the workbook.
     */
    enum class CalculationMode
    {
        Manual, /**< Calculations in the workbook are triggered manually by the user. */
        Auto, /**< Calculations in the workbook are performed automatically when
cell values change. The application recalculates those cells that are dependent
on other cells that contain changed values. This mode of calculation helps to avoid
unnecessary calculations. (default.) */
        AutoNoTable /**< Calculations in the workbook are performed automatically when
cell values change. Tables are excluded during calculation. */
    };
    /**
     * @brief The ReferenceMode enum defines the supported reference styles or
     * modes for a workbook.
     */
    enum class ReferenceMode
    {
        A1, /**< Workbook uses A1 reference style (default). A1 reference style
refers to columns with letters and refers to rows with numbers. For example, A1
refers to the cell at the intersection of column A and row 1. */
        R1C1 /**< Workbook uses the R1C1 reference style. R1C1 reference style
refers to both the rows and the columns on the worksheet with numbers. The location
of a cell is indicated with an "R" followed by a row number and a "C" followed by
a column number. For example, R1C1 refers to the cell at the intersection
of row R1 and column C1. */
    };

    ~Workbook();
    /**
     * @brief returns the sheets count (both worksheets and chartsheets).
     */
    int sheetsCount() const;
    /**
     * @brief returns the count of sheets of specified @a type
     * @param type sheets type.
     */
    int sheetsCount(AbstractSheet::Type type) const;
    /**
     * @brief returns the list of sheets names (both worksheets and chartsheets).
     */
    QStringList sheetNames() const;
    /**
     * @brief returns the list of sheets of @a type.
     */
    QList<AbstractSheet*> sheets(AbstractSheet::Type type) const;
    /**
     * @brief returns the list of sheets (both worksheets and chartsheets).
     */
    QList<AbstractSheet*> sheets() const;
    /**
     * @brief returns the list of worksheets.
     */
    QList<Worksheet*> worksheets() const;
    /**
     * @brief returns the list of chartsheets
     */
    QList<Chartsheet*> chartsheets() const;
    /**
     * @brief returns the sheet with @a index.
     * @param index the sheet index (0 to #sheetsCount()-1).
     * @return pointer to the sheet or `nullptr` if @a index is invalid.
     */
    AbstractSheet *sheet(int index) const;
    /**
     * @overload
     * @brief returns the sheet with @a name.
     * @param name the sheet name.
     * @return pointer to the sheet or `nullptr` if no sheet with @a name exists.
     */
    AbstractSheet *sheet(const QString &name) const;
    /**
     * @brief adds new sheet to the workbook.
     * @param name optional sheet name. If empty, new sheet will have the name
     * Sheet# or Chart# depending on @a type.
     * @param type optional sheet type.
     * @return pointer to the new sheet on success, `nullptr` otherwise.
     */
    AbstractSheet *addSheet(const QString &name = QString(), AbstractSheet::Type type = AbstractSheet::Type::Worksheet);
    /**
     * @brief adds new chartsheet to the workbook.
     * @param name optional sheet name. If empty, new sheet will have the name
     * Chart# where # is the sequential chartsheet number.
     * @return pointer to the new chartsheet on success, `nullptr` otherwise.
     */
    Chartsheet *addChartsheet(const QString &name = QString());
    /**
     * @brief adds new worksheet to the workbook.
     * @param name optional sheet name. If empty, new sheet will have the name
     * Sheet# where # is the sequential worksheet number.
     * @return pointer to the new worksheet on success, `nullptr` otherwise.
     */
    Worksheet *addWorksheet(const QString &name = QString());
    /**
     * @brief inserts new sheet
     * @param index the sheet index (0 to #sheetsCount()-1).
     * @param name optional sheet name. If empty, new sheet will have the name
     * Sheet# or Chart# depending on @a type.
     * @param type optional sheet type.
     * @return pointer to the new sheet on success, `nullptr` otherwise.
     */
    AbstractSheet *insertSheet(int index, const QString &name = QString(),
                               AbstractSheet::Type type = AbstractSheet::Type::Worksheet);
    /**
     * @brief renames sheet with @a index.
     * @param index the sheet index (0 to #sheetsCount()-1).
     * @param newName new sheet name.
     * @return true on success.
     * @note This is equivalent to `sheet(index)->rename(newName);`.
     */
    bool renameSheet(int index, const QString &newName);
    /**
     * @overload
     * @brief renames sheet
     * @param oldName old sheet name
     * @param newName new sheet name
     * @return true on success.
     * @note This is equivalent to `sheet(oldName)->rename(newName);`.
     */
    bool renameSheet(const QString &oldName, const QString &newName);
    /**
     * @brief deletes the sheet with @a index.
     * @param index the sheet index (0 to #sheetsCount()-1).
     * @return true on success.
     */
    bool deleteSheet(int index);
    /**
     * @brief copies sheet with @a index into a new sheet and gives the new sheet @a newName.
     *
     * This method places new sheet _at the end of sheets list_. The copy will have
     * index `sheetsCount()-1`.
     *
     * @param index the sheet index (0 to #sheetsCount()-1).
     * @param newName the name of the new copy.
     * @return true on success.
     */
    bool copySheet(int index, const QString &newName=QString());
    /**
     * @brief moves the sheet from @a srcIndex to @a dstIndex.
     * @param srcIndex the old index.
     * @param dstIndex the new index.
     * @return true on success.
     */
    bool moveSheet(int srcIndex, int dstIndex);
    /**
     * @brief moves the sheet named @a sheetName to @a dstIndex.
     * @param sheetName the sheet to move.
     * @param dstIndex the new index.
     * @return true on success.
     */
    bool moveSheet(const QString &sheetName, int dstIndex);
    /**
     * @brief returns the active (current) sheet.
     * @return pointer to the active sheet. If no sheets were added to the
     * workbook, inserts a new worksheet and returns pointer to it.
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
    /**
     * @brief sets the active (current) sheet.
     * @param index the sheet index (0 to #sheetsCount()-1).
     * @return true on success.
     */
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
     * @return true if adding defined name is successful. False if @a name or @a formula is empty or there
     * is already a definedName with @a name.
     * @note Even if you specify the sheet name in @a scope, Excel will treat @a formula without
     * the sheet name part as invalid. LibreOffice is OK with that. Example:
     * ```cpp
     * xlsx.addDefinedName("MyCol_3", "=Sheet1!$C$1:$C$10", "Sheet1"); //The only way for Excel
     * //vs
     * xlsx.addDefinedName("MyCol_3", "=$C$1:$C$10", "Sheet1"); //OK for LibreOffice, error for Excel
     * ```
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
     * @return A pointer to the defined name. May be null if there's no such
     * defined name in the workbook.
     * @warning The pointer can become invalidated if you add or remove defined names.
     */
    DefinedName *definedName(const QString &name);

    /**
     * @brief returns whether the default epoch of this workbook is set to 1904.
     *
     * - In the 1900 date system, the lower limit is January 1st, 0001 00:00:00,
     * which has a serial date-time of -693593. The upper-limit is December 31st,
     * 9999, 23:59:59.999, which has a serial date-time of 2,958,465.9999884.
     * The base date for this system is 00:00:00 on December 30th, 1899, which
     * has a serial date-time of 0.
     * - In the 1904 date system, the lower limit is January 1st, 0001, 00:00:00,
     * which has a serial date-time of -695055. The upper limit is December 31st,
     * 9999, 23:59:59.999, which has a serial date-time of 2,957,003.9999884.
     * The base date for this system is 00:00:00 on January 1st, 1904, which has
     * a serial date-time of 0.
     *
     * Excel for Windows uses a default epoch of 1900 and Excel for Mac uses an
     * epoch of 1904. However, Excel on either platform will convert automatically
     * between one system and the other. QXlsx stores dates in the 1900 format
     * by default.
     *
     * @return true if all dates are calculated using the 1904 epoch.
     *
     * If not set, false is assumed (the default epoch being 1900).
     */
    std::optional<bool> date1904() const;
    /**
     * @brief sets the default epoch for this workbook.
     *
     * - In the 1900 date system, the lower limit is January 1st, 0001 00:00:00,
     * which has a serial date-time of -693593. The upper-limit is December 31st,
     * 9999, 23:59:59.999, which has a serial date-time of 2,958,465.9999884.
     * The base date for this system is 00:00:00 on December 30th, 1899, which
     * has a serial date-time of 0.
     * - In the 1904 date system, the lower limit is January 1st, 0001, 00:00:00,
     * which has a serial date-time of -695055. The upper limit is December 31st,
     * 9999, 23:59:59.999, which has a serial date-time of 2,957,003.9999884.
     * The base date for this system is 00:00:00 on January 1st, 1904, which has
     * a serial date-time of 0.
     *
     * Excel for Windows uses a default epoch of 1900 and Excel
     * for Mac uses an epoch of 1904. However, Excel on either
     * platform will convert automatically between one system
     * and the other. QXlsx stores dates in the 1900 format
     * by default.
     *
     * @param date1904 If true, then all dates will be calculated using the 1904
     * epoch.
     *
     * If not set, false is assumed (the default epoch being 1900).
     * @note This function should be called before any date/time has been written.
     */
    void setDate1904(bool date1904);
    /**
     * @brief returns whether strings are converted to numbers (floats) where
     * possible in the Worksheet::write() method in order to avoid
     * an Excel warning about "Numbers Stored as Text".
     * @return `true` if converting is enabled.
     *
     * The default value is false.
     */
    bool isStringsToNumbersEnabled() const;

    /**
     * @brief Enables the Worksheet::write() method to convert strings
     * to numbers, where possible, using float() in order to avoid
     * an Excel warning about "Numbers Stored as Text".
     * @param enable
     *
     * The default value is false.
     */
    void setStringsToNumbersEnabled(bool enable=true);
    /**
     * @brief returns whether strings are converted to hyperlinks where
     * possible in the Worksheet::write() method.
     * @return `true` if converting is enabled.
     *
     * The default value is `true`.
     */
    bool isStringsToHyperlinksEnabled() const;
    /**
     * @brief Enables the Worksheet::write() method to convert strings
     * to hyperlinks, where possible.
     * @param enable If `true` then strings are tested for being a hyperlink and
     * are written accordingly.
     *
     * The default value is `true`.
     */
    void setStringsToHyperlinksEnabled(bool enable=true);
    /**
     * @brief returns whether strings are automatically converted to rich text if
     * they contain HTML tags.
     *
     * The default value is `false`.
     */
    bool isHtmlToRichStringEnabled() const;
    /**
     * @brief Enables the Worksheet::writeString() method to convert strings to
     * rich text if strings contain HTML tags.
     * @param enable If `true` then strings that contain HTML tags are written
     * as rich text in the Worksheet::writeString() method. If `false` then such
     * string are written as plain text.
     *
     * The default value is `false`.
     */
    void setHtmlToRichStringEnabled(bool enable = true);
    /**
     * @brief returns the default date format.
     *
     * If the default date format is not set, 'yyyy-mm-dd' is used.
     */
    QString defaultDateFormat() const;
    /**
     * @brief sets the default date format.
     * @param format a string with standard QDate placeholders (y, yy, yyyy,
     * mm, mmmm, d, dd etc.)
     *
     * If not set, 'yyyy-mm-dd' is used.
     */
    void setDefaultDateFormat(const QString &format);

    // Calculation parameters (not all)

    /**
     * @brief returns the mode the workbook uses to reference cells.
     *
     * The default value is ReferenceMode::A1.
     */
    std::optional<ReferenceMode> referenceMode() const;
    /**
     * @brief sets the mode the workbook uses to reference cells.
     * @param mode ReferenceMode enum value.
     *
     * If the reference mode is not set, ReferenceMode::A1 is assumed.
     */
    void setReferenceMode(ReferenceMode mode);
    /**
     * @brief returns the mode when the application should calculate formulas in
     * the workbook.
     *
     * The default value is CalculationMode::Auto.
     */
    std::optional<CalculationMode> calculationMode() const;
    /**
     * @brief sets the mode when the application should calculate formulas in
     * the workbook.
     * @param mode CalculationMode enum value.
     *
     * If the calculation mode is not set, CalculationMode::Auto is assumed.
     */
    void setCalculationMode(CalculationMode mode);
    /**
     * @brief sets whether the application shall perform a full recalculation
     * when the workbook is opened.
     *
     * If #calculationMode() is set to CalculationMode::Manual, then this parameter
     * is ignored.
     *
     * @param recalculate If true, then the application performs a full recalculation
     * of workbook values when the workbook is opened.
     *
     * If no value is set, false is assumed.
     */
    void setRecalculationOnLoad(bool recalculate);
    /**
     * @brief returns a boolean that indicates the precision the application uses
     * when performing calculations.
     *
     * `True` indicates the application uses the stored values of the referenced
     * cells when performing calculations.
     * `False` indicates the application uses the display values of the referenced
     * cells when performing calculations.
     *
     * Example: If two cells each contain the value 10.005 and the cells are
     * formatted to display values in currency format, the value $10.01 is displayed
     * in each cell. If you add the two cells together, the result is $20.01
     * because the application adds the stored values 10.005 and 10.005, not the
     * displayed values. You can change the precision of calculations so that the
     * application uses the displayed value instead of the stored value when it
     * recalculates formulas. So, if #fullPrecisionOnCalculation() is false, then
     * the result must be $20.02, because each cell shows $10.01, so those are
     * the values to be added.
     *
     * The default value is true.
     */
    std::optional<bool> fullPrecisionOnCalculation() const;
    /**
     * @brief sets a boolean that indicates the precision the application uses
     * when performing calculations.
     *
     * Example: If two cells each contain the value 10.005 and the cells are
     * formatted to display values in currency format, the value $10.01 is displayed
     * in each cell. If you add the two cells together, the result is $20.01
     * because the application adds the stored values 10.005 and 10.005, not the
     * displayed values. You can change the precision of calculations so that the
     * application uses the displayed value instead of the stored value when it
     * recalculates formulas. So, if #fullPrecisionOnCalculation() is false, then
     * the result must be $20.02, because each cell shows $10.01, so those are
     * the values to be added.
     *
     * @param fullPrecision `true` indicates the application uses the stored values
     * of the referenced cells when performing calculations.
     * `false` indicates the application uses the display values of the referenced
     * cells when performing calculations.
     *
     * If fullPrecision is not set, true is assumed.
     */
    void setFullPrecisionOnCalculation(bool fullPrecision);
    /**
     * @brief Clears all calculation parameters and sets the default values.
     */
    void setCalculationParametersDefaults();



    /// Protection


    /**
     * @brief returns the workbook protection parameters.
     */
    WorkbookProtection workbookProtection() const;
    /**
     * @brief creates the default workbook protection and returns
     * the reference to it.
     */
    WorkbookProtection &workbookProtection();
    /**
     * @brief sets the workbook protection parameters.
     * @param protection WorkbookProtection object. If invalid, removes the
     * workbook protection.
     */
    void setWorkbookProtection(const WorkbookProtection &protection);
    /**
     * @brief returns whether the workbook is protected.
     * @return true if any of the protection parameters were set.
     * @sa #isPasswordProtectionSet().
     */
    bool isWorkbookProtected() const;
    /**
     * @brief returns whether the workbook is protected with password.
     *
     * If password protection is set, then #isWorkbookProtected() also returns true.
     */
    bool isPasswordProtectionSet() const;
    /**
     * @brief sets the password protection to the workbook.
     * @param algorithm a string that describes the hashing algorithm used.
     * See #Protection::algorithmName for some reserved values.
     * @param hash a string that contains the hashed password in a base64 form.
     * @param salt a string that contains the salt in a base64 form.
     * @param spinCount count of iterations to compute the password hash.
     *
     * The actual hashing should be done outside this library.
     * See QCryptographicHash and QPasswordDigestor.
     */
    void setPasswordProtection(const QString &algorithm, const QString &hash,
                               const QString &salt = QString(), int spinCount = 1);
    /**
     * @brief deletes any workbook protection parameters that have been set.
     */
    void removeWorkbookProtection();


private:
    SERIALIZE_ENUM(ShowObjects, {
        {ShowObjects::All, "all"},
        {ShowObjects::None, "none"},
        {ShowObjects::Placeholders, "placehoders"}
    });
    SERIALIZE_ENUM(UpdateLinks, {
        {UpdateLinks::Always, "always"},
        {UpdateLinks::Never, "never"},
        {UpdateLinks::UserSet, "userSet"}
    });
    SERIALIZE_ENUM(CalculationMode, {
        {CalculationMode::Auto, "auto"},
        {CalculationMode::AutoNoTable, "autoNoTable"},
        {CalculationMode::Manual, "manual"}
    });
    SERIALIZE_ENUM(ReferenceMode, {
        {ReferenceMode::A1, "A1"},
        {ReferenceMode::R1C1, "R1C1"}
    });
    friend class AbstractSheet;
    friend class Worksheet;
    friend class Chartsheet;
    friend class WorksheetPrivate;
    friend class Document;
    friend class DocumentPrivate;
    friend class Cell;
    friend class FillFormat;
    friend class DrawingAnchor;

    Workbook(Workbook::CreateFlag flag);

    void saveToXmlFile(QIODevice *device) const override;
    bool loadFromXmlFile(QIODevice *device) override;

    bool addMediaFile(QSharedPointer<MediaFile> media, bool force=false);
    void removeMediaFile(QSharedPointer<MediaFile> media);
    QList<QWeakPointer<MediaFile> > mediaFiles() const;
    void addChartFile(const QSharedPointer<Chart> &chartFile);
    void removeChartFile(const QSharedPointer<Chart> &chart);
    QList<QWeakPointer<Chart> > chartFiles() const;

    SharedStrings *sharedStrings() const;
    Styles *styles();
    Theme *theme();
    QList<QImage> images();
    QList<Drawing *> drawings();
    QList<QSharedPointer<AbstractSheet> > getSheetsByType(AbstractSheet::Type type) const;
    AbstractSheet *addSheet(const QString &name, int sheetId, AbstractSheet::Type type);
};

}

#endif // XLSXWORKBOOK_H
