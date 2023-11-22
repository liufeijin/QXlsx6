// xlsxvalidation.h

#ifndef QXLSX_XLSXDATAVALIDATION_H
#define QXLSX_XLSXDATAVALIDATION_H

#include <QtGlobal>
#include <QSharedDataPointer>
#include <QString>
#include <QList>
#include <QXmlStreamReader>
#include <QXmlStreamWriter>

#include "xlsxglobal.h"
#include "xlsxutility_p.h"

class QXmlStreamReader;
class QXmlStreamWriter;

namespace QXlsx {

class Worksheet;
class CellRange;
class CellReference;

class DataValidationPrivate;
/**
 * @brief The DataValidation class represents constraints on the data that can
 * be entered into a cell.
 *
 * Various data types can be selected. See #setType(), DataValidation::Type.
 *
 * Data validation allows to define two formulas which can be combined with the
 * predicate (e.g., greater than, less than, equal to, etc). See
 * #setFormula1(), #setFormula2(), #setPredicate().
 *
 * Data validation applies to a single cell, to a range of cells or to a list of
 * ranges. Use #addCell(), #addRange(), #removeRange(), #ranges() to manipulate
 * data validation ranges.
 *
 * Additional UI can be provided to help the user select values (e.g., a dropdown
 * control on the cell or hover text when the cell is active), and to help the
 * user understand why a particular entry was disallowed (e.g., alerts and messages).
 * See #setDropDownVisible().
 *
 * An input message can be specified to help the user know what kind of value is
 * expected, and a warning message (and warning type) can be specified to alert
 * the user when they've entered data which is not permitted based on the data
 * validations specified in the worksheet. See #setPromptMessage(), #setPromptMessageVisible(),
 * #setErrorMessage(), #setErrorMessageVisible(), #setErrorStyle().
 *
 * The class is _implicitly shareable_: the deep copy occurs only in the non-const methods.
 */
class QXLSX_EXPORT DataValidation
{
public:
    /**
     * @brief The Type enum defines the type of data that you wish to validate.
     */
    enum class Type
    {
        None, /**< the type of data is unrestricted. This is the same as not applying a data validation. */
        Whole, /**< restricts the cell to integer values. */
        Decimal, /**< restricts the cell to decimal values. */
        List, /**< restricts the cell to a set of user specified values. */
        Date, /**< restricts the cell to date values. */
        Time, /**< restricts the cell to time values. */
        TextLength, /**< restricts the cell data based on an integer string length. */
        Custom /**< restricts the cell based on an external Excel formula that returns a `true/false` value. */
    };
    /**
     * @brief The Predicate enum defines the criteria by which the data in the
     *  cell is validated
     */
    enum class Predicate
    {
        Between,
        NotBetween,
        Equal,
        NotEqual,
        LessThan,
        LessThanOrEqual,
        GreaterThan,
        GreaterThanOrEqual
    };
    /**
     * @brief The Error enum defines the type of error dialog that is displayed.
     */
    enum class Error
    {
        Stop,
        Warning,
        Information
    };
    /**
     * @brief The ImeMode enum specifies the IME (input method editor) mode
     * enforced by this data validation.
     */
    enum class ImeMode
    {
        NoControl,
        Off,
        On,
        Disabled,
        Hiragana,
        FullKatakana,
        HalfKatakana,
        FullAlpha,
        HalfAlpha,
        FullHangul,
        HalfHangul,
    };

    /**
     * @brief creates an invalid DataValidation
     */
    DataValidation();
    /**
    * creates a DataValidation object with the given @a type, @a predicate, @a formula1
    * @a formula2.
    */
    DataValidation(Type type,
                   const QString &formula1=QString(),
                   std::optional<Predicate> predicate = std::nullopt,
                   const QString &formula2=QString());
    /**
     * @brief creates a DataValidation object with type Type::List and a range
     * of @a allowableValues.
     * @param allowableValues the range that contains allowable values.
     * @note This constructor creates formula1 from the @a allowableValues
     * converted into a fixed range. F.e. for a range of CellRange(1,1,4,10)
     * this constructor assigns "$A$1:$J$4" to formula1.
     * There is a possibility to create the dynamic list of values by setting formula1
     * directly. Try `setFormula1("A1:J4")` to see the effects. See also DataValidation
     * example.
     */
    DataValidation(const CellRange &allowableValues);
    /**
     * @brief creates a DataValidation object with type Type::Time and two QTime values
     * combined with @a predicate.
     * @param time1 First time value.
     * @param predicate Predicate enum
     * @param time2 Second time value.
     */
    DataValidation(const QTime &time1,
                   std::optional<Predicate> predicate = std::nullopt,
                   const QTime &time2 = QTime());
    /**
     * @brief creates a DataValidation object with type Type::Date and two QDate values
     * combined with @a predicate.
     * @param date1 First date value.
     * @param predicate Predicate enum
     * @param date2 Second date value.
     */
    DataValidation(const QDate &date1,
                   std::optional<Predicate> predicate = std::nullopt,
                   const QDate &date2 = QDate());
    /**
     * @brief creates a DataValidation object with type Type::TextLength and two
     * text lengths combined with @a predicate.
     * @param len1 First text length
     * @param predicate Predicate enum
     * @param len2 Second text length
     */
    DataValidation(int len1,
                   std::optional<Predicate> predicate = std::nullopt,
                   std::optional<int> len2 = std::nullopt);

    DataValidation(const DataValidation &other);
    ~DataValidation();

    /**
     * @brief returns the type of the validation.
     *
     * Type of the validation specifies the type of data that you wish to validate.
     * @return Type enum value if type was set, `nullopt` otherwise.
     *
     * The default value is Type::None, which means no validation constraints.
     */
    std::optional<Type> type() const;
    /**
     * @brief sets the type of the validation.
     *
     * Type of the validation specifies the type of data that you wish to validate.
     * @param type Type enum value.
     * If not set, Type::None is assumed, which means no validation constraints.
     */
    void setType(Type type);
    /**
     * @brief returns predicate that defines the criteria by which the data in the
     * cell is validated.
     * @return Predicate enum value if predicate was set, `nullopt` otherwise.
     * The default value is Predicate::Between.
     */
    std::optional<Predicate> predicate() const;
    /**
     * @brief sets predicate that defines the criteria by which the data in the
     * cell is validated.
     * @param predicate Predicate enum value.
     *
     * If not set, Predicate::Between is assumed.
     */
    void setPredicate(Predicate predicate);
    /**
     * @brief return the type of error message when cell data was disallowed.
     * @return Error enum value if errorStyle was set, `nullopt` otherwise.
     *
     * The default value is Error::Stop.
     */
    std::optional<Error> errorStyle() const;
    /**
     * @brief sets the type of error message when cell data was disallowed.
     * @param errorStyle Error enum value.
     *
     * If not set, Error::Stop is assumed.
     */
    void setErrorStyle(Error errorStyle);
    /**
     * @brief returns the first formula to be used as a validation criterion.
     * @return Non-empty string if formula1 was set, empty string otherwise.
     *
     * By default formula1 is empty.
     */
    QString formula1() const;
    /**
     * @brief returns the second formula to be used as a validation criterion.
     * @return Non-empty string if formula2 was set, empty string otherwise.
     *
     * By default formula2 is empty.
     */
    QString formula2() const;
    /**
     * @brief sets the first formula to be used as a validation criterion.
     * @param formula
     *
     * If not set, empty string is assumed.
     */
    void setFormula1(const QString &formula);
    /**
     * @brief sets the second formula to be used as a validation criterion.
     * @param formula
     *
     * If not set, empty string is assumed.
     */
    void setFormula2(const QString &formula);
    /**
     * @brief returns whether the data validation allows the use of empty or blank
     * entries. `true` means empty entries are OK and do not violate the validation constraints.
     * @return if `true`, blank entries are allowed.
     *
     * The default value is `false`.
     */
    std::optional<bool> allowBlank() const;
    /**
     * @brief sets whether the data validation allows the use of empty or blank
     * entries. `true` means empty entries are OK and do not violate the validation constraints.
     * @param enable if `true`, then blank entries are allowed.
     *
     * If not set, `false` is assumed.
     */
    void setAllowBlank(bool enable);
    /**
     * @brief returns text that is shown if data entry fails validation.
     * @return Non-empty string if error message was specified.
     *
     * The default value is empty string.
     */
    QString errorMessage() const;
    /**
     * @brief returns text that is shown in the title of an UI element if
     * data entry fails validation.
     * @return Non-empty string if error message title was specified.
     *
     * The default value is empty string.
     */
    QString errorTitle() const;
    /**
     * @brief returns text of the validation prompt.
     * @return  Non-empty string if prompt message was specified.
     *
     * The default value is empty string.
     */
    QString promptMessage() const;
    /**
     * @brief returns text that is shown in the title of an UI element of the
     * validation prompt.
     * @return Non-empty string if prompt title message was specified.
     */
    QString promptTitle() const;
    /**
     * @brief returns whether to display the input prompt message.
     * @return if `true`, then prompt message is shown.
     *
     * The default value is `false`.
     * @sa #setPromptMessageVisible(), #setPromptMessage().
     */
    std::optional<bool> isPromptMessageVisible() const;
    /**
     * @brief returns whether to display the error message.
     * @return if `true`, then error message is shown.
     *
     * The default value is `false`.
     * @sa #setErrorMessageVisible(), #setErrorMessage().
     */
    std::optional<bool> isErrorMessageVisible() const;
    /**
     * @brief sets whether to display the input prompt message.
     * @param visible if `true`, then prompt message is shown.
     *
     * If not set, `false` is assumed.
     * @sa #isPromptMessageVisible(), #setPromptMessage().
     */
    void setPromptMessageVisible(bool visible);
    /**
     * @brief sets whether to display the validation error message.
     * @param visible if `true`, then error message is shown.
     *
     * If not set, `false` is assumed.
     * @sa #isErrorMessageVisible(), #setErrorMessage(), #setErrorStyle().
     */
    void setErrorMessageVisible(bool visible);
    /**
     * @brief sets text of the validation error with the optional UI element title.
     * @param error the text of the error message.
     * @param title the title of the error box.
     * @note The error message will not be visible if #isErrorMessageVisible() returns `false`.
     * @sa #setErrorMessageVisible().
     */
    void setErrorMessage(const QString &error, const QString &title=QString());
    /**
     * @brief sets the text of the prompt message with the optional UI element title.
     * @param prompt the text of the input message.
     * @param title the title of the input message.
     * @note The prompt message will not be visible if #isPromptMessageVisible() returns `false`.
     * @sa setPromptMessageVisible().
     */
    void setPromptMessage(const QString &prompt, const QString &title=QString());

    /**
     * @brief returns whether to display a dropdown combo box for a list-type
     * data validation.
     * @return if `true`, then a dropdown is shown.
     *
     * The default value is `false`.
     */
    std::optional<bool> isDropDownVisible() const;
    /**
     * @brief sets whether to display a dropdown combo box for a list-type
     * data validation.
     * @param visible if `true`, then a dropdown is shown.
     *
     * If not set, `false` is assumed.
     */
    void setDropDownVisible(bool visible);

    /**
     * @brief returns the IME (input method editor) mode enforced by this data validation.
     *
     * When imeMode is set, the input for the cell can be restricted to specific
     * sets of characters, as specified by the value of ImeMode.
     * @return One of ImeMode enumerations.
     *
     * The default value is ImeMode::NoControl.
     */
    std::optional<ImeMode> imeMode() const;
    /**
     * @brief sets the IME (input method editor) mode enforced by this data validation.
     *
     * When imeMode is set, the input for the cell can be restricted to specific
     * sets of characters, as specified by the value of ImeMode.
     * @param mode One of ImeMode enumerations.
     *
     * If not set, ImeMode::NoControl is assumed.
     */
    void setImeMode(ImeMode mode);

    /**
     * @brief adds @a cell to the list of validated ranges.
     * @param cell valid CellReference.
     * @return `true` if @a cell is valid and has previously not been validated,
     * `false` otherwise.
     */
    bool addCell(const CellReference &cell);
    /**
     * @overload
     * @brief adds cell with (1-based) @a row and (1-based) @a column to the
     * list of validated ranges.
     * @param row 1-based cell row number.
     * @param column 1-based cell column number.
     * @return `true` if a cell is valid and has previously not been validated,
     * `false` otherwise.
     */
    bool addCell(int row, int column);
    /**
     * @brief adds @a range to the list of validated ranges.
     * @param range valid CellRange.
     * @return `true` if @a range is valid and has previously not been validated,
     * `false` otherwise.
     */
    bool addRange(const CellRange &range);
    /**
     * @overload
     * @brief adds a range specified with @a firstRow, @a firstColumn, @a lastRow,
     * @a lastColumn to the list of validated ranges.
     * @param firstRow 1-based index of the first row of the range.
     * @param firstColumn 1-based index of the first column of the range.
     * @param lastRow 1-based index of the last row of the range.
     * @param lastColumn 1-based index of the last column of the range.
     * @return `true` if a range is valid and has previously not been validated,
     * `false` otherwise.
     */
    bool addRange(int firstRow, int firstColumn, int lastRow, int lastColumn);
    /**
     * @brief removes @a range from the list of validated ranges.
     * @param range valid CellRange.
     * @return `true` if @a range was found and removed, `false` otherwise.
     */
    bool removeRange(const CellRange &range);
    /**
     * @overload
     * @brief removes a range specified with @a firstRow, @a firstColumn, @a lastRow,
     * @a lastColumn from the list of validated ranges.
     * @param firstRow 1-based index of the first row of the range.
     * @param firstColumn 1-based index of the first column of the range.
     * @param lastRow 1-based index of the last row of the range.
     * @param lastColumn 1-based index of the last column of the range.
     * @return `true` if a range was found and removed, `false` otherwise.
     */
    bool removeRange(int firstRow, int firstColumn, int lastRow, int lastColumn);
    /**
     * @brief returns the list of validated ranges.
     * @return the list of validated ranges. If the list is empty, the validation
     * is invalid.
     */
    QList<CellRange> ranges() const;


    DataValidation &operator=(const DataValidation &other);
    bool operator==(const DataValidation &other) const;
    bool operator!=(const DataValidation &other) const;
    operator QVariant() const;
    bool isValid() const;

    void write(QXmlStreamWriter &writer) const;
    void read(QXmlStreamReader &reader);
private:
    SERIALIZE_ENUM(Type,
                   {{Type::None, "none"},
                    {Type::Whole, "whole"},
                    {Type::Decimal, "decimal"},
                    {Type::List, "list"},
                    {Type::Date, "date"},
                    {Type::Time, "time"},
                    {Type::TextLength, "textLength"},
                    {Type::Custom, "custom"}});

    SERIALIZE_ENUM(Predicate,
                   {{Predicate::Between, QStringLiteral("between")},
                    {Predicate::NotBetween, QStringLiteral("notBetween")},
                    {Predicate::Equal, QStringLiteral("equal")},
                    {Predicate::NotEqual, QStringLiteral("notEqual")},
                    {Predicate::LessThan, QStringLiteral("lessThan")},
                    {Predicate::LessThanOrEqual, QStringLiteral("lessThanOrEqual")},
                    {Predicate::GreaterThan, QStringLiteral("greaterThan")},
                    {Predicate::GreaterThanOrEqual, QStringLiteral("greaterThanOrEqual")}});

    SERIALIZE_ENUM(Error,
                   {{Error::Stop, QStringLiteral("stop")},
                    {Error::Warning, QStringLiteral("warning")},
                    {Error::Information, QStringLiteral("information")}});
    SERIALIZE_ENUM(ImeMode, {
        {ImeMode::NoControl, "noControl"},
        {ImeMode::Off, "off"},
        {ImeMode::On, "on"},
        {ImeMode::Disabled, "disabled"},
        {ImeMode::Hiragana, "hiragana"},
        {ImeMode::FullKatakana, "fullKatakana"},
        {ImeMode::HalfKatakana, "halfKatakana"},
        {ImeMode::FullAlpha, "fullAlpha"},
        {ImeMode::HalfAlpha, "halfAlpha"},
        {ImeMode::FullHangul, "fullHangul"},
        {ImeMode::HalfHangul, "halfHangul"},
        });
    QSharedDataPointer<DataValidationPrivate> d;
};

}

Q_DECLARE_METATYPE(QXlsx::DataValidation);
Q_DECLARE_TYPEINFO(QXlsx::DataValidation, Q_MOVABLE_TYPE);

#endif // QXLSX_XLSXDATAVALIDATION_H
