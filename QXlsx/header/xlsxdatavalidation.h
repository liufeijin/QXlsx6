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
 * relational operator (e.g., greater than, less than, equal to, etc). See
 * #setFormula1(), #setFormula2(), #setValidationOperator().
 *
 * Data validation applies to a single cell, to a range of cells or to a list of
 * ranges. Use #addCell(), #addRange(), #removeCell(), #removeRange(), #ranges()
 * to manipulate data validation ranges.
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
        Custom /**< restricts the cell based on an external Excel formula that returns a true/false value. */
    };
    /**
     * @brief The Operator enum defines the criteria by which the data in the
     *  cell is validated
     */
    enum class Operator
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
    * creates a DataValidation object with the given @a type, @a op, @a formula1
    * @a formula2, and @a allowBlank.
    */
    DataValidation(Type type, Operator op=Operator::Between, const QString &formula1=QString()
            , const QString &formula2=QString(), bool allowBlank=false);
    DataValidation(const DataValidation &other);
    ~DataValidation();

    /**
     * @brief returns the type of the validation.
     *
     * Type of the validation specifies the type of data that you wish to validate.
     * @return Type enum value.
     */
    Type type() const;
    /**
     * @brief sets the type of the validation.
     *
     * Type of the validation specifies the type of data that you wish to validate.
     * @param type Type enum value.
     */
    void setType(Type type);

    Operator validationOperator() const;
    Error errorStyle() const;
    QString formula1() const;
    QString formula2() const;
    /**
     * @brief returns whether the data validation allows the use of empty or blank
     * entries. true means empty entries are OK and do not violate the validation constraints.
     * @return If true, blank entries are allowed.
     *
     * The default value is false.
     */
    bool allowBlank() const;
    /**
     * @brief sets whether the data validation allows the use of empty or blank
     * entries. true means empty entries are OK and do not violate the validation constraints.
     * @param enable If true, then blank entries are allowed.
     *
     * If not set, false is assumed.
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
     * @return If true, then prompt message is shown.
     *
     * The default value is false.
     * @sa #setPromptMessageVivible(), #setPromptMessage().
     */
    bool isPromptMessageVisible() const;
    /**
     * @brief returns whether to display the error message.
     * @return If true, then error message is shown.
     *
     * The default value is false.
     * @sa #setErrorMessageVivible(), #setErrorMessage().
     */
    bool isErrorMessageVisible() const;
    /**
     * @brief sets whether to display the input prompt message.
     * @param visible If true, then prompt message is shown.
     *
     * If not set, false is assumed.
     * @sa #isPromptMessageVivible(), #setPromptMessage().
     */
    void setPromptMessageVisible(bool visible);
    /**
     * @brief sets whether to display the validation error message.
     * @param visible If true, then error message is shown.
     *
     * If not set, false is assumed.
     * @sa #isErrorMessageVivible(), #setErrorMessage(), #setErrorStyle().
     */
    void setErrorMessageVisible(bool visible);
    void setErrorStyle(Error es);

    void setValidationOperator(Operator op);
    void setFormula1(const QString &formula);
    void setFormula2(const QString &formula);
    /**
     * @brief sets text of the validation error with the optional UI element title.
     * @param error the text of the error message.
     * @param title the title of the error box.
     * @note The error message will not be visible if #isErrorMessageVisible() returns false.
     * @sa #setErrorMessageVisible().
     */
    void setErrorMessage(const QString &error, const QString &title=QString());
    /**
     * @brief sets the text of the prompt message with the optional UI element title.
     * @param prompt the text of the input message.
     * @param title the title of the input message.
     * @note The prompt message will not be visible if #isPromptMessageVisible() returns false.
     * @sa setPromptMessageVisible().
     */
    void setPromptMessage(const QString &prompt, const QString &title=QString());

    /**
     * @brief returns whether to display a dropdown combo box for a list-type
     * data validation.
     * @return If true, then a dropdown is shown.
     *
     * The default value is false.
     */
    bool isDropDownVisible() const;
    /**
     * @brief sets whether to display a dropdown combo box for a list-type
     * data validation.
     * @param visible If true, then a dropdown is shown.
     *
     * If not set, false is assumed.
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
    ImeMode imeMode() const;
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
     * @return true if @a cell is valid and has previously not been validated,
     * false otherwise.
     */
    bool addCell(const CellReference &cell);
    /**
     * @overload
     * @brief adds cell with (1-based) row and (1-based) column to the list of
     * validated ranges.
     * @param row 1-based cell row number.
     * @param col 1-based cell column number.
     * @return true if a cell is valid and has previously not been validated,
     * false otherwise.
     */
    bool addCell(int row, int column);
    /**
     * @brief adds @a range to the list of validated ranges.
     * @param range valid CellRange.
     * true if @a range is valid and has previously not been validated,
     * false otherwise.
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
     * @return true if a range is valid and has previously not been validated,
     * false otherwise.
     */
    bool addRange(int firstRow, int firstColumn, int lastRow, int lastColumn);
    /**
     * @brief removes @a range from the list of validated ranges.
     * @param range valid CellRange.
     * @return true if @a range was found and removed, false otherwise.
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
     * @return true if a range was found and removed, false otherwise.
     */
    bool removeRange(int firstRow, int firstColumn, int lastRow, int lastColumn);
    /**
     * @brief returns the list of validated ranges.
     * @return the list of validated ranges.
     */
    QList<CellRange> ranges() const;


    DataValidation &operator=(const DataValidation &other);
    bool operator==(const DataValidation &other) const;
    bool operator!=(const DataValidation &other) const;
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

    SERIALIZE_ENUM(Operator,
                   {{Operator::Between, QStringLiteral("between")},
                    {Operator::NotBetween, QStringLiteral("notBetween")},
                    {Operator::Equal, QStringLiteral("equal")},
                    {Operator::NotEqual, QStringLiteral("notEqual")},
                    {Operator::LessThan, QStringLiteral("lessThan")},
                    {Operator::LessThanOrEqual, QStringLiteral("lessThanOrEqual")},
                    {Operator::GreaterThan, QStringLiteral("greaterThan")},
                    {Operator::GreaterThanOrEqual, QStringLiteral("greaterThanOrEqual")}});

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

#endif // QXLSX_XLSXDATAVALIDATION_H
