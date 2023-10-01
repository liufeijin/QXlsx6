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
/*!
 * \brief Data validation for single cell or a range
 *
 * The data validation can be applied to a single cell or a range of cells.
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
     * @brief The Error enum defines the type of error dialog that
     *  is displayed.
     */
    enum class Error
    {
        Stop,
        Warning,
        Information
    };

    /**
     * @brief creates an invalid DataValidation
     */
    DataValidation();
    /*!
    * creates a DataValidation object with the given \a type, \a op, \a formula1
    * \a formula2, and \a allowBlank.
    */
    DataValidation(Type type, Operator op=Operator::Between, const QString &formula1=QString()
            , const QString &formula2=QString(), bool allowBlank=false);
    DataValidation(const DataValidation &other);
    ~DataValidation();

    Type type() const;
    Operator validationOperator() const;
    Error errorStyle() const;
    QString formula1() const;
    QString formula2() const;
    bool allowBlank() const;
    QString errorMessage() const;
    QString errorMessageTitle() const;
    QString promptMessage() const;
    QString promptMessageTitle() const;
    bool isPromptMessageVisible() const;
    bool isErrorMessageVisible() const;
    QList<CellRange> ranges() const;

    void setType(Type type);
    void setValidationOperator(Operator op);
    void setErrorStyle(Error es);
    void setFormula1(const QString &formula);
    void setFormula2(const QString &formula);
    void setErrorMessage(const QString &error, const QString &title=QString());
    void setPromptMessage(const QString &prompt, const QString &title=QString());
    void setAllowBlank(bool enable);
    void setPromptMessageVisible(bool visible);
    void setErrorMessageVisible(bool visible);

    void addCell(const CellReference &cell);
    void addCell(int row, int col);
    void addRange(int firstRow, int firstCol, int lastRow, int lastCol);
    void addRange(const CellRange &range);

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
    QSharedDataPointer<DataValidationPrivate> d;
};

}

#endif // QXLSX_XLSXDATAVALIDATION_H
