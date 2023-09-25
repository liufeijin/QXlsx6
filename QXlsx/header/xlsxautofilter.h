#ifndef XLSXAUTOFILTER_H
#define XLSXAUTOFILTER_H

#include <QSharedData>
#include <QXmlStreamWriter>

#include "xlsxglobal.h"
#include "xlsxcellrange.h"



namespace QXlsx {

// Represents filter in the column
struct Filter
{
    /**
     * @brief The Type enum specifies the type of filter.
     */
    enum class Type
    {
        Invalid, /**< Invalid filter */
        Top10, /**< Top 10 values */
        Custom, /**< Filter based on predicate(s) */
        Dynamic, /**< Dynamic filter */
        Color,
        Icon,
        Values /**< Filter based on a collection of acceptable values */
    };
    enum class Predicate
    {
        Equal,
        LessThan,
        LessThanOrEqual,
        NotEqual,
        GreaterThan,
        GreaterThanOrEqual
    };
    enum class Operation
    {
        And,
        Or
    };

    bool operator==(const Filter &other) const;

    Type type = Type::Invalid;

    QList<QVariant> filters; //filter by values

    //custom filters
    QVariant val1;
    QVariant val2;
    Predicate op1 = Predicate::Equal;
    Predicate op2 = Predicate::Equal;
    Operation op = Operation::Or;
};

/**
 * @brief The AutoFilter class represents the autofilter parameters for worksheet data.
 *
 * If set, autofilter temporarily hides rows based on a filter criteria, which is
 * applied column by column to a table of data in the worksheet.
 *
 * Autofilter is applied to a range with #setRange() and #range().
 * You can set filters to columns within #range() with #setFilter().
 * To remove filter use #removeFilter(). To get current filter use #filter().
 */

class AutoFilterPrivate;

class QXLSX_EXPORT AutoFilter
{
public:
    /**
     * @brief creates invalid autofilter.
     */
    AutoFilter();
    /**
     * @brief creates default autofilter that applies to @a range.
     * @param range A range to apply default autofilter. The range defines columns
     * that will get the autofilter options and rows that will be subjected to autofiltering.
     * If \a range is invalid, then autofilter will be applied to rows in the first column (column "A"),
     * and all filters except for the first one will be ignored.
     */
    explicit AutoFilter(const CellRange &range);
    AutoFilter(const AutoFilter &other);
    AutoFilter &operator=(const AutoFilter &other);
    ~AutoFilter();

    /**
     * @brief returns whether the autofilter is valid.
     * @return
     */
    bool isValid() const;

    /**
     * @brief returns the cell range the autofilter is applied to.
     * @return the cell range the filter is applied to. Can be invalid, this means that
     * autofilter is applied to the first column starting from the first row,
     * and all filters except for the first one will be ignored.
     */
    CellRange range() const;
    /**
     * @brief sets the cell range the autofilter is applied to.
     * @param range the cell range the filter is applied to. Can be invalid, this means that
     * autofilter is applied to the first column starting from the first row.
     * @note Column headers should also be included into the range, as they will get
     * autofilter option buttons.
     */
    void setRange(const CellRange &range);

    void setFilterByTopN(int column, double value, double filterBy, bool usePercents = false); //0-based index in range
    void setFilterByBottomN(int column, double value, double filterBy, bool usePercents = false); //0-based index in range

    /**
     * @brief sets @a column filtering by @a values.
     * @param column zero-based column index in the autofilter #range(). //TODO: consider changing to real column index
     * @param values Values to filter rows by.
     */
    void setFilterByValues(int column, const QVariantList &values);
    /**
     * @brief returns values that were set to filter @a column.
     *
     * Returns a list of values that autofilter use to hide rows not containing these values.
     *
     * @param column zero-based index of a column in the autofilter #range().
     * @return non-empty list if #filterType() for @column is Filter::Type::Values and
     * filter values were set, empty list otherwise.
     */
    QVariantList filterValues(int column) const;
    /**
     * @brief sets filtering for @a column based on a predicate or pair of predicates.
     *
     * Suppose we want to clamp values in the first column leaving only values from range [100, 1000].
     * Then the following code will do the job:
     * @code setCustomPredicate(0, 100, Filter::Predicate::GreaterThanOrEqual, 1000, Filter::Predicate::LessThanOrEqual,
     * Filter::Operation::And);
     * @endcode
     *
     * @param column zero-based index of a column in the autofilter #range().
     * @param val1 the first value to be compared with cell values according to @a op1.
     * @param op1 the predicate for @a val1.
     * @param val2 the second, optional, value to be compared with cell values according to @a op2.
     * @param op2 the predicate for @a val2.
     * @param op the predicate to connect @a op1 and @a op2.
     *
     */
    void setCustomPredicate(int column, const QVariant &val1, Filter::Predicate op1,
                            const QVariant &val2 = QVariant(), Filter::Predicate op2 = Filter::Predicate::Equal,
                            Filter::Operation op = Filter::Operation::Or);

    /**
     * @brief returns filter type of the filter that was set to @a column.
     * @param column zero-based index of a column in the autofilter #range().
     * @return Filter type. Can be Filter::Type::Invalid if no filter was set to @a column.
     */
    Filter::Type filterType(int column) const;
    /**
     * @brief removes filter from the @a column.
     * @param column zero-based index of a column in the autofilter #range().
     */
    void removeFilter(int column);
    /**
     * @brief returns filter for @a column.
     * @param column zero-based index of a column in the autofilter #range().
     * @return A copy of the column filter.
     */
    Filter filter(int column) const;
    /**
     * @brief returns filter for @a column.
     * @param column zero-based index of a column in the autofilter #range().
     * @return Reference to the column filter.
     * @note If no filter were set, inserts an invalid filter into @a column.
     * @note This is a low-level method. It is recommended to use hi-level methods
     * such as #setFilterByTopN(), #setFilterByBottomN(), #setFilterByValues(), #removeFilter()
     * etc.
     */
    Filter &filter(int column);


    void write(QXmlStreamWriter &writer, const QString &name) const;
    void read(QXmlStreamReader &reader);

    bool operator==(const AutoFilter &other) const;
    bool operator!=(const AutoFilter &other) const;
    operator QVariant() const;


private:
    QSharedDataPointer<AutoFilterPrivate> d;
#ifndef QT_NO_DEBUG_STREAM
    friend QDebug operator<<(QDebug, const AutoFilter &f);
#endif
};

#ifndef QT_NO_DEBUG_STREAM
QDebug operator<<(QDebug, const AutoFilter &f);
#endif

}

Q_DECLARE_METATYPE(QXlsx::AutoFilter)

#endif // XLSXAUTOFILTER_H
