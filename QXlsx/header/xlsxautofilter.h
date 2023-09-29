#ifndef XLSXAUTOFILTER_H
#define XLSXAUTOFILTER_H

#include <QSharedData>
#include <QXmlStreamWriter>
#include <QDateTime>

#include "xlsxglobal.h"
#include "xlsxcellrange.h"
#include "xlsxutility_p.h"


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
    /**
     * @brief The DynamicFilterType enum specifies criteria of filtering for dynamic filters.
     */
    enum class DynamicFilterType
    {
        Invalid, /**< Common filter type not available.  */
        AboveAverage, /**< Shows values that are above average. */
        BelowAverage, /**< Shows values that are below average. */
        Tomorrow, /**< Shows tomorrow's dates. */
        Today, /**< Shows today's dates. */
        Yesterday, /**< Shows yesterday's dates. */
        NextWeek, /**< Shows next week's dates, using Sunday as the first weekday. */
        ThisWeek, /**< Shows this week's dates, using Sunday as the first weekday. */
        LastWeek, /**<  Shows last week's dates, using Sunday as the first weekday. */
        NextMonth, /**< Shows next month's dates. */
        ThisMonth, /**< Shows this month's dates. */
        LastMonth, /**< Shows last month's dates. */
        NextQuarter, /**< Shows next calendar quarter's dates. */
        ThisQuarter, /**< Shows this calendar quarter's dates. */
        LastQuarter, /**< Shows last calendar quarter's dates. */
        NextYear, /**< Shows next year's dates. */
        ThisYear, /**< Shows this year's dates. */
        LastYear, /**< Shows last year's dates.  */
        YearToDate, /**< Shows the dates between the beginning of the year and today, inclusive. */
        Q1, /**< Shows the dates that are in the 1st calendar quarter, regardless of year. */
        Q2, /**< Shows the dates that are in the 2nd calendar quarter, regardless of year. */
        Q3, /**< Shows the dates that are in the 3rd calendar quarter, regardless of year. */
        Q4, /**< Shows the dates that are in the 4th calendar quarter, regardless of year. */
        M1, /**< Shows the dates that are in January, regardless of year. */
        M2, /**< Shows the dates that are in February, regardless of year. */
        M3, /**<  Shows the dates that are in March, regardless of year. */
        M4, /**< Shows the dates that are in April, regardless of year. */
        M5, /**< Shows the dates that are in May, regardless of year. */
        M6, /**< Shows the dates that are in June, regardless of year.  */
        M7, /**<  Shows the dates that are in July, regardless of year. */
        M8, /**<  Shows the dates that are in August, regardless of year. */
        M9, /**< Shows the dates that are in September, regardless of year.  */
        M10, /**< Shows the dates that are in October, regardless of year. */
        M11, /**< Shows the dates that are in November, regardless of year. */
        M12, /**<  Shows the dates that are in December, regardless of year. */
    };

    bool operator==(const Filter &other) const;

    Type type = Type::Invalid;

    QList<QVariant> filters; //filter by values

    //custom filters
    QVariant val1;
    QVariant val2;
    std::optional<Predicate> op1;// = Predicate::Equal;
    std::optional<Predicate> op2;// = Predicate::Equal;
    std::optional<Operation> op;// = Operation::Or;

    //dynamic filters
    DynamicFilterType dynamicType = DynamicFilterType::Invalid;
    std::optional<QDateTime> maxDateTime;
    std::optional<QDateTime> dateTime;
    std::optional<double> val;
    SERIALIZE_ENUM(DynamicFilterType, {
        {DynamicFilterType::Invalid, "null"},
        {DynamicFilterType::AboveAverage, "aboveAverage"},
        {DynamicFilterType::BelowAverage, "belowAverage"},
        {DynamicFilterType::Tomorrow, "tomorrow"},
        {DynamicFilterType::Today, "today"},
        {DynamicFilterType::Yesterday, "yesterday"},
        {DynamicFilterType::NextWeek, "nextWeek"},
        {DynamicFilterType::ThisWeek, "thisWeek"},
        {DynamicFilterType::LastWeek, "lastWeek"},
        {DynamicFilterType::NextMonth, "nextMonth"},
        {DynamicFilterType::ThisMonth, "thisMonth"},
        {DynamicFilterType::LastMonth, "lastMonth"},
        {DynamicFilterType::NextQuarter, "nextQuarter"},
        {DynamicFilterType::ThisQuarter, "thisQuarter"},
        {DynamicFilterType::LastQuarter, "lastQuarter"},
        {DynamicFilterType::NextYear, "nextYear"},
        {DynamicFilterType::ThisYear, "thisYear"},
        {DynamicFilterType::LastYear, "lastYear"},
        {DynamicFilterType::YearToDate, "yearToDate"},
        {DynamicFilterType::Q1, "q1"},
        {DynamicFilterType::Q2, "q2"},
        {DynamicFilterType::Q3, "q3"},
        {DynamicFilterType::Q4, "q4"},
        {DynamicFilterType::M1, "m1"},
        {DynamicFilterType::M2, "m2"},
        {DynamicFilterType::M3, "m3"},
        {DynamicFilterType::M4, "m4"},
        {DynamicFilterType::M5, "m5"},
        {DynamicFilterType::M6, "m6"},
        {DynamicFilterType::M7, "m7"},
        {DynamicFilterType::M8, "m8"},
        {DynamicFilterType::M9, "m9"},
        {DynamicFilterType::M10, "m10"},
        {DynamicFilterType::M11, "m11"},
        {DynamicFilterType::M12, "m12"},
    });
};



class AutoFilterPrivate;
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

    bool filterButtonHidden(int column) const; //default = false
    void setFilterButtonHidden(int column, bool hidden);
    void setFilterButtonHidden(bool hidden);

    bool showFilterOptions(int column) const; //default = true
    void setShowFilterOptions(int column, bool show);
    void setShowFilterOptions(bool show);


    void setFilterByTopN(int column, double value, double filterBy, bool usePercents = false); //0-based index in range
    void setFilterByBottomN(int column, double value, double filterBy, bool usePercents = false); //0-based index in range

    /**
     * @brief sets @a column filtering by @a values.
     * @param column zero-based column index in the autofilter #range(). //TODO: consider changing to real column index
     * @param values Values to filter rows by.
     * @note The previous filter in this column will be deleted.
     */
    void setFilterByValues(int column, const QVariantList &values);
    /**
     * @brief returns values that were set to filter @a column.
     *
     * Returns a list of values that autofilter use to hide rows not containing these values.
     *
     * @param column zero-based index of a column in the autofilter #range().
     * @return non-empty list if #filterType() for @a column is Filter::Type::Values and
     * filter values were set, empty list otherwise.
     */
    QVariantList filterValues(int column) const;
    /**
     * @brief sets filtering for @a column based on a predicate or pair of predicates.
     *
     * Suppose we want to clamp values in the first column leaving only values from range [100, 1000].
     * Then the following code will do the job:
     * @code setCustomPredicate(0, Filter::Predicate::GreaterThanOrEqual, 100, Filter::Operation::And,
     * Filter::Predicate::LessThanOrEqual, 1000);
     * @endcode
     *
     * @param column zero-based index of a column in the autofilter #range().
     * @param val1 the first value to be compared with cell values according to @a op1.
     * @param op1 the predicate for @a val1.
     * @param val2 the second, optional, value to be compared with cell values according to @a op2.
     * @param op2 the predicate for @a val2.
     * @param op the predicate to connect @a op1 and @a op2.
     * @note The previous filter in this column will be deleted.
     */
    void setPredicate(int column,
                      Filter::Predicate op1, const QVariant &val1,
                      Filter::Operation op = Filter::Operation::Or,
                      Filter::Predicate op2 = Filter::Predicate::Equal, const QVariant &val2 = QVariant());
    /**
     * @brief sets filtering for dynamically changing data.
     *
     * A dynamic filter returns a result set which might vary due to a change in the
     * data itself or a change in the date on which the filter is being applied.

     * @param column zero-based index of a column in the autofilter #range().
     * @param type one of the following Filter::DynamicFilterType enums.
     */
    void setDynamicFilter(int column, Filter::DynamicFilterType type);
    /**
     * @brief returns dynamic filter type of the filter that was set to @a column.
     * @param column zero-based index of a column in the autofilter #range().
     * @return dynamic filter type. Can be Filter::DynamicFilterType::Invalid if
     * no filter was set to @a column.
     */
    Filter::DynamicFilterType dynamicFilterType(int column) const;

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