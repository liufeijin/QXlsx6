#ifndef XLSXAUTOFILTER_H
#define XLSXAUTOFILTER_H

#include <QSharedData>
#include <QXmlStreamWriter>
#include <QDateTime>

#include "xlsxglobal.h"
#include "xlsxcellrange.h"
#include "xlsxutility_p.h"
#include "xlsxmain.h"

namespace QXlsx {

/**
 * @brief The Filter struct represents the filter parameters in a worksheet column.
 */
struct QXLSX_EXPORT Filter
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
        Color, /**< Filter based on cell color */
        Icon, /**< Filter based on icon stored in cell. */
        Values /**< Filter based on a collection of acceptable values. */
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
    bool operator!=(const Filter &other) const;

    Type type = Type::Invalid;

    //filter by values
    QList<QVariant> filters;

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

    //top 10 filters
    std::optional<bool> top; //default = true;
    std::optional<bool> usePercents; //default = false;
    std::optional<double> top10val;
    std::optional<double> top10filterVal;

};

/**
 * @brief The SortCondition struct specifies the sort condition parameters in
 * a SortState object.
 */
struct QXLSX_EXPORT SortCondition
{
    /**
     * @brief The SortBy enum specifies the type of sorting.
     */
    enum class SortBy
    {
        Value, /**< Sort by the cell value (default) */
        CellColor, /**< Sort by the cell color */
        FontColor, /**< Sort by the cell font (foreground) color */
        Icon /**< Sort by the cell icon */
    };
    std::optional<Qt::SortOrder> descending;// default=Qt::SortAscending
    std::optional<SortBy> sortBy;// default="value"
    CellRange range; //required
    QString customList;
    std::optional<int> dxfId;
    bool operator==(const SortCondition &other) const;
    //TODO: std::optional<IconSet> iconSet;// default="3Arrows"
    //TODO: std::optional<int> iconId;

    SERIALIZE_ENUM(SortBy, {
        {SortBy::CellColor, "cellColor"},
        {SortBy::FontColor, "fontColor"},
        {SortBy::Icon, "icon"},
        {SortBy::Value, "value"}
    });
};

/**
 * @brief The SortState class represents the sorting parameters in the autofilter.
 *
 * Worksheet sorting is specified for a #range (required parameter). The sort method
 * is specified with #sortMethod, case sensitivity with #caseSensitive.
 *
 * Sort can have up to 64 sort conditions, each of them is applied one by one.
 */
class QXLSX_EXPORT SortState
{
public:

    /**
     * @brief The SortMethod enum specifies the sort method.
     */
    enum class SortMethod
    {
        None, /**< The sort method is not specified*/
        Stroke, /**< Sort by stroke count of the characters. Only applies to
Chinese Simplified, Chinese Traditional, and Japanese. */
        PinYin, /**< Sort phonetically. This is the default sort method. */
    };

    std::optional<bool> columnSort; // default="false"
    std::optional<Qt::CaseSensitivity> caseSensitive; // default="false"
    std::optional<SortMethod> sortMethod; // default="none"
    CellRange range; //required
    ExtensionList extLst;
    QList<SortCondition> sortConditions;

    SERIALIZE_ENUM(SortMethod, {
       {SortMethod::None,   "none"},
       {SortMethod::PinYin, "pinYin"},
       {SortMethod::Stroke, "stroke"}
    });

    bool operator==(const SortState &other) const;
    bool operator!=(const SortState &other) const;
    bool isValid() const;
    void write(QXmlStreamWriter &writer, const QString &name) const;
    void read(QXmlStreamReader &reader);
};

class AutoFilterPrivate;
/**
 * @brief The AutoFilter class represents the autofilter and sorting parameters for worksheet data.
 *
 * This class is _implicitly shareable_.
 *
 * # Autofilter
 *
 * If set, autofilter temporarily hides rows based on a filter criteria, which is
 * applied column by column to a table of data in the worksheet.
 *
 * Autofilter is applied to a range. See #setRange() and #range().
 * You can set filters to columns within #range() with #setFilter(). Or you can
 * use specialized methods like #setFilterByValues(), #setPredicate() etc.
 *
 * To remove filter use #removeFilter(). To get current filter use #filter().
 * To check if filtering was set for specific column use `filterType(column) != Filter::Type::Invalid`.
 *
 * #setShowFilterButton() and #setShowFilterOptions() let you to change the visibility
 * of UI elements in the column header.
 *
 * # Sorting
 *
 * Sorting is applied to a range. See #setSortRange() and #sortRange(). If sort range is invalid,
 * so is sort.
 *
 * To test whether sorting is enabled use #sortingEnabled() or ```sortState().isValid()```.
 *
 * Sorting can have up to 64 sort conditions. See #sortConditions(), #sortCondition(),
 * #addSortCondition(), #setSortCondition(), #removeSortCondition() and #clearSortConditions()
 * to find out how to manipulate sort conditions.
 *
 * There is a convenience method #setSorting() to easily add simple sorting by values for a range
 * of cells. This method replaces all previously defined sort parameters.
 */
class QXLSX_EXPORT AutoFilter
{
public:
    /**
     * @brief creates invalid autofilter.
     * @note
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
     *
     * The range defines columns that will get the autofilter options and rows
     * that will be subjected to autofiltering.
     *
     * @return the cell range the filter is applied to. Can be invalid, this means that
     * autofilter is applied to the first column starting from the first row,
     * and all filters except for the first one will be ignored.
     */
    CellRange range() const;
    /**
     * @brief sets the cell range the autofilter is applied to.
     *
     * The range defines columns that will get the autofilter options and rows
     * that will be subjected to autofiltering.
     *
     * @param range the cell range the filter is applied to. Can be invalid, this means that
     * autofilter is applied to the first column starting from the first row.
     * @note Column headers should also be included into the range, as they will get
     * autofilter option buttons.
     */
    void setRange(const CellRange &range);
    /**
     * @brief returns whether to hide the filtering button in the column header.
     * @param column zero-based column index in the autofilter #range().
     * @return valid bool if the parameter was set, nullopt otherwise.
     *
     * The default value is false (filter button is shown).
     */
    std::optional<bool> hideFilterButton(int column) const;
    /**
     * @brief sets whether to hide the filtering button in the column header.
     * @param column zero-based column index in the autofilter #range().
     * @param hide If true, then filtering button is hidden.
     *
     * If not set, false is assumed.
     */
    void setHideFilterButton(int column, bool hide);
    /**
     * @overload
     * @brief sets whether to show the filtering button in the headers of all columns
     * in the #range().
     * @param show If true, then filtering button is shown.
     *
     * If not set, true is assumed.
     */
    void setShowFilterButton(bool show);
    /**
     * @brief returns whether to show filtering options in the UI element.
     * @param column zero-based column index in the autofilter #range().
     * @return valid bool if the parameter was set, nullopt otherwise.
     *
     * The default value is true.
     */
    std::optional<bool> showFilterOptions(int column) const;
    /**
     * @brief sets whether to show filtering options in the UI element.
     * @param column zero-based column index in the autofilter #range().
     * @param show If true, then filtering options are shown.
     *
     * If not set, true is assumed.
     */
    void setShowFilterOptions(int column, bool show);
    /**
     * @brief sets whether to show filtering options in the UI element of all columns
     * in the #range().
     * @param show If true, then filtering options are shown.
     *
     * If not set, true is assumed.
     */
    void setShowFilterOptions(bool show);

    /**
     * @brief sets @a column filtering by top N (percent or number of items).
     * @param column zero-based column index in the autofilter #range().
     * @param value If @a usePercents is true, sets N persents to filter by. If
     * @a usePercents is false, sets N items to filter by.
     * @param filterBy The actual cell value in the #range() which is used to
     * perform the comparison for this filter. Allows to narrow down the filtering criterion.
     * @param usePercents If true, then filters by top @a value percents. If false,
     * then filters by top @a value items.
     */
    void setFilterByTopN(int column, double value, std::optional<double> filterBy = std::nullopt,
                         bool usePercents = false);
    /**
     * @brief sets @a column filtering by bottom N (percent or number of items).
     * @param column zero-based column index in the autofilter #range().
     * @param value If @a usePercents is true, sets N persents to filter by. If
     * @a usePercents is false, sets N items to filter by.
     * @param filterBy The actual cell value in the #range() which is used to
     * perform the comparison for this filter. Allows to narrow down the filtering criterion.
     * @param usePercents If true, then filters by bottom @a value percents.
     * If false, then filters by bottom @a value items.
     */
    void setFilterByBottomN(int column, double value, std::optional<double> filterBy = std::nullopt,
                            bool usePercents = false);

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
     * such as #setFilterByTopN(), #setFilterByBottomN(), #setFilterByValues()
     * etc.
     */
    Filter &filter(int column);
    /**
     * @brief sets filter for @a column.
     * @param column zero-based index of a column in the autofilter #range().
     * @param filter Filter object.
     * @note This is a low-level method. It is easier to use the specialized methods
     * like #setPredicate(), #setFilterByValues() etc.
     */
    void setFilter(int column, const Filter &filter);

    /// Sorting methods

    /**
     * @brief returns wheter sorting was enabled. Equivalent to ```sortState().isValid()```
     * @return true if sorting is enabled.
     * @note Sort parameters may have no sort conditions specified, but if the sort
     * range is valid, Excel applied default sorting to this range, thus to set the
     * simples sorting it is enough to specify the sort range.
     */
    bool sortingEnabled() const;
    /**
     * @brief clears all sort parameters.
     */
    void clearSortState();
    /**
     * @brief returns the sort parameters.
     * @return A copy of the SortState object.
     */
    SortState sortState() const;
    /**
     * @brief returns the sort parameters.
     * @return A reference to the SortState object.
     */
    SortState &sortState();
    /**
     * @brief sets the sort parameters.
     *
     * If @a sort is invalid, does nothing and returns false. To remove sorting use #clearSortState()
     * or ```setSortState({})```.
     *
     * @param sort A SortState object.
     */
    bool setSortState(const SortState &sort);
    /**
     * @brief sets the sort range, which sorting will be applied to.
     *
     * If @a range is invalid, does nothing. To remove sorting use #clearSortState()
     * or ```setSortState({})```.
     *
     * @param range a valid CellRange object.
     * @return true on success (@a range is valid and includes all sort conditions ranges).
     */
    bool setSortRange(const CellRange &range);
    /**
     * @brief sets the sort range, which sorting is applied to.
     * @return A CellRange object.
     */
    CellRange sortRange() const;

    /**
     * @brief returns the list of sort conditions defined in the sort.
     * @return a list.
     */
    QList<SortCondition> sortConditions() const;
    /**
     * @brief returns the sort condition with @a index.
     * @param index must be within 0 and #sortConditionsCount()-1.
     * @return a reference to the sort condition.
     */
    SortCondition &sortCondition(int index);
    /**
     * @brief returns the sort condition with @a index.
     * @param index must be within 0 and #sortConditionsCount()-1.
     * @return  a copy of the SortCondition object if @a index is valid,
     * a default-constructed SortCondition otherwise.
     */
    SortCondition sortCondition(int index) const;
    /**
     * @brief returns the sort conditions count.
     * @return an int value. If it is negative or > 64, the file is ill-formed.
     */
    int sortConditionsCount() const;
    /**
     * @brief adds a new sort condition.
     * @param condition SortCondition object
     * @return true if @a condition is valid and was successfully added.
     */
    bool addSortCondition(const SortCondition& condition);
    /**
     * @brief sets sort condition with @a index.
     * @param index must be within 0 and #sortConditionsCount()-1.
     * @param condition SortCondition object
     * @return true if @a index and @a condition are valid and @a condition was successfully
     * set.
     */
    bool setSortCondition(int index, const SortCondition& condition);
    /**
     * @brief removes the sort condition with @a index
     * @param index must be within 0 and #sortConditionsCount()-1.
     * @return true if the sort condition was successfully removed.
     */
    bool removeSortCondition(int index);
    /**
     * @brief removes all sort conditions.
     */
    void clearSortConditions();

    bool addSortByValue(const CellRange &range,
                        Qt::SortOrder sortOrder = Qt::AscendingOrder,
                        const QString &customList = "");
    bool addSortByCellColor(const CellRange &range, int formatIndex, Qt::SortOrder sortOrder = Qt::AscendingOrder);
    bool addSortByFontColor(const CellRange &range, int formatIndex, Qt::SortOrder sortOrder = Qt::AscendingOrder);
    //TODO: bool addSortByIcon()

    std::optional<bool> sortBycolumns() const;
    void setSortByColumns(bool sortByColumns);

    std::optional<Qt::CaseSensitivity> caseSensitivity();
    void setCaseSensitivity(Qt::CaseSensitivity caseSensitivity);

    std::optional<SortState::SortMethod> sortMethod() const;
    void setSortMethod(SortState::SortMethod sortMethod);

    /**
     * @brief sets sort parameters.
     *
     * This method is intended to quickly add sorting to the worksheet.
     * It uses the following parameters:
     * - SortState::range is @a sortRange;
     * - SortState::columnSort is unspecified (defaults to false);
     * - SortState::caseSensitive is @a caseSensitive;
     * - SortState::sortMethod is unspecified (defaults to SortMethod::PinYin;
     * - sort condition is only one:
     *     - SortCondition::range is @a sortBy;
     *     - SortCondition::descending is @a descending;
     *     - SortCondition::sortBy is unspecified (defaults to SortCondition::SortBy::Value);
     *     - other SortCondition parameters are unspecified.
     *
     * @param sortRange The range to be sorted.
     * @param sortBy The range to sort by.
     * @param sortOrder Ascending/descending sort.
     * @param caseSensitivity the sort case sensitivity.
     * @return true if the ranges are valid and sorting is successful.
     * @note To fine-tune sorting, use #sortState(), #setSortState(), #addSortCondition(),
     * #sortCondition(int index) methods.
     */
    bool setSorting(const CellRange &sortRange,
                    const CellRange &sortBy = CellRange(),
                    Qt::SortOrder sortOrder = Qt::AscendingOrder,
                    Qt::CaseSensitivity caseSensitivity = Qt::CaseInsensitive);

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
Q_DECLARE_TYPEINFO(QXlsx::AutoFilter, Q_MOVABLE_TYPE);

#endif // XLSXAUTOFILTER_H
