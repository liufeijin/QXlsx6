// xlsxchart.h

#ifndef QXLSX_CHART_H
#define QXLSX_CHART_H

#include <QtGlobal>
#include <QXmlStreamReader>
#include <QXmlStreamWriter>

#include "xlsxabstractooxmlfile.h"
#include "xlsxshapeformat.h"
#include "xlsxlineformat.h"
#include "xlsxaxis.h"
#include "xlsxlabel.h"
#include "xlsxseries.h"
#include "xlsxlegend.h"
#include "xlsxpagesetup.h"
#include "xlsxpagemargins.h"
#include "xlsxabstractsheet.h"

namespace QXlsx {

class AbstractSheet;
class Worksheet;
class ChartPrivate;
class CellRange;
class DrawingAnchor;
class Workbook;

class QXLSX_EXPORT UpDownBar
{
public:
    std::optional<int> gapWidth;
    ShapeFormat upBar;
    ShapeFormat downBar;

    void read(QXmlStreamReader &reader);
    void write(QXmlStreamWriter &writer, const QString &name) const;
    bool isValid() const;
};

/**
 * @brief The DataTable class holds the properties of the chart data table, that is
 * the table with the chart series data.
 */
class QXLSX_EXPORT DataTable
{
public:
    std::optional<bool> showHorizontalBorder; /**< table cells horizontal border visibility */
    std::optional<bool> showVerticalBorder; /**< table cells vertical border visibility */
    std::optional<bool> showOutline; /**< table outline border visibility */
    std::optional<bool> showKeys; /**< legend keys visibility */
    ShapeFormat shape; /**< line and fill of the data table cells */
    TextFormat textProperties; /**< text, paragraph and character properties */

    void read(QXmlStreamReader &reader);
    void write(QXmlStreamWriter &writer, const QString &name) const;
    bool isValid() const;
private:
    ExtensionList extension;
};

/**
 * @brief The Chart class represents a chart in the worksheet or the chartsheet.
 *
 * ## Subcharts and axes
 *
 * Though the chart can have as many axes as you wish, only a subset of them can be used
 * to display the chart series. So a notion of subchart is used.
 *
 * Each subchart has a type (see Chart::Type enum), a number of axes references (see
 * #addSubchart(), #setSubchartAxes(), #addDefaultAxes()), a number of associated
 * series (see #addSeries() overloads) and parameters specific to the chart (f.e. #bubbleScale()).
 *
 * Each subchart type requires different count and types of axes. Here is the table of requirements:
 *
 * Chart type | axes count and types
 * -----------|---------------------
 * Line, Bar, Area, Radar, Stock, Bubble | 2: Category and Value
 * Line3D, Surface3D | 3: Category, Value and Series
 * Bar3D, Area3D, Surface | 2-3: Category, Value and optionally Series
 * Scatter | 2: Value and Value
 * Pie, Pie3D, Doughnut, OfPie | 0 (no axes required)
 *
 * The easiest way of adding chart is to use only one subchart:
 * @code
 * Chart *chart = sheet->insertChart(3, 3, QSize(300, 300));
 * chart->addSubchart(Chart::Type::Line);
 * chart->addSeries(CellRange("A1:A9"));
 * chart->addSeries(CellRange("B1:B9"));
 * chart->addSeries(CellRange("C1:C9"));
 * @endcode
 *
 * The code above also creates default axes that are needed for the line chart (category axis
 * and value axis). You can access these axes via #axis() and #axes() methods.
 *
 * If more than one subchart is needed, f.e. if you need to move some series to the right/top axis
 * or to place series of different types on the same chart, then adding a new subchart comes in handy.
 *
 * This code creates two subcharts of different types.
 *
 * @code
 * Chart *chart = sheet->insertChart(3, 3, QSize(300, 300));
 * // Create the first subchart
 * chart->addSubchart(Chart::Type::Line);
 * chart->addSeries(CellRange("A1:A9"));
 *
 * // Create the second subchart
 *
 * // the second subchart has different type, but uses the same axes as the first one.
 * chart->addSubchart(Chart::Type::Bar, chart->subchartAxes(0));
 * chart->addSeries(CellRange("B1:B9"), nullptr, false, false, true, 1);
 * @endcode
 *
 * The code below adds the right axis to the chart
 *
 * @code
 * Chart *chart = sheet->insertChart(3, 3, QSize(300, 300));
 * // Create the first subchart
 * chart->addSubchart(Chart::Type::Line);
 *
 * // Create the second subchart of the same type, but with different set of axes
 * auto rightAxis = chart->addAxis(QXlsx::Axis::Type::Value, QXlsx::Axis::Position::Right);
 * auto bottomAxis = chart->axis(QXlsx::Axis::Position::Bottom);
 * rightAxis.setCrossAxis(bottomAxis->id()); // required
 * rightAxis.setCrossesType(QXlsx::Axis::CrossesType::Maximum); //this actually moves rightAxis to the right
 * chart->addSubchart(QXlsx::Chart::Type::Line, {bottomAxis->id(), rightAxis.id()});
 *
 * // Add series
 * chart->addSeries(CellRange("A1:A9"), nullptr, false, false, true, 0);
 * chart->addSeries(CellRange("B1:B9"), nullptr, false, false, true, 1);
 * @endcode
 */
class QXLSX_EXPORT Chart : public AbstractOOXmlFile
{
    Q_DECLARE_PRIVATE(Chart)
public:
    /**
     * @brief The Type enum specifies the chart type
     */
    enum class Type
    {
        None = 0, /**< Invalid chart type*/
        Area, /**< Area chart */
        Area3D, /**< Area3D chart*/
        Line, /**< Line chart */
        Line3D, /**< Line3D chart*/
        Stock, /**< Stock chart */
        Radar, /**< Radar chart*/
        Scatter, /**< Scater chart */
        Pie, /**< Pie chart */
        Pie3D, /**< Pir3D chart */
        Doughnut, /**< Doughnut chart */
        Bar, /**< Bar chart */
        Bar3D, /**< Bar3D chart*/
        OfPie, /**< Pie of pie chart (or bar of pie based on the ofPieChart parameter.) */
        Surface, /**< Surface chart */
        Surface3D, /**< Surface3D chart*/
        Bubble, /**< Bubble chart */
    };
    /**
     * @brief The DisplayBlanksAs enum specifies how blank cells are plotted on the chart.
     */
    enum class DisplayBlanksAs
    {
        Span, /**< Blank cells are plotted as previous cells (previous cell value
is spanned to fill the gap). */
        Gap, /**< Blank cells create gaps in the chart (not plotted at all.) */
        Zero /**< Blank cells are plotted as zeroes. */
    };
    /**
     * @brief The Grouping enum specifies the kind of grouping for an area or line chart.
     */
    enum class Grouping
    {
        Standard, /**< Chart series are drawn on the value axis */
        Stacked, /**< Chart series are drawn next to each other on the value axis */
        PercentStacked /**< Chart series are drawn next to each
other along the value axis and scaled to total 100% */
    };
    /**
     * @brief The BarGrouping enum specifies the kind grouping for a bar chart.
     */
    enum class BarGrouping
    {
        Standard, /**< Chart series are drawn next to each other on the depth axis*/
        Stacked, /**< Chart series are drawn next to each other on the value axis*/
        PercentStacked, /**< Chart series are drawn next to each
other along the value axis and scaled to total 100%*/
        Clustered /**< Chart series are drawn next to each other along the category axis */
    };
    /**
     * @brief The BubbleSizeRepresents enum specifies how the bubble size values
     * are represented on the chart.
     */
    enum class BubbleSizeRepresents
    {
        Area, /**< Series::bubbleSizeDataSource values are used as the bubbles area. */
        Width /**< Series::bubbleSizeDataSource values are used as the bubbles width (diameter).*/
    };
    /**
     * @brief The OfPieType enum specifies the second pie type on a OfPie chart.
     */
    enum class OfPieType
    {
        Bar,  /**< @brief The second pie is a bar chart. */
        Pie /**< @brief The second pie is a pie chart.*/
    };
    /**
     * @brief The SplitType enum specifies how to split the data points between
     * the first pie and second pie or bar on a OfPie chart.
     */
    enum class SplitType
    {
        Auto, /**< @brief automatic selection of data points that go to the second pie.  */
        Custom, /**< @brief the second pie contains data points selected manually via
                     setCustomSplit().*/
        Position, /**< @brief last splitPos() data points go to the second pie.*/
        Percent, /**< @brief data points with value less than splitPos() percents
                      of the sum of all values go to the second pie. */
        Value /**< @brief data points with value less than splitPos() go to the second pie. */
    };
    /**
     * @brief The BarDirection enum specifies whether the series form a Bar (horizontal)
     * chart or a Column (vertical) chart
     */
    enum class BarDirection
    {
        Bar, /**< Bars of a bar chart are horizontal.*/
        Column /**< Bars of a bar chart are vertical. This is the default value. */
    };
    /**
     * @brief The ScatterStyle enum specifies the style of the scatter chart.
     */
    enum class ScatterStyle
    {
        None, /**< No lines or markers are drawn. */
        Line, /**< Points are connected with straight lines, no markers shown. */
        LineMarker, /**< Points are connected with straight lines, each point is marked with a marker. */
        Marker, /**< Points are not connected with lines, each point is marked with a marker. */
        Smooth, /**< Points are connected with smoothed lines, no markers shown. */
        SmoothMarker, /**< Points are connected with smoothed lines, each point is marked with a marker. */
    };
    /**
     * @brief The RadarStyle enum  specifies what type of radar chart shall be drawn.
     */
    enum class RadarStyle
    {
        Standard, /**< Points are connected with straight lines, no markers shown. */
        Marker, /**< Points are connected with straight lines, each point is marked with a marker. */
        Filled /**< Points are connected with straight lines, areas under lines are filled. */
    };

private:
    friend class AbstractSheet;
    friend class Worksheet;
    friend class Chartsheet;
    friend class DrawingAnchor;
    friend class LineFormat;
    SERIALIZE_ENUM(DisplayBlanksAs,
    {
        {DisplayBlanksAs::Gap, "gap"},
        {DisplayBlanksAs::Span, "span"},
        {DisplayBlanksAs::Zero, "zero"},
    });
    SERIALIZE_ENUM(Type,
    {
        {Type::Bar, "barChart"},
        {Type::Pie, "pieChart"},
        {Type::Area, "areaChart"},
        {Type::Line, "lineChart"},
        {Type::None, ""},
        {Type::Bar3D, "bar3DChart"},
        {Type::OfPie, "ofPieChart"},
        {Type::Pie3D, "pie3DChart"},
        {Type::Radar, "radarChart"},
        {Type::Stock, "stockChart"},
        {Type::Area3D, "area3DChart"},
        {Type::Bubble, "bubbleChart"},
        {Type::Line3D, "line3DChart"},
        {Type::Scatter, "scatterChart"},
        {Type::Surface, "surfaceChart"},
        {Type::Doughnut, "doughnutChart"},
        {Type::Surface3D, "surface3DChart"},
    });
    SERIALIZE_ENUM(Grouping,
    {
        {Grouping::Standard, "standard"},
        {Grouping::Stacked, "stacked"},
        {Grouping::PercentStacked, "percentStacked"},
    });
    SERIALIZE_ENUM(BarGrouping,
    {
        {BarGrouping::Standard, "standard"},
        {BarGrouping::Stacked, "stacked"},
        {BarGrouping::PercentStacked, "percentStacked"},
        {BarGrouping::Clustered, "clustered"},
    });
    SERIALIZE_ENUM(SplitType,
    {
        {SplitType::Auto, "auto"},
        {SplitType::Custom,"cust"},
        {SplitType::Position,"pos"},
        {SplitType::Percent,"percent"},
        {SplitType::Value, "val"},
    });
    SERIALIZE_ENUM(ScatterStyle,
    {
        {ScatterStyle::None,"none"},
        {ScatterStyle::Line,"line"},
        {ScatterStyle::LineMarker,"lineMarker"},
        {ScatterStyle::Marker,"marker"},
        {ScatterStyle::Smooth,"smooth"},
        {ScatterStyle::SmoothMarker,"smoothMarker"},
    });
    SERIALIZE_ENUM(RadarStyle,
    {
        {RadarStyle::Standard,"standard"},
        {RadarStyle::Marker,"marker"},
        {RadarStyle::Filled,"filled"},
    });
private:
    Chart(AbstractSheet *parent, CreateFlag flag);
    Chart &operator=(const Chart &other);
public:
    /**
     * Destroys the chart.
     */
    ~Chart();
public:
    /**
     * @brief Adds new subchart of @a type and with axes specified with @a axesIDs.
     * @param type subchart type.
     * @param axesIDs ids of axes for this subchart. Each subchart must have a
     * subset of axes defined in the chart. If axesIDs is empty, then new set of default
     * axes are created for the new subchart.
     */
    void addSubchart(Type type, const QList<int> &axesIDs = {});
    /**
     * @brief sets @a axesIDs to the subchart with index @a subchartIndex.
     * @param subchartIndex zero-based subchart index.
     * @param axesIDs a list of required axes IDs.
     * @sa #addDefaultAxes().
     */
    void setSubchartAxes(int subchartIndex, const QList<int> &axesIDs);
    /**
     * @brief returns axes associated with a subchart
     * @param subchartIndex zero-based subchart index.
     * @return a list of axes IDs.
     * @sa #setSubchartAxes(), #addDefaultAxes(), #addSubchart().
     */
    QList<int> subchartAxes(int subchartIndex) const;
    /**
     * @brief returns whether subchart with @a subchartIndex has series added to it.
     * @param subchartIndex zero-based index of a subchart.
     * @return `true` if series were added to the subchart. `false` if subchart has no
     * series or @a subchartIndex is invalid.
     */
    bool subchartHasSeries(int subchartIndex) const;
    /**
     * @brief removes subchart.
     * @param subchartIndex zero-based index of a subchart.
     * @return `true` if removal was successful, `false` if @a subchartIndex is invalid.
     * @note After that the chart may have no subcharts. See #subchartsCount(). This method
     * also removes any series that were added to the subchart.
     */
    bool removeSubchart(int subchartIndex);
    /**
     * @brief returns the subcharts count.
     */
    int subchartsCount() const;
    /**
     * @brief returns the subchart type
     * @param subchartIndex zero-based index of a subchart.
     * @return Type enum value. If @a subchartIndex is invalid, returns Type::None.
     */
    Type subchartType(int subchartIndex) const;
    /**
     * @brief sets the subchart type
     * @param subchartIndex zero-based index of a subchart.
     * @param type Type enum value.
     * @note Check axes that this subchart is associated with after invoking this method.
     */
    void setSubchartType(int subchartIndex, Type type);

    /**
     * @brief sets line format to the chart space (the outer area of the chart).
     *
     * This is a convenience method, equivalent to `chartShape()->setLineFormat(format)`
     * @param format
     */
    void setChartLineFormat(const LineFormat &format);
    /**
     * @brief sets line format to the plot area (the inner area of the chart
     * where actual series plotting occurs).
     *
     * This is a convenience method, equivalent to `plotAreaShape()->setLineFormat(format)`
     * @param format
     */
    void setPlotAreaLineFormat(const LineFormat &format);
    /**
     * @brief sets shape format to the chart space (the outer area of the chart).
     * @param shape
     */
    void setChartShape(const ShapeFormat &shape);
    /**
     * @brief returns shape format of the chart space (the outer area of the chart).
     * @return reference to the ShapeFormat object.
     * @note You can use this method to set the default chartShape, as the method will create
     * a valid ShapeFormat object with its parameters not set.
     */
    ShapeFormat &chartShape();
    /**
     * @brief returns the header/footer parameters of the chart space (the outer area of the chart).
     * @return reference to the HeaderFooter object.
     */
    HeaderFooter &headerFooter();
    /**
     * @brief returns the page margins parameters of the chart space (the outer area of the chart).
     * @return reference to the PageMargins object.
     */
    PageMargins &pageMargins();
    /**
     * @brief returns the paper and printing parameters of the chart space (the outer area of the chart).
     * @return reference to the PageSetup object.
     */
    PageSetup &pageSetup();

    /**
     * @brief sets shape format to the plot area (the inner area of the chart
     * where actual plotting occurs).
     * @param shape
     */
    void setPlotAreaShape(const ShapeFormat &shape);
    /**
     * @brief returns shape format of the plot area (the inner area of the chart
     * where actual plotting occurs).
     * @return reference to the ShapeFormat object.
     * @note You can use this method to set the default plotAreaShape, as the method will create
     * a valid ShapeFormat object with its parameters not set.
     */
    ShapeFormat &plotAreaShape();

    /**
     * @brief returns text format of the chart.
     * @return copy of a TextFormat object.
     */
    TextFormat textProperties() const;
    /**
     * @brief returns text format of the chart.
     * @return reference to the TextFormat object.
     */
    TextFormat &textProperties();
    /**
     * @brief sets text format of the chart.
     * @param textProperties
     */
    void setTextProperties(const TextFormat &textProperties);

    /**
     * @brief adds an empty series.
     * @param subchart zero-based index of a subchart into which the series are being added.
     * @return index of the added series.
     */
    int addSeries(int subchart = 0);

    /**
     * @overload
     * @brief adds one or more series with data defined by @a range.
     * @param range valid CellRange
     * @param sheet data source for @a range
     * @param firstRowContainsHeaders specifies that the 1st row (the 1st column
     * if @a columnBased is `false`) contains the series titles and should be
     * excluded from the series data.
     * @param firstColumnContainsCategoryData specifies that the 1st column
     * (the 1st row if @a columnBased is `false`) contains data for the category (x)
     * axis.
     * @param columnBased specifies that @a range should be treated as column-based
     * or row-based: if @a columnBased is `true`, the first column is category data
     * (if @a firstColumnContainsCategoryData is set to `true`),
     * other columns are value data for new series. If @a columnBased is `false`, the
     * first row is category data, other rows are value data for new series.
     * @param subchart zero-based index of a subchart into which the series are being added.
     * @note the series will have the type specified by the subchart type.
     * @return a list of indexes the added series.
     */
    QList<int> addSeries(const CellRange &range, AbstractSheet *sheet = nullptr,
                   bool firstRowContainsHeaders = false,
                   bool firstColumnContainsCategoryData = false,
                   bool columnBased = true,
                   int subchart = 0);
    /**
     * @overload
     * @brief adds series defined by keyRange and valRange.
     * @param keyRange category range or x axis range.
     * @param valRange value range or y axis range.
     * @param sheet data sheet reference.
     * @param keyRangeIncludesHeader if `true`, the first row or column is used as a series name reference.
     * @param subchart zero-based index of a subchart into which the series are being added.
     * @return index of the added series or -1 if no series was added.
     * @note The series will have the type specifies by the subchart type.
     *
     * @note To set bubbleSizeDataSource to the BubbleChart series (the 3rd range)
     * use #addSeries() method and manually set all three ranges.
     */
    int addSeries(const CellRange &keyRange, const CellRange &valRange,
                      AbstractSheet *sheet = NULL, bool keyRangeIncludesHeader = false,
                      int subchart = 0);
    /**
     * @brief returns the chart series by its index.
     * @param index the series index starting from 0.
     * @return pointer to the series if such series exists, `nullptr` otherwise.
     */
    Series *series(int index);

    /**
     * @brief returns the chart's series count.
     * @return
     */
    int seriesCount() const;

    /**
     * @brief removes series with @a index from the chart.
     * @param index zero-based series index.
     * @return `true` if such a series was found and successfully deleted, `false` otherwise.
     */
    bool removeSeries(int index);
    /**
     * @overload
     * @brief removes #series from the chart.
     * @param series
     * @return `true` if such a series was found and successfully deleted, `false` otherwise.
     */
    bool removeSeries(const Series &series);
    /**
     * @brief removes all series from the chart.
     */
    void removeAllSeries();

    /**
     * @brief sets axes for the specific series
     * @param series pointer to the series
     * @param axesIds a list of 0, 2 or 3 items.
     * @note @a axesIds are ids, not indexes. @see Axis::id().
     * This method also adds a new subchart if no subchart with @a axesIds was found.
     *
     */
    void setSeriesAxesIDs(Series *series, const QList<int> &axesIds);

    /**
     * @brief changes the series order in which it is drawn on the chart.
     * @param oldOrder
     * @param newOrder
     */
    void moveSeries(int oldOrder, int newOrder);

    /**
     * @brief adds default axes required for the chart @a type.
     *
     * If you want to later fine-tune them, use axis(int idx) method.
     *
     * Here is the table of axes requirements for each chart type
     *
     * Chart type | axes count and types
     * -----------|---------------------
     * Line, Bar, Area, Radar, Stock, Bubble | 2: Category (bottom) and Value (left)
     * Line3D, Surface3D | 3: Category (bottom), Value (left) and Series (bottom)
     * Bar3D, Area3D, Surface | 2-3: Category (bottom), Value (left) and Series (bottom)
     * Scatter | 2: Value (bottom) and Value (left)
     * Pie, Pie3D, Doughnut, OfPie | 0 (no axes required)
     *
     * @param type chart type.
     * @return list of axes IDs. Can be empty.
     * @note This method does not check if the required axes are already added to the chart,
     * it simply creates a new set of axes.
     */
    QList<int> addDefaultAxes(Type type);

    /**
     * @brief adds new axis with specified parameters.
     * @param type axis type: Cat, Val, Date or Ser.
     * @param pos Axis position: Bottom, Left, Top or Right.
     * @param title optional axis title.
     * @return index of the newly added axis.
     */
    int addAxis(Axis::Type type, Axis::Position pos, QString title = QString());
    /**
     * @brief adds new axis created elsewhere.
     *
     * The method can be used to duplicate existing axis, change the type and position of the copy
     * and add the result to the chart. The method makes sure new axis will have a unique ID.
     * @param axis
     * @return index of the newly added axis.
     */
    int addAxis(const Axis &axis);
    /**
     * @brief returns axis that has index idx
     * @param idx valid index (0 <= idx < axesCount()).
     * @return pointer to the axis if such axis exists, `nullptr` otherwise.
     */
    Axis *axis(int idx);
    /**
     * @overload
     * @brief returns axis that has position @a pos.
     * @param pos Axis::Position.
     * @note A chart can have several axes positioned at @a pos. This method returns
     * _the first added_ axis that has position @a pos.
     * @return pointer to the axis if such axis exists, `nullptr` otherwise.
     */
    Axis *axis(Axis::Position pos);
    /**
     * @overload
     * @brief returns an axis that has type @a type.
     * @param type Axis::Type.
     * @note A chart can have several axes of the same type (f.e. scatter chart
     * usually has 2 value axes). This method returns _the first added_ axis of the
     * type.
     * @return pointer to the axis if such axis exists, `nullptr` otherwise.
     */
    Axis *axis(Axis::Type type);
    /**
     * @overload
     * @brief returns an axis that has position @a pos and @a type.
     * @param pos Axis::Position.
     * @param type Axis::Type.
     * @note A chart can have several axes of the same position and type. This
     * method returns _the first added_ axis of these position and type.
     * @return pointer to the axis if such axis exists, `nullptr` otherwise.
     */
    Axis *axis(Axis::Type type, Axis::Position pos);
    /**
     * @brief returns the list of axes defined in the chart.
     * @return a list of axes. As Axis is an explicitly shareable class, no actual copying
     * of axes occurs. @see Axis.
     */
    QList<QXlsx::Axis> axes() const;
    /**
     * @brief returns the list of axes ids defined in the chart.
     * @return a list of axes IDs.
     */
    QList<int> axesIDs() const;

    /**
     * @brief tries to remove axis and returns `true` if axis has been removed.
     * @param axisID the id of the axis to be removed.
     * @return `true` if axis has been removed, `false` otherwise (no axis with this id or axis is used
     * in some series).
     * @note This method does not check if any series uses this axis. Use #seriesThatUseAxis()
     * to do the checking.
     */
    bool removeAxis(int axisID);
    /**
     * @overload
     * @brief tries to remove axis and returns `true` if axis has been removed.
     * @param axis the pointer to the axis to be removed.
     * @return `true` if axis has been removed, `false` otherwise (no such axis or axis is used
     * in some series).
     * @note This method does not check if any series uses this axis. Use #seriesThatUseAxis()
     * to do the checking.
     */
    bool removeAxis(Axis *axis);
    /**
     * @brief removes all axes defined in the chart.
     *
     * This method removes all axes defined in the chart and clears the series references to the axes.
     * After that until you add the required axes and set these axes for the chart series
     * the chart is ill-formed.
     */
    void removeAxes();
    /**
     * @brief returns the list of series that use axis with @a axisID.
     * @param axisID axis ID. @see Axis::id().
     * @return list of series indexes.
     */
    QList<int> seriesThatUseAxis(int axisID) const;
    /**
     * @overload
     * @brief returns the list of series that use @a axis.
     * @param axis pointer to the axis.
     * @return list of series indexes.
     */
    QList<int> seriesThatUseAxis(Axis *axis) const;
    /**
     * @brief returns the axes count
     * @return
     */
    int axesCount() const;

    /**
     * @brief sets formatted title to the chart.
     * @param title
     */
    void setTitle(const Title &title);
    /**
     * @overload
     * @brief sets plain text title to the chart.
     * @param title
     */
    void setTitle(const QString &title);
    /**
     * @brief returns reference to the chart title
     * @return
     */
    Title &title();
    /**
     * @overload
     * @brief returns a copy of the chart's title
     * @return shallow copy of the Title object.
     */
    Title title() const;
    /**
     * @brief returns whether to show the automatically generated title of the chart.
     * @return
     */
    std::optional<bool> autoTitleDeleted() const;
    /**
     * @brief sets whether to show the automatically generated title of the chart.
     * @param value if `true` then the automatically generated title is not shown.
     * @note If the chart has only one series and autoTitleDeleted is set to `false`,
     * then the chart will have the title that contains the series name. Set
     * autoTitleDeleted to `true` to hide this title.
     */
    void setAutoTitleDeleted(bool value);

    /**
     * @brief returns whether only visible cells should be plotted on the chart.
     * @return
     *
     * If not set, the default value is `true`.
     */
    std::optional<bool> plotOnlyVisibleCells() const;
    /**
     * @brief sets whether only visible cells should be plotted on the chart.
     * @param value if `true`, only visible cells are plotted. If `false`, all cells are plotted,
     * including hidden ones.
     *
     * If not set, the default value is `true`.
     */
    void setPlotOnlyVisibleCells(bool value);

    /**
     * @brief returns  how blank cells shall be plotted on a chart.
     * @return
     *
     * If not set, the default value is DisplayBlanksAs::Zero
     */
    std::optional<DisplayBlanksAs> displayBlanksAs() const;
    /**
     * @brief sets how blank cells shall be plotted on a chart.
     * @param value DisplayBlanksAs value.
     *
     * If not set, the default value is DisplayBlanksAs::Zero.
     */
    void setDisplayBlanksAs(DisplayBlanksAs value);

    /**
     * @brief returns whether data labels over the maximum of the chart shall be shown.
     *
     * If not set, the default value is `true`.
     *
     * @return
     */
    std::optional<bool> showDataLabelsOverMaximum() const;
    /**
     * @brief sets whether data labels over the maximum of the chart shall be shown.
     * @param value
     *
     * If not set, the default value is `true`.
     */
    void setShowDataLabelsOverMaximum(bool value);

    /**
     * @brief adds default legend to the chart. Legend parameters will be set to
     * their default values.
     */
    void setDefaultLegend();
    /**
     * @brief moves the chart legend to position pos with chart overlay
     * @param position
     * @param overlay
     */
    void setLegend(Legend::Position position, bool overlay = false);
    /**
     * @brief returns reference to the chart legend.
     * @return reference to the chart legend.
     */
    Legend &legend();
    /**
     * @brief removes the chart legend.
     */
    void removeLegend();


    /**
     * @brief returns the chart's plot area layout.
     * @return a copy of the Layout object.
     */
    Layout layout() const;
    /**
     * @overload
     * @brief returns the chart's plot area layout.
     * @return reference to the Layout object.
     */
    Layout &layout();
    /**
     * @brief sets the chart's plot area layout.
     * @param layout
     */
    void setLayout(const Layout &layout);

    /**
     * @brief sets the visibility of the chart data table
     * @param visible if `true`, the chart data table is shown below the chart. If `false`,
     * the chart data table is hidden.
     *
     * @warning Setting the data table visibility to `false` deletes the current data table properties.
     *
     */
    void setDataTableVisible(bool visible);
    /**
     * @brief returns the chart data table.
     *
     * Changing the data table properties also makes it visible.
     *
     * @return reference to the data table.
     */
    DataTable &dataTable();
    /**
     * @brief returns the chart data table.
     *
     * @return copy of the chart data table.
     */
    DataTable dataTable() const;
    /**
     * @brief sets the chart data table.
     *
     * This method also makes dataTable visible.
     *
     * @param dataTable
     */
    void setDataTable(const DataTable &dataTable);

    /**
     * @brief styleID
     * @return
     */
    std::optional<int> styleID() const;
    /**
     * @brief setStyleID
     * @param id
     */
    void setStyleID(int id);

    /**
     * @brief returns whether the chart uses the 1904 date system.
     *
     * If the 1904 date system is used, then all dates and times shall be specified as a
     * decimal number of days since Dec. 31, 1903. If the 1904 date system is
     * not used, then all dates and times shall be specified as a decimal number
     * of days since Dec. 31, 1899.
     *
     * @return valid bool if the 1904 date system is set to `true` or `false`, `nullopt` otherwise.
     */
    std::optional<bool> date1904() const;
    /**
     * @brief sets that the chart uses the 1904 date system.
     *
     * This method sets that the chart uses the 1904 date system. If the 1904
     * date system is used, then all dates and times shall be specified as a
     * decimal number of days since Dec. 31, 1903. If the 1904 date system is
     * not used, then all dates and times shall be specified as a decimal number
     * of days since Dec. 31, 1899.
     *
     * @param isDate1904
     */
    void setDate1904(bool isDate1904);

    /**
     * @brief returns the primary editing language (f.e. "en-US")
     * which was used when this chart was last modified.
     *
     * @return language id or empty string if no last used language was set.
     */
    QString lastUsedLanguage() const;
    /**
     * @brief sets the primary editing language (f.e. "en-US")
     * which was used when this chart was last modified.
     *
     * @param lang language id
     */
    void setLastUsedLanguage(QString lang);

    /**
     * @brief returns whether the chart area has rounded corners.
     * @return valid bool if the property is set, `nullopt` otherwise.
     */
    std::optional<bool> roundedCorners() const;
    /**
     * @brief sets whether the chart area shall have rounded corners.
     * @param rounded
     */
    void setRoundedCorners(bool rounded);

    // Parameters specific to the subcharts

    /**
     * @brief returns the kind of grouping for a line or area chart.
     *
     * Applicable to: Area, Area3D, Line, Line3D.
     * @param subchartIndex zero-based index of a subchart.
     * @return valid Chart::Grouping if property is set, `nullopt` otherwise.
     */
    std::optional<Chart::Grouping> grouping(int subchartIndex) const;
    /**
     * @brief sets the kind of grouping for a line or area chart.
     *
     * Applicable to: Area, Area3D, Line, Line3D.
     * @param subchartIndex zero-based index of a subchart.
     * @param grouping Chart::Grouping
     */
    void setGrouping(int subchartIndex, Chart::Grouping grouping);

    /**
     * @brief returns the kind of grouping for a bar chart.
     *
     * Applicable to: Bar, Bar3D.
     * @param subchartIndex zero-based index of a subchart.
     * @return valid Chart::BarGrouping if property is set, `nullopt` otherwise.
     */
    std::optional<Chart::BarGrouping> barGrouping(int subchartIndex) const;
    /**
     * @brief sets the kind of grouping for a bar chart.
     *
     * Applicable to chart types: Bar, Bar3D.
     * @param subchartIndex zero-based index of a subchart.
     * @param grouping Chart::BarGrouping
     */
    void setBarGrouping(int subchartIndex, Chart::BarGrouping grouping);

    /**
     * @brief specifies that each data marker in the series has a different color.
     *
     * Applicable to chart types: Line, Line3D, Scatter, Radar, Bar, Bar3D, Area, Area3D,
     * Pie, Pie3D, Doughnut, OfPie, Bubble.
     * @param subchartIndex zero-based index of a subchart.
     * @return valid bool value if property is set, `nullopt` otherwise.
     *
     * @note by default varyColors is not set. But in order to get sensible charts
     * all new Pie, Pie3D, Doughnut, OfPie, Bubble charts automatically set varyColors to `true`.
     *
     */
    std::optional<bool> varyColors(int subchartIndex) const;
    /**
     * @brief sets that each data marker in the series has a different color.
     *
     * Applicable to chart types: Line, Line3D, Scatter, Radar, Bar, Bar3D, Area, Area3D,
     * Pie, Pie3D, Doughnut, OfPie, Bubble.
     * @param subchartIndex zero-based index of a subchart.
     * @param varyColors `true` if each data marker in the series has a different color, `false`
     * if every data markers in the series have the same color.
     */
    void setVaryColors(int subchartIndex, bool varyColors);

    /**
     * @brief returns the entire chart labels properties.
     *
     * This element serves as a root element that specifies the settings for the
     * data labels for an entire series or the entire chart.  It contains child
     * elements that specify the specific formatting and positioning settings.
     *
     * To set the series specific labels use Series::labels() and Series::setLabels().
     *
     * Applicable to chart types: Line, Line3D, Stock, Scatter, Radar, Bar, Bar3d,
     * Area, Area3D, Pie, Pie3D, Doughnut, OfPie, Bubble.
     * @param subchartIndex zero-based index of a subchart.
     * @return a copy of the chart labels.
     *
     */
    Labels labels(int subchartIndex) const;
    /**
     * @brief returns the entire chart labels properties.
     *
     * This element serves as a root element that specifies the settings for the
     * data labels for an entire series or the entire chart.  It contains child
     * elements that specify the specific formatting and positioning settings.
     *
     * To set the series specific labels use Series::labels() and Series::setLabels().
     *
     * Applicable to chart types: Line, Line3D, Stock, Scatter, Radar, Bar, Bar3d,
     * Area, Area3D, Pie, Pie3D, Doughnut, OfPie, Bubble.
     * @param subchartIndex zero-based index of a subchart.
     * @return reference to the chart labels.
     */
    Labels &labels(int subchartIndex);
    /**
     * @brief sets the entire chart labels properties.
     *
     * This element serves as a root element that specifies the settings for the
     * data labels for an entire series or the entire chart.  It contains child
     * elements that specify the specific formatting and positioning settings.
     *
     * To set the series specific labels use Series::labels() and Series::setLabels().
     *
     * Applicable to chart types: Line, Line3D, Stock, Scatter, Radar, Bar, Bar3d,
     * Area, Area3D, Pie, Pie3D, Doughnut, OfPie, Bubble.
     * @param subchartIndex zero-based index of a subchart.
     * @param labels Labels properties object.
     */
    void setLabels(int subchartIndex, const Labels &labels);

    /**
     * @brief returns the chart drop lines.
     *
     * Applicable to chart types: Line, Line3D, Stock, Area, Area3D.
     * @param subchartIndex zero-based index of a subchart.
     * @return a copy of the chart dropLines.
     */
    ShapeFormat dropLines(int subchartIndex) const;
    /**
     * @brief returns the chart drop lines.
     *
     * Applicable to chart types: Line, Line3D, Stock, Area, Area3D.
     * @param subchartIndex zero-based index of a subchart.
     * @return a reference to the chart dropLines.
     */
    ShapeFormat &dropLines(int subchartIndex);
    /**
     * @brief sets the chart drop lines.
     *
     * Applicable to chart types: Line, Line3D, Stock, Area, Area3D.
     * @param subchartIndex zero-based index of a subchart.
     * @param dropLines If not valid, sets default drop lines.
     * @note To remove drop lines use #removeDropLines().
     */
    void setDropLines(int subchartIndex, const ShapeFormat &dropLines);
    /**
     * @brief removes the chart drop lines.
     *
     * Applicable to chart types: Line, Line3D, Stock, Area, Area3D.
     * @param subchartIndex zero-based index of a subchart.
     */
    void removeDropLines(int subchartIndex);

    /**
     * @brief returns the chart high-low lines.
     *
     * Applicable to chart types: Line, Stock.
     * @param subchartIndex zero-based index of a subchart.
     * @return a copy of the chart hiLowLines.
     */
    ShapeFormat hiLowLines(int subchartIndex) const;
    /**
     * @overload
     * @brief returns the chart high-low lines.
     *
     * Applicable to chart types: Line, Stock
     * @param subchartIndex zero-based index of a subchart.
     * @return a reference to the chart hiLowLines.
     */
    ShapeFormat &hiLowLines(int subchartIndex);
    /**
     * @brief sets the chart high-low lines.
     *
     * Applicable to chart types: Line, Stock.
     * @param subchartIndex zero-based index of a subchart.
     * @param hiLowLines If not valid, sets default hi-low lines.
     * @note to remove hi-low lines use #removeHiLowLines().
     */
    void setHiLowLines(int subchartIndex, const ShapeFormat &hiLowLines);
    /**
     * @brief removes the chart high-low lines.
     * @param subchartIndex zero-based index of a subchart.
     */
    void removeHiLowLines(int subchartIndex);

    /**
     * @brief returns the chart up-dow lines.
     *
     * Applicable to chart types: Line, Stock.
     * @param subchartIndex zero-based index of a subchart.
     * @return a copy of the chart upDownBars.
     */
    UpDownBar upDownBars(int subchartIndex) const;
    /**
     * @brief returns the chart up-down lines.
     *
     * Applicable to chart types: Line, Stock
     * @param subchartIndex zero-based index of a subchart.
     * @return a reference to the chart upDownBars.
     */
    UpDownBar &upDownBars(int subchartIndex);
    /**
     * @brief sets the chart up-down lines.
     *
     * Applicable to chart types: Line, Stock
     * @param subchartIndex zero-based index of a subchart.
     * @param upDownBars UpDownBar object.
     */
    void setUpDownBars(int subchartIndex, const UpDownBar &upDownBars);

    /**
     * @brief if `true`, the chart series marker is shown.
     *
     * Applicable to chart types: Line.
     * @param subchartIndex zero-based index of a subchart.
     * @return valid boolean value if property is set, `nullopt` otherwise.
     */
    std::optional<bool> markerShown(int subchartIndex) const;
    /**
     * @brief sets that the chart series marker shall be shown.
     *
     * Applicable to chart types: Line.
     * @param subchartIndex zero-based index of a subchart.
     * @param markerShown
     */
    void setMarkerShown(int subchartIndex, bool markerShown);

    /**
     * @brief returns smoothing of the chart series. if `true`, the line
     * connecting the points on the chart is smoothed using Catmull-Rom splines.
     *
     * Applicable to chart types: Line.
     * @param subchartIndex zero-based index of a subchart.
     * @return valid boolean value if property is set, `nullopt` otherwise.
     */
    std::optional<bool> smooth(int subchartIndex) const;
    /**
     * @brief sets smoothing of the chart series. if `true`, the line
     * connecting the points on the chart shall be smoothed using Catmull-Rom splines.
     *
     * Applicable to chart types: Line.
     * @param subchartIndex zero-based index of a subchart.
     * @param smooth if `true`, the line
     * connecting the points on the chart is smoothed using Catmull-Rom splines.
     */
    void setSmooth(int subchartIndex, bool smooth);

    /**
     * @brief specifies the space between bar or column clusters, as a
     * percentage of the bar or column width.
     *
     * Applicable to chart types: Line3D, Bar3D, Area3D.
     * @param subchartIndex zero-based index of a subchart.
     * @return valid double ([0..]) value if property is set, `nullopt` otherwise.
     */
    std::optional<int> gapDepth(int subchartIndex) const;
    /**
     * @brief sets the space between bar or column clusters, as a
     * percentage of the bar or column width.
     *
     * Applicable to chart types: Line3D, Bar3D, Area3D.
     *
     * The default value is 100.
     * @param subchartIndex zero-based index of a subchart.
     * @param gapDepth value of the gap depth, in percents (100 equals the bar width).
     */
    void setGapDepth(int subchartIndex, int gapDepth);

    /**
     * @brief returns the style of the scatter chart.
     *
     * The default value is ScatterStyle::Marker.
     * @param subchartIndex zero-based index of a subchart.
     * @return Chart::ScatterStyle
     */
    ScatterStyle scatterStyle(int subchartIndex) const;
    /**
     * @brief sets the style of the scatter chart.
     *
     * If not set, the default value is ScatterStyle::Marker.
     * @param subchartIndex zero-based index of a subchart.
     * @param scatterStyle
     */
    void setScatterStyle(int subchartIndex, ScatterStyle scatterStyle);

    /**
     * @brief returns the style of the radar chart.
     *
     * The default value is RadarStyle::Standard.
     * @param subchartIndex zero-based index of a subchart.
     * @return RadarStyle
     */
    RadarStyle radarStyle(int subchartIndex) const;
    /**
     * @brief sets the style of the radar chart.
     *
     * If not set, the default value is RadarStyle::Standard.
     * @param subchartIndex zero-based index of a subchart.
     * @param radarStyle
     */
    void setRadarStyle(int subchartIndex, RadarStyle radarStyle);

    /**
     * @brief returns whether the series form a bar (horizontal)
     * chart or a column (vertical) chart.
     *
     * The default value is BarDirection::Column.
     *
     * Applicable to chart types: Bar, Bar3D.
     * @param subchartIndex zero-based index of a subchart.
     * @return
     */
    BarDirection barDirection(int subchartIndex) const;
    /**
     * @brief sets whether the series form a bar (horizontal)
     * chart or a column (vertical) chart.
     *
     * If not set, the default value is BarDirection::Column.
     *
     * Applicable to chart types: Bar, Bar3D.
     * @param subchartIndex zero-based index of a subchart.
     * @param barDirection
     */
    void setBarDirection(int subchartIndex, BarDirection barDirection);

    /**
     * @brief returns the space between bar or column clusters, as a
     * percentage of the bar or column width.
     *
     * Applicable to chart types: Bar, Bar3D, OfPie.
     * @param subchartIndex zero-based index of a subchart.
     * @return valid % ([0..500]) value if property is set, `nullopt` otherwise.
     */
    std::optional<int> gapWidth(int subchartIndex) const;
    /**
     * @brief sets the space between bar or column clusters, as a
     * percentage of the bar or column width.
     *
     * Applicable to chart types: Bar, Bar3D, OfPie.
     *
     * The default value is 150.
     * @param subchartIndex zero-based index of a subchart.
     * @param gapWidth gap width in percents (0..500).
     */
    void setGapWidth(int subchartIndex, int gapWidth);

    /**
     * @brief specifies how much bars and columns shall overlap on 2-D charts.
     *
     * Applicable to chart types: Bar.
     * @param subchartIndex zero-based index of a subchart.
     * @return valid int value (percentage) if property is set, `nullopt` otherwise.
     */
    std::optional<int> overlap(int subchartIndex) const;
    /**
     * @brief sets how much bars and columns shall overlap on 2-D charts.
     *
     * If not set, the default value is 0.
     *
     * Applicable to chart types: Bar.
     * @param subchartIndex zero-based index of a subchart.
     * @param overlap overlap percentage in the range [-100..100].
     */
    void setOverlap(int subchartIndex, int overlap);

    /**
     * @brief seriesLines returns the series lines of the chart.
     *
     * Applicable to chart types: Bar, OfPie.
     * @param subchartIndex zero-based index of a subchart.
     * @return A list of the series lines formats.
     */
    QList<ShapeFormat> seriesLines(int subchartIndex) const;
    /**
     * @brief seriesLines returns the series lines of the chart.
     *
     * Applicable to chart types: Bar, OfPie.
     * @param subchartIndex zero-based index of a subchart.
     * @return a reference to the series lines formats.
     */
    QList<ShapeFormat> &seriesLines(int subchartIndex);
    /**
     * @brief setSeriesLines sets the series lines of the chart.
     *
     * Applicable to chart types: Bar, OfPie.
     * @param subchartIndex zero-based index of a subchart.
     * @param seriesLines
     */
    void setSeriesLines(int subchartIndex, const QList<ShapeFormat> &seriesLines);

    /**
     * @brief barShape returns the shape of a series or a 3-D bar chart.
     *
     * Applicable to chart types: Bar3D.
     * @param subchartIndex zero-based index of a subchart.
     * @return valid Series::BarShape value or `nullopt` if the property is not set.
     */
    std::optional<Series::BarShape> barShape(int subchartIndex) const;
    /**
     * @brief setBarShape sets the shape of a series or a 3-D bar chart.
     *
     * Applicable to chart types: Bar3D.
     * @param subchartIndex zero-based index of a subchart.
     * @param barShape Series::BarShape
     */
    void setBarShape(int subchartIndex, Series::BarShape barShape);

    /**
     * @brief firstSliceAngle returns the angle of the first pie or doughnut
     * chart slice, in degrees (clockwise from up).
     *
     * Applicable to chart types: Pie, Doughnut.
     * @param subchartIndex zero-based index of a subchart.
     * @return valid int value if the property is set, `nullopt` otherwise.
     */
    std::optional<int> firstSliceAngle(int subchartIndex) const;
    /**
     * @brief setFirstSliceAngle sets  the angle of the first pie or doughnut
     * chart slice, in degrees (clockwise from up).
     *
     * Applicable to chart types: Pie, Doughnut.
     * @param subchartIndex zero-based index of a subchart.
     * @param angle value in degrees in the range [0..360].
     */
    void setFirstSliceAngle(int subchartIndex, int angle);

    /**
     * @brief returns the size, in percents, of the hole in a doughnut chart.
     *
     * Applicable to chart types: Doughnut.
     *
     * @return valid int value if the property is set (in the range [1..90]), `nullopt` otherwise.
     */
    std::optional<int> holeSize(int subchartIndex) const;
    /**
     * @brief sets the size, in percents, of the hole in a doughnut chart.
     *
     * Applicable to chart types: Doughnut.
     *
     * The default value is 10.
     * @param subchartIndex zero-based index of a subchart.
     * @param holeSize the hole size in percents in the range [1..90]
     */
    void setHoleSise(int subchartIndex, int holeSize);

    //+ofpie
    /**
     * @brief returns the OfPie chart type.
     *
     * The default value is OfPieType::Pie.
     * @param subchartIndex zero-based index of a subchart.
     * @return OfPieType::Bar or OfPieType::Pie enum value.
     */
    Chart::OfPieType ofPieType(int subchartIndex) const;
    /**
     * @brief sets the OfPie chart type.
     * @param subchartIndex zero-based index of a subchart.
     * @param type OfPieType::Bar or OfPieType::Pie enum value.
     */
    void setOfPieType(int subchartIndex, OfPieType type);

    /**
     * @brief returns the OfPie chart split type (how to split the data points
     * between the first pie and second pie or bar.)
     *
     * The default value is SplitType::Auto.
     * @param subchartIndex zero-based index of a subchart.
     */
    std::optional<SplitType> splitType(int subchartIndex) const;
    /**
     * @brief sets the OfPie chart split type (how to split the data points
     * between the first pie and second pie or bar.)
     * @param subchartIndex zero-based index of a subchart.
     * @param splitType SplitType enum.
     */
    void setSplitType(int subchartIndex, SplitType splitType);
    /**
     * @brief returns a value used to determine which data points are in the
     * second pie on an OfPie chart.
     * @param subchartIndex zero-based index of a subchart.
     * @return If splitType() is Position, splitPos equals to the number of last data points.
     * If Percent, splitPos is the percentage of the data points sum, and data points
     * with value less than splitPos go to the second pie.
     * If Value, splitPos is the value below which data points go to the second pie.
     */
    std::optional<double> splitPos(int subchartIndex) const;
    /**
     * @brief sets the value that shall be used to determine which data points are in the
     * second pie on an OfPie chart.
     * @param subchartIndex zero-based index of a subchart.
     * @param value If splitType() is Position, value equals to the number of last data points.
     * If Percent, value is the percentage of the data points sum, and data points
     * with value less than that go to the second pie.
     * If Value, value is the value below which data points go to the second pie.
     */
    void setSplitPos(int subchartIndex, double value);
    /**
     * @brief returns the list of data points indexes (starting from 1) that go to
     * the second pie on an OfPie chart.
     *
     * This parameter is valid only if splitType() is SplitType::Custom.
     * @param subchartIndex zero-based index of a subchart.
     */
    QList<int> customSplit(int subchartIndex) const;
    /**
     * @brief sets the list of data points indexes (starting from 1) that go to
     * the second pie on an OfPie chart.
     * @param subchartIndex zero-based index of a subchart.
     * @param indexes the list of valid data points indexes (starting from 1).
     */
    void setCustomSplit(int subchartIndex, const QList<int> &indexes);
    /**
     * @brief returns the second pie size of an OfPie chart, as a percentage
     * of the size of the first pie.
     * @param subchartIndex zero-based index of a subchart.
     * @return valid percentage value (5..200%) or `nullopt` if the parameter is not set.
     */
    std::optional<int> secondPieSize(int subchartIndex) const;
    /**
     * @brief sets the second pie size of an OfPie chart, as a percentage
     * of the size of the first pie.
     *
     * If not set, the default value is 75.
     * @param subchartIndex zero-based index of a subchart.
     * @param size percentage value (5..200).
     */
    void setSecondPieSize(int subchartIndex, int size);

    /**
     * @brief returns whether the Bubble chart has a 3-D effect applied to series.
     * @param subchartIndex zero-based index of a subchart.
     * @return valid optional value if bubble3D property is set, `nullopt` otherwise.
     */
    std::optional<bool> bubble3D(int subchartIndex) const;
    /**
     * @brief sets whether the Bubble chart has a 3-D effect applied to series.
     * @param subchartIndex zero-based index of a subchart.
     * @param bubble3D
     *
     * Applicable to char types: Bubble.
     */
    void setBubble3D(int subchartIndex, bool bubble3D);
    /**
     * @brief returns whether negative sized bubbles shall be shown on a bubble chart.
     * @param subchartIndex zero-based index of a subchart.
     * @return valid optional value if bubble3D property is set, `nullopt` otherwise.
     *
     * If the parameter is not set, the default value is `true`.
     *
     * Applicable to char types: Bubble.
     */
    std::optional<bool> showNegativeBubbles(int subchartIndex) const;
    /**
     * @brief sets whether negative sized bubbles shall be shown on a bubble chart.
     * @param subchartIndex zero-based index of a subchart.
     * @param show
     *
     * The default value is `true`.
     *
     * Applicable to char types: Bubble.
     */
    void setShowNegativeBubbles(int subchartIndex, bool show);

    /**
     * @brief returns the scale factor for a bubble chart.
     * @param subchartIndex zero-based index of a subchart.
     * @return a percentage value from 0 to 300, corresponding to a percentage of the default size.
     *
     * Applicable to char types: Bubble.
     */
    std::optional<int> bubbleScale(int subchartIndex) const;
    /**
     * @brief sets the scale factor for a bubble chart.
     * @param subchartIndex zero-based index of a subchart.
     * @param scale a percentage value from 0 to 300, corresponding to a percentage of the default size.
     *
     * Applicable to char types: Bubble.
     */
    void setBubbleScale(int subchartIndex, int scale);

    /**
     * @brief returns how the bubble size values are represented on the chart.
     * @param subchartIndex zero-based index of a subchart.
     * @return valid optional value if parameter is set, `nullopt` otherwise.
     *
     * Applicable to char types: Bubble.
     */
    std::optional<BubbleSizeRepresents> bubbleSizeRepresents(int subchartIndex) const;
    /**
     * @brief sets how the bubble size values are represented on the bubble chart.
     * @param subchartIndex zero-based index of a subchart.
     * @param value
     *
     * Applicable to char types: Bubble.
     */
    void setBubbleSizeRepresents(int subchartIndex, BubbleSizeRepresents value);

    /**
     * @brief returns whether the surface chart is drawn as a wireframe.
     * @param subchartIndex zero-based index of a subchart.
     * @return valid optional value if the parameter is set, `nullopt` otherwise.
     *
     * The default value is `true`.
     *
     * Applicable to chart types: Surface, Surface3D.
     */
    std::optional<bool> wireframe(int subchartIndex) const;
    /**
     * @brief sets whether the surface chart is drawn as a wireframe.
     * @param subchartIndex zero-based index of a subchart.
     * @param wireframe
     *
     * The default value is `true`.
     *
     * Applicable to chart types: Surface, Surface3D
     */
    void setWireframe(int subchartIndex, bool wireframe);

    /**
     * @brief returns a map of formatting bands for a surface chart indexed from low to high.
     * @param subchartIndex zero-based index of a subchart.
     * @return
     *
     * Applicable to chart types: Surface, Surface3D.
     */
    QMap<int, ShapeFormat> bandFormats(int subchartIndex) const;
    /**
     * @brief sets a map of formatting bands for a surface chart.
     * @param subchartIndex zero-based index of a subchart.
     * @param bandFormats
     *
     * Applicable to chart types: Surface, Surface3D.
     */
    void setBandFormats(int subchartIndex, QMap<int, ShapeFormat> bandFormats);

    bool loadFromXmlFile(QIODevice *device) override;
    void saveToXmlFile(QIODevice *device) const override;
    void saveMediaFiles(Workbook *workbook);
    void loadMediaFiles(Workbook *workbook);

private:
    Series* seriesByOrder(int order);
    bool hasAxis(int id) const;
    friend class SubChart;
};


}

#endif // QXLSX_CHART_H
