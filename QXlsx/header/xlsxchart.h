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
        OfPie, /**< @brief Pie of pie chart (or bar of pie based on the ofPieChart parameter.) */
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
    /*!
     * Destroys the chart.
     */
    ~Chart();
public:
    /**
     * @brief sets type to the chart.
     * @warning This method does not change the previous subcharts type nor the
     * previous series type.
     * @note OOXML schema allows having as many subcharts of different types
     * as you need. Each subchart is used to group series of the same type.
     * If you need to add both bar series and line series to the chart, you can do it this way:
     * @code
     * //set Type::Bar to the chart
     * chart->setType(Chart::Type::Bar);
     * //add bar series
     * chart.addSeries(QXlsx::CellRange("A1:A10"), QXlsx::CellRange("B1:B10"),
     *     xlsx.sheet("Sheet1"), true);
     * //set Type::Line to the chart
     * chart->setType(Chart::Type::Line);
     * //add line series
     * chart->addSeries(QXlsx::CellRange("A1:A10"), QXlsx::CellRange("C1:C10"),
     *     xlsx.sheet("Sheet1"), true);
     * @endcode
     * @note All parameters of the chart are applied only to the last subchart that was
     * added with #setType() method. Right now there is no possibility to change
     * parameters of previous subcharts.
     * @param type Chart type. If set to Type::None the chart is ill-formed.
     */
    void setType(Type type);
    /**
     * @brief returns the chart type.
     * @return the type
     */
    Type type() const;
    /**
     * @brief sets line format to the chart space (the outer area of the chart).
     *
     * This is a convenience method, equivalent to ```chartShape()->setLineFormat(format)```
     * @param format
     */
    void setChartLineFormat(const LineFormat &format);
    /**
     * @brief sets line format to the plot area (the inner area of the chart
     * where actual plotting occurs).
     *
     * This is a convenience method, equivalent to ```plotAreaShape()->setLineFormat(format)```
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
     * @return pointer to the added series
     */
    Series *addSeries();

    /**
     * @overload
     * @brief adds one or more series with data defined by #range.
     * @param range valid CellRange
     * @param sheet data source for range
     * @param firstRowContainsHeaders specifies that the 1st row (the 1st column
     * if columnBased is false) contains the series titles and should be
     * excluded from the series data.
     * @param firstColumnContainsCategoryData specifies that the 1st column
     * (the 1st row if columnBased is false) contains data for the category (x)
     * axis.
     * @param columnBased specifies that range should be treated as column-based
     * or row-based: if columnBased is true, the first column is category data
     * (if firstColumnContainsCategoryData is set to true),
     * other columns are value data for new series. If columnBased is false, the
     * first row is category data, other rows are value data for new series.
     * @note the series will have the type specified by the most recently set chart type.
     * So it is possible to have series of different types on the same chart.
     */
    void addSeries(const CellRange &range, AbstractSheet *sheet = NULL,
                   bool firstRowContainsHeaders = false,
                   bool firstColumnContainsCategoryData = false,
                   bool columnBased = true);
    /**
     * @overload
     * @brief adds series defined by keyRange and valRange.
     * @param keyRange category range or x axis range.
     * @param valRange value range or y axis range.
     * @param sheet data sheet reference.
     * @param keyRangeIncludesHeader if true, the first row or column is used as a series name reference.
     * @return pointer to the added series or nullptr if no series was added.
     * @note The series will have the type specifies by the most recently set chart type.
     * So it is possible to have series of different types on the same chart.
     *
     * @note To set bubbleSizeDataSource to the BubbleChart series (the 3rd range)
     * use #addSeries() method and manually set all three ranges.
     */
    Series* addSeries(const CellRange &keyRange, const CellRange &valRange,
                       AbstractSheet *sheet = NULL, bool keyRangeIncludesHeader = false);
    /**
     * @brief returns the chart series by its index.
     * @param index the series index starting from 0.
     * @return pointer to the series if such series exists, nullptr otherwise.
     */
    Series *series(int index);

    /**
     * @brief returns the chart's series count.
     * @return
     */
    int seriesCount() const;

    /**
     * @brief removes series with #index from the chart.
     * @param index series index starting from 0.
     * @return true if such a series was found and successfully deleted, false otherwise.
     */
    bool removeSeries(int index);
    /**
     * @overload
     * @brief removes #series from the chart.
     * @param series
     * @return true if such a series was found and successfully deleted, false otherwise.
     */
    bool removeSeries(const Series &series);
    /**
     * @brief removes all series from the chart.
     * @note This method will also remove all subcharts and create a new subchart of type
     * equal to the last set chart type. This method does not remove chart axes.
     */
    void removeAllSeries();

    /**
     * @brief sets axes for all series on the chart.
     *
     * The method doesn't check for the axes availability and sets axesIds for all series
     * added to the chart. To set specific axes for some series (f.e. to move the series to the
     * right axis) use setSeriesAxesIDs(Series *series, const QList<int> &axesIds) method.
     *
     * @param axesIds a list of 0, 2 or 3 items.
     * @note axesIds are ids, not indexes. \see Axis::id()
     */
    void setSeriesAxesIDs(const QList<int> &axesIds);
    /**
     * @overload
     * @brief sets axes for the specific series
     * @param series pointer to the series
     * @param axesIds a list of 0, 2 or 3 items.
     * @note axesIds are ids, not indexes. \see Axis::id()
     *
     * @note Each series (except for Pie) shall be attached to at least two axes.
     * If no axes are specified, then the file is ill-formed. The best way to set
     * axes is as follows:
     *
     * @code
     * //first add axes to the chart
     * auto bottom = chart.addAxis(QXlsx::Axis::Type::Val, QXlsx::Axis::Position::Bottom);
     * auto left = chart.addAxis(QXlsx::Axis::Type::Val, QXlsx::Axis::Position::Left);
     * left->setCrossAxis(bottom); //also sets bottom to cross left
     * QList<int> axesIds {b->id(), l->id()};
     *
     * //or create default axes:
     * QList<int> axesIds = chart.addDefaultAxes();
     *
     * //then add series and specify its axes
     * auto series = chart.addSeries(QXlsx::CellRange("A1:A10"), QXlsx::CellRange("B1:B10"),
     * xlsx.sheet("Sheet1"), true);
     * chart.setSeriesAxesIDs(series, axesIds);
     *
     * @endcode
     */
    void setSeriesAxesIDs(Series *series, const QList<int> &axesIds);

    /**
     * @brief sets all available axes for all series added to the chart.
     *
     * The method doesn't check for the axes availability and sets _all_ available
     * axes for _all_ series added to the chart. To set specific axes for some
     * series (f.e. to move the series to the right axis) use
     * setSeriesAxesIDs() method.
     *
     * @note invoke this method after you have added all the series and all the
     * axes if all the series share the same axes.
     * @sa setSeriesAxesIDs
     */
    void setSeriesDefaultAxes();
    /**
     * @brief changes the series order in which it is drawn on the chart.
     * @param oldOrder
     * @param newOrder
     */
    void moveSeries(int oldOrder, int newOrder);

    /**
     * @brief adds all necessary axes to the chart.
     *
     * You can use this method to create all default axes. If you want to later fine-tune them,
     * use axis(int idx) method.
     *
     * @warning A chart without the required axes is considered ill-formed. Here is the table of
     * axes requirements for each chart type
     *
     * Chart type | axes count and default axes
     * -----------|---------------------
     * Line, Bar, Area, Radar, Stock, Bubble | 2: Category (bottom) and Value (left)
     * Line3D, Surface3D | 3: Category (bottom), Value (left) and Series (bottom)
     * Bar3D, Area3D, Surface | 2-3: Category (bottom), Value (left) and Series (bottom)
     * Scatter | 2: Value (bottom) and Value (left)
     * Pie, Pie3D, Doughnut, OfPie | 0 (no axes required)
     *
     * @note this method does not delete axes already present in the chart.
     *
     * @return a list of axes ids, that can later be used to set them for series
     */
    QList<int> addDefaultAxes();
    /**
     * @brief adds new axis with specified parameters.
     * @param type axis type: Cat, Val, Date or Ser.
     * @param pos Axis position: Bottom, Left, Top or Right.
     * @param title optional axis title.
     * @return reference to the newly added axis.
     */
    Axis &addAxis(Axis::Type type, Axis::Position pos, QString title = QString());
    /**
     * @brief adds new axis created elsewhere.
     *
     * The method can be used to duplicate existing axis, change the type and position of the copy
     * and add the result to the chart. The method makes sure new axis will have a unique ID.
     * @param axis
     */
    void addAxis(const Axis &axis);
    /**
     * @brief returns axis that has index idx
     * @param idx valid index (0 <= idx < axesCount()).
     * @return pointer to the axis if such axis exists, nullptr otherwise.
     *
     * This is the same as ```&axes().at(idx)```
     */
    Axis *axis(int idx); //TODO: replace with std::optional<std::reference_wrapper<Axis>> axis(int idx);
    /**
     * @overload
     * @brief returns axis that has position #pos.
     * @param pos Axis::Position.
     * @note A chart can have several axes positioned at #pos. This method returns
     * _the first added_ axis that has position #pos.
     * @return pointer to the axis if such axis exists, nullptr otherwise.
     */
    Axis *axis(Axis::Position pos);
    /**
     * @overload
     * @brief returns an axis that has type #type.
     * @param type Axis::Type.
     * @note A chart can have several axes of the same type (f.e. scatter chart
     * usually has 2 value axes). This method returns _the first added_ axis of the
     * type.
     * @return pointer to the axis if such axis exists, nullptr otherwise.
     */
    Axis *axis(Axis::Type type);
    /**
     * @overload
     * @brief returns an axis that has position #pos and type #type.
     * @param pos Axis::Position.
     * @param type Axis::Type.
     * @note A chart can have several axes of the same position and type. This
     * method returns _the first added_ axis of these position and type.
     * @return pointer to the axis if such axis exists, nullptr otherwise.
     */
    Axis *axis(Axis::Type type, Axis::Position pos);
    /**
     * @brief returns the list of axes defined in the chart.
     * @return a list of axes. As Axis is an explicitly shareable class, no actual copying
     * of axes occurs. @see Axis.
     */
    QList<QXlsx::Axis> axes() const;

    /**
     * @brief tries to remove axis and returns true if axis has been removed.
     * @param axisID the id of the axis to be removed.
     * @return true if axis has been removed, false otherwise (no axis with this id or axis is used
     * in some series).
     */
    bool removeAxis(int axisID);
    /**
     * @overload
     * @brief tries to remove axis and returns true if axis has been removed.
     * @param axis the pointer to the axis to be removed.
     * @return true if axis has been removed, false otherwise (no such axis or axis is used
     * in some series).
     */
    bool removeAxis(Axis *axis);
    /**
     * @brief removes all axes defined in the chart.
     * This method removes all axes defined in the chart and clears the series references to the axes.
     * After that until you add the required axes and set these axes for the chart series
     * the chart is ill-formed.
     */
    void removeAxes();
    /**
     * @brief returns the list of series that use axis with #axisID.
     * @param axisID axis ID (_not axis index!_). @see Axis::id().
     * @return
     */
    QList<Series> seriesThatUseAxis(int axisID) const;
    /**
     * @overload
     * @brief returns the list of series that use #axis.
     * @param axis pointer to the axis.
     * @return
     */
    QList<Series> seriesThatUseAxis(Axis *axis) const;
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
     * @param value if true then the automatically generated title is not shown.
     * @note If the chart has only one series and autoTitleDeleted is set to false,
     * then the chart will have the title that contains the series name. Set
     * autoTitleDeleted to true to hide this title.
     */
    void setAutoTitleDeleted(bool value);

    /**
     * @brief return whether only visible cells should be plotted on the chart.
     * @return
     *
     * If not set, the default value is true.
     */
    std::optional<bool> plotOnlyVisibleCells() const;
    /**
     * @brief sets whether only visible cells should be plotted on the chart.
     * @param plot
     *
     * If not set, the default value is true.
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
     * If not set, the default value is true.
     *
     * @return
     */
    std::optional<bool> showDataLabelsOverMaximum() const;
    /**
     * @brief sets whether data labels over the maximum of the chart shall be shown.
     * @param value
     *
     * If not set, the default value is true.
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
     * @brief returns the kind of grouping for a line or area chart.
     *
     * Applicable to: Area, Area3D, Line, Line3D.
     *
     * @return valid Chart::Grouping if property is set, nullopt otherwise.
     */
    std::optional<Chart::Grouping> grouping() const;
    /**
     * @brief sets the kind of grouping for a line or area chart.
     *
     * Applicable to: Area, Area3D, Line, Line3D.
     *
     * @param grouping Chart::Grouping
     */
    void setGrouping(Chart::Grouping grouping);

    /**
     * @brief returns the kind of grouping for a bar chart.
     *
     * Applicable to: Bar, Bar3D.
     *
     * @return valid Chart::BarGrouping if property is set, nullopt otherwise.
     */
    std::optional<Chart::BarGrouping> barGrouping() const;
    /**
     * @brief sets the kind of grouping for a bar chart.
     *
     * Applicable to chart types: Bar, Bar3D.
     *
     * @param grouping Chart::BarGrouping
     */
    void setBarGrouping(Chart::BarGrouping grouping);

    /**
     * @brief specifies that each data marker in the series has a different color.
     *
     * Applicable to chart types: Line, Line3D, Scatter, Radar, Bar, Bar3D, Area, Area3D,
     * Pie, Pie3D, Doughnut, OfPie, Bubble.
     *
     * @return valid bool value if property is set, nullopt otherwise.
     *
     * @note by default varyColors is not set. But in order to get sensible charts
     * all new Pie, Pie3D, Doughnut, OfPie, Bubble charts automatically set varyColors to true.
     *
     */
    std::optional<bool> varyColors() const;
    /**
     * @brief sets that each data marker in the series has a different color.
     *
     * Applicable to chart types: Line, Line3D, Scatter, Radar, Bar, Bar3D, Area, Area3D,
     * Pie, Pie3D, Doughnut, OfPie, Bubble.
     *
     * @param varyColors true if each data marker in the series has a different color, false
     * if every data markers in the series have the same color.
     */
    void setVaryColors(bool varyColors);

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
     *
     * @return a copy of the chart labels.
     *
     */
    Labels labels() const;
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
     *
     * @return reference to the chart labels.
     */
    Labels &labels();
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
     *
     * @param labels
     */
    void setLabels(const Labels &labels);

    /**
     * @brief returns the chart drop lines.
     *
     * Applicable to chart types: Line, Line3D, Stock, Area, Area3D.
     *
     * @return a copy of the chart dropLines.
     */
    ShapeFormat dropLines() const;
    /**
     * @brief returns the chart drop lines.
     *
     * Applicable to chart types: Line, Line3D, Stock, Area, Area3D.
     *
     * @return a reference to the chart dropLines.
     */
    ShapeFormat &dropLines();
    /**
     * @brief sets the chart drop lines.
     *
     * Applicable to chart types: Line, Line3D, Stock, Area, Area3D.
     *
     * @param dropLines
     */
    void setDropLines(const ShapeFormat &dropLines);

    /**
     * @brief returns the chart high-low lines.
     *
     * Applicable to chart types: Line, Stock.
     *
     * @return a copy of the chart hiLowLines.
     */
    ShapeFormat hiLowLines() const;
    /**
     * @brief returns the chart high-low lines.
     *
     * Applicable to chart types: Line, Stock
     *
     * @return a reference to the chart hiLowLines.
     */
    ShapeFormat &hiLowLines();
    /**
     * @brief sets the chart high-low lines.
     *
     * Applicable to chart types: Line, Stock
     *
     * @param hiLowLines
     */
    void setHiLowLines(const ShapeFormat &hiLowLines);

    /**
     * @brief returns the chart up-dow lines.
     *
     * Applicable to chart types: Line, Stock.
     *
     * @return a copy of the chart upDownBars.
     */
    UpDownBar upDownBars() const;
    /**
     * @brief returns the chart up-down lines.
     *
     * Applicable to chart types: Line, Stock
     *
     * @return a reference to the chart upDownBars.
     */
    UpDownBar &upDownBars();
    /**
     * @brief sets the chart up-down lines.
     *
     * Applicable to chart types: Line, Stock
     *
     * @param upDownBars
     */
    void setUpDownBars(const UpDownBar &upDownBars);

    /**
     * @brief if true, the chart series marker is shown.
     *
     * Applicable to chart types: Line.
     *
     * @return valid boolean value if property is set, nullopt otherwise.
     */
    std::optional<bool> markerShown() const;
    /**
     * @brief sets that the chart series marker shall be shown.
     *
     * Applicable to chart types: Line.
     *
     * @param markerShown
     */
    void setMarkerShown(bool markerShown);

    /**
     * @brief returns smoothing of the chart series. If true, the line
     * connecting the points on the chart is smoothed using Catmull-Rom splines.
     *
     * Applicable to chart types: Line.
     *
     * @return valid boolean value if property is set, nullopt otherwise.
     */
    std::optional<bool> smooth() const;
    /**
     * @brief sets smoothing of the chart series. If true, the line
     * connecting the points on the chart shall be smoothed using Catmull-Rom splines.
     *
     * Applicable to chart types: Line.
     *
     * @param smooth if true, the line
     * connecting the points on the chart is smoothed using Catmull-Rom splines.
     */
    void setSmooth(bool smooth);

    /**
     * @brief specifies the space between bar or column clusters, as a
     * percentage of the bar or column width.
     *
     * Applicable to chart types: Line3D, Bar3D, Area3D.
     *
     * @return valid double ([0..]) value if property is set, nullopt otherwise.
     */
    std::optional<int> gapDepth() const;
    /**
     * @brief sets the space between bar or column clusters, as a
     * percentage of the bar or column width.
     *
     * Applicable to chart types: Line3D, Bar3D, Area3D.
     *
     * The default value is 100.
     *
     * @param gapDepth value of the gap depth, in percents (100 equals the bar width).
     */
    void setGapDepth(int gapDepth);

    /**
     * @brief returns the style of the scatter chart.
     *
     * The default value is ScatterStyle::Marker.
     *
     * @return Chart::ScatterStyle
     */
    ScatterStyle scatterStyle() const;
    /**
     * @brief sets the style of the scatter chart.
     *
     * If not set, the default value is ScatterStyle::Marker.
     *
     * @param scatterStyle
     */
    void setScatterStyle(ScatterStyle scatterStyle);

    /**
     * @brief returns the style of the radar chart.
     *
     * The default value is RadarStyle::Standard.
     *
     * @return RadarStyle
     */
    RadarStyle radarStyle() const;
    /**
     * @brief sets the style of the radar chart.
     *
     * If not set, the default value is RadarStyle::Standard.
     *
     * @param radarStyle
     */
    void setRadarStyle(RadarStyle radarStyle);

    /**
     * @brief returns whether the series form a bar (horizontal)
     * chart or a column (vertical) chart.
     *
     * The default value is BarDirection::Column.
     *
     * Applicable to chart types: Bar, Bar3D.
     *
     * @return
     */
    BarDirection barDirection() const;
    /**
     * @brief sets whether the series form a bar (horizontal)
     * chart or a column (vertical) chart.
     *
     * If not set, the default value is BarDirection::Column.
     *
     * Applicable to chart types: Bar, Bar3D.
     *
     * @param barDirection
     */
    void setBarDirection(BarDirection barDirection);

    /**
     * @brief returns the space between bar or column clusters, as a
     * percentage of the bar or column width.
     *
     * Applicable to chart types: Bar, Bar3D, OfPie.
     *
     * @return valid % ([0..500]) value if property is set, nullopt otherwise.
     */
    std::optional<int> gapWidth() const;
    /**
     * @brief sets the space between bar or column clusters, as a
     * percentage of the bar or column width.
     *
     * Applicable to chart types: Bar, Bar3D, OfPie.
     *
     * The default value is 150.
     *
     * @param gapWidth gap width in percents (0..500).
     */
    void setGapWidth(int gapWidth);

    /**
     * @brief specifies how much bars and columns shall overlap on 2-D charts.
     *
     * Applicable to chart types: Bar.
     *
     * @return valid int value (percentage) if property is set, nullopt otherwise.
     */
    std::optional<int> overlap() const;
    /**
     * @brief sets how much bars and columns shall overlap on 2-D charts.
     *
     * If not set, the default value is 0.
     *
     * Applicable to chart types: Bar.
     *
     * @param overlap overlap percentage in the range [-100..100].
     */
    void setOverlap(int overlap);

    /**
     * @brief seriesLines returns the series lines of the chart.
     *
     * Applicable to chart types: Bar, OfPie.
     *
     * @return A list of the series lines formats.
     */
    QList<ShapeFormat> seriesLines() const;
    /**
     * @brief seriesLines returns the series lines of the chart.
     *
     * Applicable to chart types: Bar, OfPie.
     *
     * @return a reference to the series lines formats.
     */
    QList<ShapeFormat> &seriesLines();
    /**
     * @brief setSeriesLines sets the series lines of the chart.
     *
     * Applicable to chart types: Bar, OfPie.
     *
     * @param seriesLines
     */
    void setSeriesLines(const QList<ShapeFormat> &seriesLines);

    /**
     * @brief barShape returns the shape of a series or a 3-D bar chart.
     *
     * Applicable to chart types: Bar3D.
     *
     * @return valid Series::BarShape value or nullopt if the property is not set.
     */
    std::optional<Series::BarShape> barShape() const;
    /**
     * @brief setBarShape sets the shape of a series or a 3-D bar chart.
     *
     * Applicable to chart types: Bar3D.
     *
     * @param barShape Series::BarShape
     */
    void setBarShape(Series::BarShape barShape);

    /**
     * @brief firstSliceAngle returns the angle of the first pie or doughnut
     * chart slice, in degrees (clockwise from up).
     *
     * Applicable to chart types: Pie, Doughnut.
     *
     * @return valid int value if the property is set, nullopt otherwise.
     */
    std::optional<int> firstSliceAngle() const;
    /**
     * @brief setFirstSliceAngle sets  the angle of the first pie or doughnut
     * chart slice, in degrees (clockwise from up).
     *
     * Applicable to chart types: Pie, Doughnut.
     *
     * @param angle value in degrees in the range [0..360].
     */
    void setFirstSliceAngle(int angle);

    /**
     * @brief returns the size, in percents, of the hole in a doughnut chart.
     *
     * Applicable to chart types: Doughnut.
     *
     * @return valid int value if the property is set (in the range [1..90]), nullopt otherwise.
     */
    std::optional<int> holeSize() const;
    /**
     * @brief sets the size, in percents, of the hole in a doughnut chart.
     *
     * Applicable to chart types: Doughnut.
     *
     * The default value is 10.
     *
     * @param holeSize the hole size in percents in the range [1..90]
     */
    void setHoleSise(int holeSize);

    //+ofpie
    /**
     * @brief returns the OfPie chart type.
     *
     * The default value is OfPieType::Pie.
     *
     * @return OfPieType::Bar or OfPieType::Pie enum value.
     */
    Chart::OfPieType ofPieType() const;
    /**
     * @brief sets the OfPie chart type.
     * @param type OfPieType::Bar or OfPieType::Pie enum value.
     */
    void setOfPieType(OfPieType type);

    /**
     * @brief returns the OfPie chart split type (how to split the data points
     * between the first pie and second pie or bar.)
     *
     * The default value is SplitType::Auto
     */
    std::optional<SplitType> splitType() const;
    /**
     * @brief sets the OfPie chart split type (how to split the data points
     * between the first pie and second pie or bar.)
     * @param splitType SplitType enum.
     */
    void setSplitType(SplitType splitType);
    /**
     * @brief returns a value used to determine which data points are in the
     * second pie on an OfPie chart.
     *
     * If splitType() is Position, splitPos equals to the number of last data points.
     * If Percent, splitPos is the percentage of the data points sum, and data points
     * with value less than splitPos go to the second pie.
     * If Value, splitPos is the value below which data points go to the second pie.
     */
    std::optional<double> splitPos() const;
    /**
     * @brief sets the value that shall be used to determine which data points are in the
     * second pie on an OfPie chart.
     * @param value If splitType() is Position, value equals to the number of last data points.
     * If Percent, value is the percentage of the data points sum, and data points
     * with value less than that go to the second pie.
     * If Value, value is the value below which data points go to the second pie.
     *
     * Applicable to chart types: OfPie.
     */
    void setSplitPos(double value);
    /**
     * @brief returns the list of data points indexes (starting from 1) that go to
     * the second pie on an OfPie chart.
     *
     * This parameter is valid only if splitType() is SplitType::Custom.
     */
    QList<int> customSplit() const;
    /**
     * @brief sets the list of data points indexes (starting from 1) that go to
     * the second pie on an OfPie chart.
     * @param indexes the list of valid data points indexes (starting from 1).
     */
    void setCustomSplit(const QList<int> &indexes);
    /**
     * @brief returns the second pie size of an OfPie chart, as a percentage
     * of the size of the first pie.
     * @return valid percentage value (5..200%) or nullopt if the parameter is not set.
     */
    std::optional<int> secondPieSize() const;
    /**
     * @brief sets the second pie size of an OfPie chart, as a percentage
     * of the size of the first pie.
     *
     * If not set, the default value is 75.
     *
     * @param size percentage value (5..200).
     */
    void setSecondPieSize(int size);

    /**
     * @brief returns whether the Bubble chart has a 3-D effect applied to series.
     * @return valid optional value if bubble3D property is set, nullopt otherwise.
     *
     * Applicable to char types: Bubble.
     */
    std::optional<bool> bubble3D() const;
    /**
     * @brief sets whether the Bubble chart has a 3-D effect applied to series.
     * @param bubble3D
     *
     * Applicable to char types: Bubble.
     */
    void setBubble3D(bool bubble3D);
    /**
     * @brief returns whether negative sized bubbles shall be shown on a bubble chart.
     * @return valid optional value if bubble3D property is set, nullopt otherwise.
     *
     * If the parameter is not set, the default value is true.
     *
     * Applicable to char types: Bubble.
     */
    std::optional<bool> showNegativeBubbles() const;
    /**
     * @brief sets whether negative sized bubbles shall be shown on a bubble chart.
     * @param show
     *
     * The default value is true.
     *
     * Applicable to char types: Bubble.
     */
    void setShowNegativeBubbles(bool show);

    /**
     * @brief returns the scale factor for a bubble chart.
     * @return a percentage value from 0 to 300, corresponding to a percentage of the default size.
     *
     * Applicable to char types: Bubble.
     */
    std::optional<int> bubbleScale() const;
    /**
     * @brief sets the scale factor for a bubble chart.
     * @param scale a percentage value from 0 to 300, corresponding to a percentage of the default size.
     *
     * Applicable to char types: Bubble.
     */
    void setBubbleScale(int scale);

    /**
     * @brief returns how the bubble size values are represented on the chart.
     * @return valid optional value if parameter is set, nullopt otherwise.
     *
     * Applicable to char types: Bubble.
     */
    std::optional<BubbleSizeRepresents> bubbleSizeRepresents() const;
    /**
     * @brief sets how the bubble size values are represented on the bubble chart.
     * @param value
     *
     * Applicable to char types: Bubble.
     */
    void setBubbleSizeRepresents(BubbleSizeRepresents value);

    /**
     * @brief returns whether the surface chart is drawn as a wireframe.
     * @return valid optional value if the parameter is set, nullopt otherwise.
     *
     * The default value is true.
     *
     * Applicable to chart types: Surface, Surface3D.
     */
    std::optional<bool> wireframe() const;
    /**
     * @brief sets whether the surface chart is drawn as a wireframe.
     * @param wireframe
     *
     * The default value is true.
     *
     * Applicable to chart types: Surface, Surface3D
     */
    void setWireframe(bool wireframe);

    /**
     * @brief returns a map of formatting bands for a surface chart indexed from low to high.
     * @return
     *
     * Applicable to chart types: Surface, Surface3D.
     */
    QMap<int, ShapeFormat> bandFormats() const;
    /**
     * @brief sets a map of formatting bands for a surface chart.
     * @param bandFormats
     *
     * Applicable to chart types: Surface, Surface3D.
     */
    void setBandFormats(QMap<int, ShapeFormat> bandFormats);

    /**
     * @brief sets the visibility of the chart data table
     * @param visible if true, the chart data table is shown below the chart. If false,
     * the chart data table is hidden.
     *
     * @warning Setting the data table visibility to false deletes the current data table properties.
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
    void setStyleID(int id);

    /**
     * @brief returns whether the chart uses the 1904 date system.
     *
     * If the 1904 date system is used, then all dates and times shall be specified as a
     * decimal number of days since Dec. 31, 1903. If the 1904 date system is
     * not used, then all dates and times shall be specified as a decimal number
     * of days since Dec. 31, 1899.
     *
     * @return valid bool if the 1904 date system is set to true or false, nullopt otherwise.
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
     * @return valid bool if the property is set, nullopt otherwise.
     */
    std::optional<bool> roundedCorners() const;
    /**
     * @brief sets whether the chart area shall have rounded corners.
     * @param rounded
     */
    void setRoundedCorners(bool rounded);

    bool loadFromXmlFile(QIODevice *device) override;
    void saveToXmlFile(QIODevice *device) const override;
    void saveMediaFiles(Workbook *workbook);
    void loadMediaFiles(Workbook *workbook);

private:
    friend class CT_XXXChart;
    Series* seriesByOrder(int order);
    bool hasAxis(int id) const;
};

}

#endif // QXLSX_CHART_H
