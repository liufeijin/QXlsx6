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

QT_BEGIN_NAMESPACE_XLSX

class AbstractSheet;
class Worksheet;
class ChartPrivate;
class CellRange;
class DrawingAnchor;

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
    Text textProperties; /**< text, paragraph and character properties */
    ExtensionList extension;

    void read(QXmlStreamReader &reader);
    void write(QXmlStreamWriter &writer, const QString &name) const;
    bool isValid() const;
};

class QXLSX_EXPORT Chart : public AbstractOOXmlFile
{
    Q_DECLARE_PRIVATE(Chart)
public:
    enum class Type
    { // 16 type of chart (ECMA 376)
        None = 0, // Zero is internally used for unknown types
        Area,
        Area3D,
        Line,
        Line3D,
        Stock,
        Radar,
        Scatter,
        Pie,
        Pie3D,
        Doughnut,
        Bar,
        Bar3D,
        OfPie,
        Surface,
        Surface3D,
        Bubble,
    };
    enum class DisplayBlanksAs
    {
        Span,
        Gap,
        Zero
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

    enum class BubbleSizeRepresents
    {
        Area,
        Width
    };
    enum class OfPieType
    {
        Bar,
        Pie
    };
    enum class SplitType
    {
        Auto,
        Custom,
        Position,
        Percent,
        Value
    };
    enum class BarDirection
    {
        Bar,
        Column
    };
    enum class ScatterStyle
    {
        None,
        Line,
        LineMarker,
        Marker,
        Smooth,
        SmoothMarker,
    };
    enum class RadarStyle
    {
        Standard,
        Marker,
        Filled
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
public:
    ~Chart();
public:
    /**
     * @brief setType sets type to the chart.
     *
     * @note OOXML schema allows having as many subcharts of different types
     * as you need.
     *
     * @warning This method does not change the previous subcharts type nor the
     * prevoius series type.
     *
     * @example If you need to add both bar series and line
     * series to the chart, you can do it this way:
     *
     * @code
     * //set Type::Bar to the chart
     * chart->setType(Chart::Type::Bar);
     * //add bar series
     * chart.addSeries(QXlsx::CellRange("A1:A10"), QXlsx::CellRange("B1:B10"),
     *     xlsx.sheet("Sheet1"), true);
     * //set Type::Line to the chart
     * chart->setType(Chart::Type::Line);
     * //add line series
     * chart.addSeries(QXlsx::CellRange("A1:A10"), QXlsx::CellRange("C1:C10"),
     *     xlsx.sheet("Sheet1"), true);
     * @endcode
     *
     * @param type Chart type. If set to Type::None the chart is ill-formed.
     */
    void setType(Type type);
    /**
     * @brief type returns the chart type.
     *
     * @return
     */
    Type type() const;

    void setChartLineFormat(const LineFormat &format);
    void setPlotAreaLineFormat(const LineFormat &format);

    void setChartShape(const ShapeFormat &shape);
    void setPlotAreaShape(const ShapeFormat &shape);

    /**
     * @brief addSeries adds one or more series defined by range.
     * @param range valid CellRange
     * @param sheet data source for range
     * @param firstRowContainsHeaders
     * @param columnBased specifies that range should be treated as column-based
     * or row-based: if columnBased is true, the first column is category data,
     * other columns are value data for new series. If columnBased is false, the
     * first row is category data, other rows are value data for new series
     */
    void addSeries(const CellRange &range, AbstractSheet *sheet = NULL,
                   bool firstRowContainsHeaders = false, bool columnBased = true);
    /**
     * @brief addSeries adds series
     * @param keyRange category range or x axis range
     * @param valRange value range or y axis range
     * @param sheet data sheet reference
     * @param keyRangeIncludesHeader if true, the first row or column is used as a series name reference
     * @return pointer to the added series or nullptr if no series was added
     */
    Series* addSeries(const CellRange &keyRange, const CellRange &valRange,
                       AbstractSheet *sheet = NULL, bool keyRangeIncludesHeader = false);
    /**
     * @brief addSeries adds an empty series
     * @return reference to the added series
     */
    Series *addSeries();
    /**
     * @brief series
     * @param index
     * @return
     */
    Series *series(int index);

    /**
     * @brief seriesCount
     * @return
     */
    int seriesCount() const;

    /**
     * @brief sets axes for all series on the chart.
     *
     * The method doesn't check for the axes availability and sets axesIds for all series
     * added to the chart. To set specific axes for some series (f.e. to move the series to the
     * right axis) use Series::setAxesIDs(Series *series, const QList<int> &axesIds) method.
     *
     * @param axesIds a list of 0, 2 or 3 items.
     * @note axesIds are ids, not indexes. \see Axis::id()
     */
    void setSeriesAxesIDs(const QList<int> &axesIds);
    /**
     * @brief setSeriesAxesIDs sets axes for the specific series
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
     * @brief setSeriesDefaultAxes sets all available axes for all series
     * added to the chart.
     *
     * The method doesn't check for the axes availability and sets all available
     * axes for all series added to the chart. To set specific axes for some
     * series (f.e. to move the series to the right axis) use
     * Series::setAxesIDs() method.
     *
     * @note invoke this method after you have added all the series and all the
     * axes if all the series share the same axes.
     */
    void setSeriesDefaultAxes();
    /**
     * @brief moveSeries changes the series order in which it is drawn on the chart.
     * @param oldOrder
     * @param newOrder
     */
    void moveSeries(int oldOrder, int newOrder);

    /**
     * @brief addDefaultAxes adds all necessary axes to the chart.
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
     * @brief addAxis adds a new axis with specified parameters.
     * @param type axis type: Cat, Val, Date or Ser.
     * @param pos Axis position: Bottom, Left, Top or Right.
     * @param title optional axis title.
     * @return reference to the newly added axis.
     */
    Axis &addAxis(Axis::Type type, Axis::Position pos, QString title = QString());
    void addAxis(const Axis &axis);
    /**
     * @brief axis returns axis that has index idx
     * @param idx valid index (0 <= idx < axesCount())
     * @return pointer to the axis if such axis exists, nullptr otherwise
     */
    Axis *axis(int idx);

    /**
     * @brief tries to remove axis and returns true if axis has been removed.
     * @param axisID the id of the axis to be removed.
     * @return true if axis has been removed, false otherwise (no axis with this id or axis is used
     * in some series).
     */
    bool removeAxis(int axisID);
    /**
     * @brief tries to remove axis and returns true if axis has been removed.
     * @param axis the pointer to the axis to be removed.
     * @return true if axis has been removed, false otherwise (no such axis or axis is used
     * in some series).
     */
    bool removeAxis(Axis *axis);
    /**
     * @brief axesCount returns the axes count
     * @return
     */
    int axesCount() const;

    /**
     * @brief setTitle sets formatted title to the chart.
     * @param title
     */
    void setTitle(const Title &title);
    /**
     * @brief setTitle sets plain text title to the chart.
     * @param title
     */
    void setTitle(const QString &title);
    /**
     * @brief title returns reference to the chart title
     * @return
     */
    Title &title();
    /**
     * @brief title returns a copy of the chart's title
     * @return
     */
    Title title() const;

    /**
     * @brief setLegend moves the chart legend to position pos with chart overlay
     * @param position
     * @param overlay
     */
    void setLegend(Legend::Position position, bool overlay = false);
    /**
     * @brief legend returns reference to the chart legend.
     * @return reference to the default legend.
     */
    Legend &legend();

    /**
     * @brief grouping returns the kind of grouping for a line or area chart.
     *
     * ApplicableTo: Area, Area3D, Line, Line3D.
     *
     * @return valid Chart::Grouping if property is set, nullopt otherwise.
     */
    std::optional<Chart::Grouping> grouping() const;
    /**
     * @brief setGrouping sets the kind of grouping for a line or area chart.
     *
     * ApplicableTo: Area, Area3D, Line, Line3D.
     *
     * @param grouping Chart::Grouping
     */
    void setGrouping(Chart::Grouping grouping);

    /**
     * @brief barGrouping returns the kind of grouping for a bar chart.
     *
     * ApplicableTo: Bar, Bar3D.
     *
     * @return valid Chart::BarGrouping if property is set, nullopt otherwise.
     */
    std::optional<Chart::BarGrouping> barGrouping() const;
    /**
     * @brief setBarGrouping sets the kind of grouping for a bar chart.
     *
     * Applicable to chart types: Bar, Bar3D.
     *
     * @param grouping Chart::BarGrouping
     */
    void setBarGrouping(Chart::BarGrouping grouping);

    /**
     * @brief varyColors specifies that each data marker in the series has a different color.
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
     * @brief setVaryColors sets that each data marker in the series has a different color.
     *
     * Applicable to chart types: Line, Line3D, Scatter, Radar, Bar, Bar3D, Area, Area3D,
     * Pie, Pie3D, Doughnut, OfPie, Bubble.
     *
     * @param varyColors true if each data marker in the series has a different color, false
     * if every data markers in the series have the same color.
     */
    void setVaryColors(bool varyColors);

    /**
     * @brief labels returns the entire chart labels properties.
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
     * @brief labels returns the entire chart labels properties.
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
     * @brief setLabels sets the entire chart labels properties.
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
     * @brief dropLines returns the chart drop lines.
     *
     * Applicable to chart types: Line, Line3D, Stock, Area, Area3D.
     *
     * @return a copy of the chart dropLines.
     */
    ShapeFormat dropLines() const;
    /**
     * @brief dropLines returns the chart drop lines.
     *
     * Applicable to chart types: Line, Line3D, Stock, Area, Area3D.
     *
     * @return a reference to the chart dropLines.
     */
    ShapeFormat &dropLines();
    /**
     * @brief setDropLines sets the chart drop lines.
     *
     * Applicable to chart types: Line, Line3D, Stock, Area, Area3D.
     *
     * @param dropLines
     */
    void setDropLines(const ShapeFormat &dropLines);

    /**
     * @brief hiLowLines returns the chart high-low lines.
     *
     * Applicable to chart types: Line, Stock.
     *
     * @return a copy of the chart hiLowLines.
     */
    ShapeFormat hiLowLines() const;
    /**
     * @brief hiLowLines returns the chart high-low lines.
     *
     * Applicable to chart types: Line, Stock
     *
     * @return a reference to the chart hiLowLines.
     */
    ShapeFormat &hiLowLines();
    /**
     * @brief setHLowLines sets the chart high-low lines.
     *
     * Applicable to chart types: Line, Stock
     *
     * @param hiLowLines
     */
    void setHiLowLines(const ShapeFormat &hiLowLines);

    /**
     * @brief upDownBars returns the chart up-dow lines.
     *
     * Applicable to chart types: Line, Stock.
     *
     * @return a copy of the chart upDownBars.
     */
    UpDownBar upDownBars() const;
    /**
     * @brief upDownBars returns the chart up-down lines.
     *
     * Applicable to chart types: Line, Stock
     *
     * @return a reference to the chart upDownBars.
     */
    UpDownBar &upDownBars();
    /**
     * @brief setUpDownBars sets the chart up-down lines.
     *
     * Applicable to chart types: Line, Stock
     *
     * @param upDownBars
     */
    void setUpDownBars(const UpDownBar &upDownBars);

    /**
     * @brief markerShown if true, the chart series marker is shown.
     *
     * Applicable to chart types: Line.
     *
     * @return valid boolean value if property is set, nullopt otherwise.
     */
    std::optional<bool> markerShown() const;
    /**
     * @brief setMarkerShown sets that the chart series marker shall be shown.
     *
     * Applicable to chart types: Line.
     *
     * @param markerShown
     */
    void setMarkerShown(bool markerShown);

    /**
     * @brief smooth returns smoothing of the chart series. If true, the line
     * connecting the points on the chart is smoothed using Catmull-Rom splines.
     *
     * Applicable to chart types: Line.
     *
     * @return valid boolean value if property is set, nullopt otherwise.
     */
    std::optional<bool> smooth() const;
    /**
     * @brief setSmooth sets smoothing of the chart series. If true, the line
     * connecting the points on the chart shall be smoothed using Catmull-Rom splines.
     *
     * Applicable to chart types: Line.
     *
     * @param smooth
     */
    void setSmooth(bool smooth);

    /**
     * @brief gapDepth specifies the space between bar or column clusters, as a
     * percentage of the bar or column width.
     *
     * Applicable to chart types: Line3D, Bar3D, Area3D.
     *
     * @return valid double ([0..]) value if property is set, nullopt otherwise.
     */
    std::optional<int> gapDepth() const;
    /**
     * @brief setGapDepth sets the space between bar or column clusters, as a
     * percentage of the bar or column width.
     *
     * Applicable to chart types: Line3D, Bar3D, Area3D.
     *
     * @param gapDepth value of the gap depth, in percents (100 equals the bar width).
     */
    void setGapDepth(int gapDepth);

    /**
     * @brief scatterStyle returns the style of the scatter chart.
     *
     * The default value is ScatterStyle::Marker.
     *
     * @return Chart::ScatterStyle
     */
    ScatterStyle scatterStyle() const;
    /**
     * @brief setScatterStyle sets the style of the scatter chart.
     *
     * If not set, the default value is ScatterStyle::Marker.
     *
     * @param scatterStyle
     */
    void setScatterStyle(ScatterStyle scatterStyle);

    /**
     * @brief radarStyle returns the style of the radar chart.
     *
     * The default value is RadarStyle::Standard.
     *
     * @return RadarStyle
     */
    RadarStyle radarStyle() const;
    /**
     * @brief setRadarStyle sets the style of the radar chart.
     *
     * If not set, the default value is RadarStyle::Standard.
     *
     * @param radarStyle
     */
    void setRadarStyle(RadarStyle radarStyle);

    /**
     * @brief barDirection returns whether the series form a bar (horizontal)
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
     * @brief setBarDirection sets whether the series form a bar (horizontal)
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
     * @brief gapWidth returns the space between bar or column clusters, as a
     * percentage of the bar or column width.
     *
     * Applicable to chart types: Bar, Bar3D, OfPie.
     *
     * @return valid int ([0..500]) value if property is set, nullopt otherwise.
     */
    std::optional<int> gapWidth() const;
    /**
     * @brief setGapWidth sets the space between bar or column clusters, as a
     * percentage of the bar or column width.
     *
     * Applicable to chart types: Bar, Bar3D, OfPie.
     *
     * @param gapWidth gap width in percents (0..500).
     */
    void setGapWidth(int gapWidth);

    /**
     * @brief overlap specifies how much bars and columns shall overlap on 2-D charts.
     *
     * Applicable to chart types: Bar.
     *
     * @return valid int value (percentage) if property is set, nullopt otherwise.
     */
    std::optional<int> overlap() const;
    /**
     * @brief setOverlap sets how much bars and columns shall overlap on 2-D charts.
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
     * @brief holeSize returns the size, in percents, of the hole in a doughnut chart.
     *
     * Applicable to chart types: Doughnut.
     *
     * @return valid int value if the property is set (in the range [1..90]), nullopt otherwise.
     */
    std::optional<int> holeSize() const;
    /**
     * @brief setHoleSise sets the size, in percents, of the hole in a doughnut chart.
     *
     * Applicable to chart types: Doughnut.
     *
     * @param holeSize the hole size in percents in the range [1..90]
     */
    void setHoleSise(int holeSize);

    /**
     * @brief setDataTableVisible sets the visibility of the chart data table
     * @param visible if true, the chart data table is shown below the chart. If false,
     * the chart data table is hidden.
     *
     * @warning Setting the data table visibility to false deletes the current data table properties.
     *
     */
    void setDataTableVisible(bool visible);
    /**
     * @brief dataTable returns the chart data table.
     *
     * Changing the data table properties also makes it visible.
     *
     * @return reference to the data table.
     */
    DataTable &dataTable();
    /**
     * @brief dataTable returns the chart data table.
     *
     * @return copy of the chart data table.
     */
    DataTable dataTable() const;
    /**
     * @brief setDataTable sets the chart data table.
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
     * @brief date1904 returns that the chart uses the 1904 date system.
     *
     * This method returns that the chart uses the 1904 date system. If the 1904
     * date system is used, then all dates and times shall be specified as a
     * decimal number of days since Dec. 31, 1903. If the 1904 date system is
     * not used, then all dates and times shall be specified as a decimal number
     * of days since Dec. 31, 1899.
     *
     * @return valid bool if the 1904 date system is set to true or false, nullopt otherwise.
     */
    std::optional<bool> date1904() const;
    /**
     * @brief setDate1904 sets that the chart uses the 1904 date system.
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
     * @brief lastUsedLanguage returns the primary editing language (f.e. "en-US")
     * which was used when this chart was last modified.
     *
     * @return language id or empty string if no last used language was set.
     */
    QString lastUsedLanguage() const;
    /**
     * @brief setLastUsedLanguage sets the primary editing language (f.e. "en-US")
     * which was used when this chart was last modified.
     *
     * @param lang language id
     */
    void setLastUsedLanguage(QString lang);

    /**
     * @brief roundedCorners returns whether the chart area has rounded corners.
     * @return valid bool if the property is set, nullopt otherwise.
     */
    std::optional<bool> roundedCorners() const;
    /**
     * @brief setRoundedCorners sets whether the chart area shall have rounded corners.
     * @param rounded
     */
    void setRoundedCorners(bool rounded);

    bool loadFromXmlFile(QIODevice *device) override;
    void saveToXmlFile(QIODevice *device) const override;
private:
    friend class CT_XXXChart;
    Series* seriesByOrder(int order);
    bool hasAxis(int id) const;


};

QT_END_NAMESPACE_XLSX

#endif // QXLSX_CHART_H
