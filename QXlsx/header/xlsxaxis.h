#ifndef XLSXAXIS_H
#define XLSXAXIS_H

#include <QFont>
#include <QColor>
#include <QByteArray>
#include <QList>
#include <QVariant>
#include <QXmlStreamWriter>
#include <QXmlStreamReader>
#include <QSharedData>

#include "xlsxglobal.h"
#include "xlsxshapeformat.h"
#include "xlsxutility_p.h"
#include "xlsxtitle.h"
#include "xlsxtext.h"

namespace QXlsx {

class DisplayUnitsPrivate;
/**
 * @brief The DisplayUnits class specifies the units properties for the value axis
 */
class QXLSX_EXPORT DisplayUnits
{
public:
    enum class BuiltInUnit
    {
        Hundreds,
        Thousands,
        TenThousands,
        HundredThousands,
        Millions,
        TenMillions,
        HundredMillions,
        Billions,
        Trillions,
    };
    DisplayUnits();
    DisplayUnits(double customUnit);
    DisplayUnits(BuiltInUnit unit);
    DisplayUnits(const DisplayUnits &other);
    ~DisplayUnits();
    DisplayUnits &operator=(const DisplayUnits &other);

    std::optional<double> customUnit() const;
    void setCustomUnit(double customUnit);

    std::optional<BuiltInUnit> builtInUnit() const;
    void setBuiltInUnit(BuiltInUnit builtInUnit);

    /**
     * @brief labelVisible returns the visibility of the display units label
     * @return
     */
    bool labelVisible() const;
    /**
     * @brief setLabelVisible shows/hides the display units label
     * @param visible
     */
    void setLabelVisible(bool visible);

    Title title() const;
    Title &title();
    void setTitle(const Title &title);

    bool isValid() const;

    void read(QXmlStreamReader &reader);
    void write(QXmlStreamWriter &writer, const QString &name) const;

    bool operator ==(const DisplayUnits &other) const;
    bool operator !=(const DisplayUnits &other) const;
private:
    SERIALIZE_ENUM(BuiltInUnit,
    {
        {BuiltInUnit::Hundreds, "hundreds"},
        {BuiltInUnit::Thousands, "thousands"},
        {BuiltInUnit::TenThousands, "tenThousands"},
        {BuiltInUnit::HundredThousands, "hundredThousands"},
        {BuiltInUnit::Millions, "millions"},
        {BuiltInUnit::TenMillions, "tenMillions"},
        {BuiltInUnit::HundredMillions, "hundredMillions"},
        {BuiltInUnit::Billions, "billions"},
        {BuiltInUnit::Trillions, "trillions"},
    });

    QSharedDataPointer<DisplayUnitsPrivate> d;
};

class AxisPrivate;
/**
 * @brief The Axis class represents an axis on a chart.
 *
 * The Axis class has a number of required parameters: #type(), #position(), #crossAxis()
 * and various optional parameters that specify its look and behaviour.
 *
 * The class is _explicitly shareable_: unless you add a copy on the chart
 * the two copies of Axis represent the same object.
 *
 * A chart can have as many axes as you need. But caution should be applied to properly
 * position axes on the chart. See [CombinedChart example](../../CombinedChart/main.cpp) of
 * three axes: bottom, left and right one.
 *
 * To add an axis on a chart use Chart::addDefaultAxes and Chart::addAxis methods.
 */
class QXLSX_EXPORT Axis
{
public:
    /**
     * @brief The Scaling class is used to set the axis range and scaling parameters
     */
    struct Scaling {
        std::optional<double> logBase;
        std::optional<bool> reversed;
        std::optional<double> min;
        std::optional<double> max;
        /**
         * @brief isValid returns true if some properties of axis scaling were set, false otherwise.
         * @return
         */
        bool isValid() const;
        void write(QXmlStreamWriter &writer) const;
        void read(QXmlStreamReader &reader);
        bool operator ==(const Scaling &other) const;
        bool operator !=(const Scaling &other) const;
    };
    /**
     * @brief The Type enum specifies the axis type
     */
    enum class Type {
        None = (-1), /**< @brief the invalid axis type */
        Category, /**< @brief Axis represents categorized values */
        Value,  /**< @brief Axis represents numerical values */
        Date,  /**< @brief Axis represents date and time values */
        Series  /**< @brief Axis represents series names */
    };
    /**
     * @brief The Position enum specifies the axis position
     */
    enum class Position {
        None = (-1), /**< @brief the invalid axis position */
        Left, /**< @brief the axis is positioned on the left */
        Right, /**< @brief the axis is positioned on the right */
        Top, /**< @brief the axis is positioned on top of the chart */
        Bottom /**< @brief the axis is positioned on bottom of the chart */
    };
    /**
     * @brief The CrossesType enum specifies the possible crossing points for an axis.
     */
    enum class CrossesType {
        Minimum, /**< @brief Axis crosses at the minimum value of the chart. */
        Maximum, /**< @brief Axis crosses at the maximum value of the chart. */
        Position, /**< @brief Axis crosses at the value specified with setCrossesAt() */
        AutoZero /**< @brief The category axis crosses at the zero point of the value axis (if possible),
                      or the minimum value (if the minimum is greater than zero) or the maximum
                      (if the maximum is less than zero) */
    };
    /**
     * @brief The CrossesBetweenType enum  specifies the possible crossing states of an axis.
     */
    enum class CrossesBetweenType
    {
        Between, /**< @brief the value axis shall cross the category axis between data markers. */
        MidCat /**< @brief the value axis shall cross the category axis at the midpoint of a category. */
    };
    /**
     * @brief The TickMark enum specifies the possible positions for tick marks.
     */
    enum class TickMark
    {
        None, /**< @brief there shall be no tick marks.*/
        In, /**< @brief the tick marks shall be inside the plot area. */
        Out, /**<  @brief the tick marks shall be outside the plot area. */
        Cross /**< @brief the tick marks shall cross the axis. */
    };
    /**
     * @brief The TickLabelPosition enum specifies the possible positions for tick labels
     */
    enum class TickLabelPosition
    {
        High, /**< @brief  the axis labels shall be at the high end of the perpendicular axis */
        Low, /**< @brief the axis labels shall be at the low end of the perpendicular axis */
        NextTo, /**< @brief the axis labels shall be next to the axis */
        None /**< @brief the axis labels are not drawn */
    };

    /**
     * @brief The LabelAlignment enum specifies the possible ways to align the tick labels
     */
    enum class LabelAlignment
    {
        Center, /**< @brief the text shall be centered */
        Left, /**< @brief the text shall be left justified */
        Right /**< @brief the text shall be right justified */
    };
    /**
     * @brief The TimeUnit enum  specifies a unit of time to use on a date axis.
     */
    enum class TimeUnit
    {
        Days, /**< @brief the chart data shall be shown in days. */
        Months, /**< @brief the chart data shall be shown in months. */
        Years /**<  @brief the chart data shall be shown in years. */
    };

    /**
     * @brief creates an invalid axis. To make it valid set type and position
     * and add this axis on the chart using Chart::addAxis(const Axis &axis)
     * method.
     */
    Axis();
    /**
     * @brief creates an axis with specified type and position. To add this axis
     * on the chart use Chart::addAxis(const Axis &axis) method.
     * @param type
     * @param position
     */
    Axis(Type type, Position position);
    /**
     * @brief creates a copy of the other axis. The copy will be the same axis as
     * the other axis, until you invoke Chart::addAxis(const Axis &axis) method
     * to add the copy on the chart.
     * @param other
     */
    Axis(const Axis &other);
    Axis &operator=(const Axis &other);
    ~Axis();

    bool isValid() const;

    /**
     * @brief returns the axis type.
     * @return
     */
    Type type() const;
    /**
     * @brief sets the axis type.
     * @param type
     */
    void setType(Type type);
    /**
     * @brief returns the axis position.
     * @return
     */
    Position position() const;
    /**
     * @brief sets the axis position.
     * @param position
     */
    void setPosition(Position position);
    /**
     * @brief returns whether the axis is visible on the chart.
     * @return valid optional if the parameter was set, nullopt otherwise. By default
     * the axis is visible.
     */
    std::optional<bool> visible() const;
    /**
     * @brief sets whether the axis is visible on the chart.
     * @param visible false if the axis should be hidden. By default
     * the axis is visible.
     * @note This method does not allow to hide the axis title. To do this, simply
     * set the empty title.
     */
    void setVisible(bool visible);

    /**
     * @brief returns the unique ID of the axis. This ID is used by processing apps
     * (like Excel) to identify the axis on a chart.
     * @return positive value if the axis was added to a chart, -1 otherwise.
     */
    int id() const;

    /**
     * @brief returns the axis that should cross the current axis.
     * @return pointer to the axis if cross axis was set, `nullptr` otherwise.
     */
    Axis *crossAxis() const;
    /**
     * @brief sets the axis that should cross the current axis.
     * @param axis pointer to the axis.
     */
    void setCrossAxis(Axis *axis);

    /**
     * @brief returns the position on the cross axis in which this axis should cross.
     * This parameter applies only if crossesType() is set to CrossesType::Position.
     * @return double value if the parameter was set, nullopt otherwise. The default value
     * is chosen by the processing app automatically.
     */
    std::optional<double> crossesAt() const;
    /**
     * @brief sets the position on the cross axis in which this axis should cross.
     * This parameter applies only if crossesType() is set to CrossesType::Position.
     * @note This method does not change the crossesType.
     * @param val
     */
    void setCrossesAt(double val);
    /**
     * @brief returns the mode this axis should cross the crossAxis.
     * @return
     */
    std::optional<CrossesType> crossesType() const;
    /**
     * @brief sets the mode this axis should cross the crossAxis. If val is CrossesType::Position,
     * #setCrossesAt should also be invoked.
     * @param val
     */
    void setCrossesType(CrossesType val);

    /**
     * @brief sets the axis' major grid lines as a ShapeFormat
     * @param val
     */
    void setMajorGridLines(const ShapeFormat &val);
    /**
     * @brief returns the axis' major grid lines.
     * @return reference to the ShapeFormat object.
     */
    ShapeFormat &majorGridLines();
    /**
     * @brief returns the axis' major grid lines.
     * @return shallow copy of the ShapeFormat object.
     */
    ShapeFormat majorGridLines() const;
    /**
     * @brief sets the axis' minor grid lines as a ShapeFormat
     * @param val
     */
    void setMinorGridLines(const ShapeFormat &val);
    /**
     * @brief returns the axis' minor grid lines.
     * @return reference to the ShapeFormat object.
     */
    ShapeFormat &minorGridLines();
    /**
     * @brief returns the axis' minor grid lines.
     * @return shallow copy of the ShapeFormat object.
     */
    ShapeFormat minorGridLines() const;

    /**
     * @brief sets the axis' major grid lines properties in the more convenient way.
     * This method can be used to set the basic parameters of the major grid lines.
     * To fine-tune parameters, use #setMajorGridLines(const ShapeFormat &val)
     * and #majorGridLines().
     * @param color grid line color.
     * @param width grid line width in points.
     * @param strokeType grid line type.
     */
    void setMajorGridLines(const QColor &color, double width, LineFormat::StrokeType strokeType);
    /**
     * @brief turns on/off the default major grid lines.
     * @param on
     */
    void setMajorGridLines(bool on);
    /**
     * @brief sets the axis' minor grid lines properties in the more convenient way.
     * This method can be used to set the basic parameters of the minor grid lines.
     * To fine-tune parameters, use #setMinorGridLines(const ShapeFormat &val)
     * and #minorGridLines().
     * @param color grid line color.
     * @param width grid line width in points.
     * @param strokeType grid line type.
     */
    void setMinorGridLines(const QColor &color, double width, LineFormat::StrokeType strokeType);
    /**
     * @brief turns on/off the default minor grid lines.
     * @param on
     */
    void setMinorGridLines(bool on);
    /**
     * @brief sets the position of the axis' major tick marks.
     * @param tickMark
     */
    void setMajorTickMark(TickMark tickMark);
    /**
     * @brief sets the position of the axis' minor tick marks.
     *
     * @param tickMark
     */
    void setMinorTickMark(TickMark tickMark);
    /**
     * @brief returns the position of the axis' major tick marks.
     * @return nullopt if the parameter was not set.
     */
    std::optional<TickMark> majorTickMark() const;
    /**
     * @brief returns the position of the axis' minor tick marks.
     * @return nullopt if the parameter was not set.
     */
    std::optional<TickMark> minorTickMark() const;
    /**
     * @brief returns the axis title stripped of all formatting.
     * @return non-empty string if title was set.
     */
    QString titleAsString() const;
    /**
     * @brief returns the axis title.
     * @return shallow copy of the Title object.
     */
    Title title() const;
    /**
     * @brief returns the axis title.
     * @return reference to the Title object.
     */
    Title &title();
    /**
     * @brief sets the axis title as a plain string.
     * @param title axis title.
     * @note If you want to set the axis title as a string reference or as a
     * formatted text, create new title and use the overloaded method.
     * @code
     * Title t;
     * t.setStringReference("Sheet1!A2");
     * axis->setTitle(t);
     * @endcode
     */
    void setTitle(const QString &title);
    /**
     * @brief sets the axis title.
     * @param title
     */
    void setTitle(const Title &title);

    /**
     * @brief returns the axis text formatting.
     * @return reference to the TextFormat object.
     * @note This object specifies formatting parameters for the axis text elements
     * except the axis title.
     */
    TextFormat &textProperties();
    /**
     * @brief returns the axis text formatting.
     * @return shallow copy of the TextFormat object.
     * @note This object specifies formatting parameters for the axis text elements
     * except the axis title.
     */
    TextFormat textProperties() const;
    /**
     * @brief sets the axis text formatting.
     * @param textProperties
     * @note This object specifies formatting parameters for the axis text elements
     * except the axis title.
     */
    void setTextProperties(const TextFormat &textProperties);
    /**
     * @brief returns the axis number format
     * @return non-empty string if the parameter is set.
     * By default numberFormat is "General".
     */
    //TODO: create a class NumberFormat with validation and use it everywhere.
    QString numberFormat() const;
    /**
     * @brief sets the axis number format.
     * @param formatCode string representation of the number format.
     */
    void setNumberFormat(const QString &formatCode);
    /**
     * @brief returns the axis range/scaling parameters.
     * @return reference to the Scaling object.
     */
    Scaling &scaling();
    /**
     * @brief returns the axis range/scaling parameters.
     * @return copy of the Scaling object.
     */
    Scaling scaling() const;
    /**
     * @brief sets the axis range/scaling parameters.
     * @param scaling
     */
    void setScaling(Scaling scaling);
    /**
     * @brief returns the axis range.
     * @note This is a convenience method. It is equivalent to this code:
     * @code
     * const auto scale = axis->scaling();
     * QPair<double, double> range = qMakePair<double, double>(scale.min.value_or(0.0),
     *     scale.max.value_or(0.0));
     * @endcode
     * @return
     */
    QPair<double, double> range() const;
    /**
     * @brief sets the axis range.
     * @note This is a convenience method. It is equivalent to this code:
     * @code
     * auto &scale = axis->scaling();
     * scale.min = min;
     * scale.max = max;
     * @endcode
     * @param min
     * @param max
     */
    void setRange(double min, double max);

    /**
     * @brief specifies that this axis is a date or text axis based on
     * the data that is used for the axis labels, not a specific choice.
     *
     * Applicable to: Category axis, Date axis.
     *
     * @return valid bool value if the property is set, nullopt otherwise.
     */
    std::optional<bool> autoAxis() const;
    /**
     * @brief sets that this axis is a date or text axis based on
     * the data that is used for the axis labels, not a specific choice.
     *
     * Applicable to: Category axis, Date axis.
     *
     * @param autoAxis
     */
    void setAutoAxis(bool autoAxis);

    /**
     * @brief returns labels alignment for the category axis.
     * @return
     */
    std::optional<Axis::LabelAlignment> labelAlignment() const;
    /**
     * @brief sets labels alignment for the category axis.
     * @param labelAlignment
     */
    void setLabelAlignment(Axis::LabelAlignment labelAlignment);

    /**
     * @brief returns the distance of labels, in percents, from the axis.
     *
     * Applicable to: Category axis, Date axis.
     *
     * @return valid optional value if labelOffset is set, nullopt otherwise.
     */
    std::optional<int> labelOffset() const;
    /**
     * @brief sets the distance of labels, in percents, from the axis.
     *
     * If not set, the default value is 100.
     *
     * Applicable to: Category axis, Date axis.
     *
     * @param labelOffset value in percents in the range [0..1000].
     */
    void setLabelOffset(int labelOffset);

    /**
     * @brief returns the distance between major axis ticks.
     *
     * Applicable to: Date axis, Value axis.
     *
     * @return positive value if the property is set, nullopt otherwise.
     */
    std::optional<double> majorTickDistance() const;
    /**
     * @brief sets the distance between major axis ticks.
     * Applicable to: Date axis, Value axis.
     * @param distance positive value.
     */
    void setMajorTickDistance(double distance);
    /**
     * @brief returns the distance between minor axis ticks.
     *
     * Applicable to: Date axis, Value axis.
     *
     * @return positive value if the property is set, nullopt otherwise.
     */
    std::optional<double> minorTickDistance() const;
    /**
     * @brief sets the distance between minor axis ticks.
     * Applicable to: Date axis, Value axis.
     * @param distance positive value.
     */
    void setMinorTickDistance(double distance);

    /**
     * @brief returns the smallest time unit that is represented on the date axis.
     * @return
     */
    std::optional<Axis::TimeUnit> baseTimeUnit() const;
    /**
     * @brief sets the smallest time unit that is represented on the date axis.
     * @param baseTimeUnit
     */
    void setBaseTimeUnit(TimeUnit baseTimeUnit);
    /**
     * @brief returns time unit for major tick marks of a date axis.
     * @return
     */
    std::optional<Axis::TimeUnit> majorTimeUnit() const;
    /**
     * @brief sets time unit for major tick marks of a date axis.
     * @param majorTimeUnit
     */
    void setMajorTimeUnit(TimeUnit majorTimeUnit);
    /**
     * @brief returns time unit for minor tick marks of a date axis.
     * @return
     */
    std::optional<Axis::TimeUnit> minorTimeUnit() const;
    /**
     * @brief sets time unit for minor tick marks of a date axis.
     * @param minorTimeUnit
     */
    void setMinorTimeUnit(TimeUnit minorTimeUnit);

    /**
     * @brief returns how many tick labels is skipped between label that is drawn.
     *
     * Applicable to: Category axis, Series axis.
     *
     * @return
     */
    std::optional<int> tickLabelSkip();
    /**
     * @brief sets how many tick labels shall be skipped between label that is drawn.
     *
     * Applicable to: Category axis, Series axis.
     *
     * @param skip
     */
    void setTickLabelSkip(int skip);
    /**
     * @brief returns how many tick marks shall be skipped before
     * the next one shall be drawn.
     *
     * Applicable to: Category axis, Series axis.
     *
     * @return
     */
    std::optional<int> tickMarkSkip();
    /**
     * @brief sets  how many tick marks shall be skipped before
     * the next one shall be drawn.
     *
     * Applicable to: Category axis, Series axis.
     *
     * @param skip
     */
    void setTickMarkSkip(int skip);

    /**
     * @brief specifies the labels on the category axis shall
     * be shown as flat text.
     * If this element is not included or is set to false, then the labels shall
     * be drawn as a hierarchy.
     * @return
     */
    std::optional<bool> noMultiLevelLabels() const;
    /**
     * @brief specifies the labels on the category axis shall
     * be shown as flat text.
     *
     * @param val true means that the labels shall be drawn as a flat text.
     * False means that the labels shall be drawn as a hierarchy.
     */
    void setNoMultiLevelLabels(bool val);

    /**
     * @brief specifies whether the value axis crosses the category
     * axis between categories.
     * @return
     */
    std::optional<CrossesBetweenType> crossesBetween() const;
    /**
     * @brief sets where the value axis crosses the category axis.
     * @param crossesBetween CrossesBetweenType::Between means the value axis shall cross the category axis
     * between categories. CrossesBetweenType::MidCat means the value axis shall cross the category
     * axis on categories.
     */
    void setCrossesBetween(CrossesBetweenType crossesBetween);

    /**
     * @brief returns the properties of the display units for the value axis.
     * @return shallow copy of the DisplayUnits object.
     */
    DisplayUnits displayUnits() const;
    /**
     * @brief returns the properties of the display units for the value axis.
     * @return reference to the DisplayUnits object.
     */
    DisplayUnits &displayUnits();
    /**
     * @brief sets the properties of the display units for the value axis.
     * @param displayUnits
     */
    void setDisplayUnits(const DisplayUnits &displayUnits);

    ShapeFormat shape() const;
    ShapeFormat &shape();
    void setShape(const ShapeFormat &shape);

    void write(QXmlStreamWriter &writer) const;
    void read(QXmlStreamReader &reader);

    bool operator == (const Axis &other) const;
    bool operator != (const Axis &other) const;

private:
    friend class ChartPrivate;
    void setId(int id);
    SERIALIZE_ENUM(Type, {
        {Type::None, "none"},
        {Type::Category, "cat"},
        {Type::Value, "val"},
        {Type::Date, "date"},
        {Type::Series, "ser"},
    });

    SERIALIZE_ENUM(Position, {
        {Position::None, "none"},
        {Position::Top, "t"},
        {Position::Bottom, "b"},
        {Position::Left, "l"},
        {Position::Right, "r"}
    });
    SERIALIZE_ENUM(TickMark, {
        {TickMark::None, "none"},
        {TickMark::In, "in"},
        {TickMark::Out, "out"},
        {TickMark::Cross, "cross"},
    });
    SERIALIZE_ENUM(LabelAlignment, {
        {LabelAlignment::Left, "l"},
        {LabelAlignment::Center, "ctr"},
        {LabelAlignment::Right, "r"},
    });
    SERIALIZE_ENUM(TimeUnit, {
        {TimeUnit::Days, "days"},
        {TimeUnit::Months, "months"},
        {TimeUnit::Years, "years"},
    });
    SERIALIZE_ENUM(CrossesBetweenType, {
        {CrossesBetweenType::MidCat, "midCat"},
        {CrossesBetweenType::Between, "between"},
    });
    SERIALIZE_ENUM(TickLabelPosition, {
        {TickLabelPosition::High, "high"},
        {TickLabelPosition::Low, "low"},
        {TickLabelPosition::NextTo, "nextTo"},
        {TickLabelPosition::None, "none"},
    });

    friend QDebug operator<<(QDebug, const Axis &axis);

    QExplicitlySharedDataPointer<AxisPrivate> d;
};

#ifndef QT_NO_DEBUG_STREAM
QDebug operator<<(QDebug dbg, const Axis &axis);
#endif



}

Q_DECLARE_TYPEINFO(QXlsx::DisplayUnits, Q_MOVABLE_TYPE);
Q_DECLARE_TYPEINFO(QXlsx::Axis, Q_MOVABLE_TYPE);


#endif // XLSXAXIS_H
