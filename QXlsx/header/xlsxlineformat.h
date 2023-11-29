// xlsxformat.h

#ifndef QXLSX_LINEFORMAT_H
#define QXLSX_LINEFORMAT_H

#include <QFont>
#include <QColor>
#include <QPen>
#include <QByteArray>
#include <QList>
#include <QVariant>
#include <QXmlStreamWriter>
#include <QSharedData>

#include "xlsxglobal.h"
#include "xlsxcolor.h"
#include "xlsxfillformat.h"
#include "xlsxmain.h"
#include "xlsxutility_p.h"

namespace QXlsx {

class Worksheet;
class WorksheetPrivate;
class RichStringPrivate;
class SharedStrings;
class LineFormatPrivate;

/**
 * @brief The LineFormat class represents parameters of lines that are used to
 * draw series, axes, grid lines, borders of shapes, etc. on a chart.
 *
 * Line can have:
 *
 * - width (#width(), #setWidth())
 * - pen alignment (#penAlignment(), #setPenAlignment())
 * - compound type (#compoundLineType(), #setCompoundLineType())
 * - stroke type (#strokeType(), #setStrokeType(), #dashPattern(), #setDashPattern())
 * - cap style (lineCap(), #setLineCap())
 * - join style (#lineJoin(), #setLineJoin(), #miterLimit(), #setMiterLimit())
 * - line start and line end shape parameters
 *   - type (#lineEndType(), #lineStartType(), #setLineEndType(), #setLineStartType())
 *   - length (#lineEndLength(), #lineStartLength(), #setLineEndLength(), #setLineStartLength())
 *   - width (#lineEndWidth(), #lineStartWidth(), #setLineEndWidth(), #setLineStartWidth)
 * - fill parameters that can be specified via #fill() and #setFill() methods.
 * See the FillFormat class documentation.
 *   - or with the convenience methods #color(), #setColor(), #type(), #setType().
 *
 *
 * The class is _implicitly shareable_: the deep copy occurs only in the
 * non-const methods.
 */
class QXLSX_EXPORT LineFormat
{
public:
    /**
     * @brief The CompoundLineType enum specifies the type of compound lines
     */
    enum class CompoundLineType {
        Single, /**< Single line */
        Double, /**< Double line of equal width */
        ThickThin, /**< Double line where the first line is thicker */
        ThinThick, /**< Double line where the the second line is thicker */
        Triple, /**< Triple line */
    };
    /**
     * @brief The StrokeType enum  specifies line dash (stroke) types.
     */
    enum class StrokeType {
        Solid, /**< @brief solid line, 1 */
        Dot, /**< @brief dot line, 1000 */
        SystemDot, /**< @brief dot line, 10 */
        Dash, /**< @brief dash line, 1111000 */
        SystemDash, /**< @brief dash line, 1110 */
        DashDot, /**< @brief dash-dot line, 11110001000 */
        SystemDashDot, /**< @brief dash-dot line, 111010 */
        LongDash, /**< @brief long dash line, 11111111000 */
        LongDashDot, /**< @brief long dash-dot line, 111111110001000 */
        LongDashDotDot, /**< @brief long dash-dot-dot line, 1111111100010001000 */
        SystemDashDotDot, /**< @brief dash-dot-dot line, 11101010 */
        Custom /**< @brief custom dash line specified with dashPattern() */
    };
    /**
     * @brief The LineJoin enum specifies the type of line joining
     */
    enum class LineJoin {
        Round, /**< Lines are rounded at the joint. */
        Bevel, /**< Joined lines form a triangular notch. */
        Miter /**< The outer edges of the lines are extended to meet at an angle */
    };
    /**
     * @brief The LineCap enum specifies the end line caps.
     */
    enum class LineCap {
        Square, /**< Line ends with a square that covers the end point and
extends beyond it by half the line width. */
        Round, /**< Line end is rounded. */
        Flat, /**< Line ends with a square that doesn't cover the end point. */
    };

    /**
     * @brief The LineEndType enum specifies the additional shape of a line end.
     */
    enum class LineEndType {
        None, /**< No shape, line ends as specified by its line cap. */
        Triangle, /**< Line ends with a filled arrow-like triangle. */
        Stealth, /**< Line ends with a curved arrow. */
        Oval, /**< Line ends with a filled oval. */
        Arrow, /**< Line ands with a straight arrow. */
        Diamond /**< Line ends with a diamond shape. */
    };

    /**
     * @brief The LineEndSize enum specifies the size of the line end.
     */
    enum class LineEndSize {
        Small, /**< Small end */
        Medium, /**< Medium end */
        Large /**< Large end */
    };

    /**
     * @brief The PenAlignment enum specifies the alignment of the line with the
     * path stroke.
     */
    enum class PenAlignment
    {
        Center, /**< Line is drawn at the center of the path stroke. */
        Inset /**< Line is drawn on the inside edge of the path stroke. */
    };
    /**
     * @brief creates an invalid line format.
     *
     * There is a distinction between the line with no fill and the line with
     * the default fill. To hide the line use
     *
     * @code
     * series(0)->setLine(LineFormat(FillFormat::FillType::NoFill));
     * @endcode
     *
     * An invalid line format can be used as a 'default' line format to clear
     * the previously set line parameters:
     *
     * @code
     * series(0)->setLine({});
     * @endcode
     */
    LineFormat();
    /**
     * @brief creates a new line format with specified fill, width and color.
     * @param fill line fill type.
     * @param widthInPt line width specified in points.
     * @param color line color as an RGB value.
     */
    LineFormat(FillFormat::FillType fill, double widthInPt, QColor color);
    /**
     * @brief creates a new line format with specified fill, width and color.
     * @param fill line fill.
     * @param widthInEMU line width specified in EMU (1 pt = 12700 EMU).
     * @param color line color as an RGB value.
     */
    LineFormat(FillFormat::FillType fill, qint64 widthInEMU, QColor color);
    /**
     * @brief creates a new line format and fills its parameters from a Qt pen.
     * @param pen pen properties.
     *
     * @note Not all QPen properties are supported:
     * - Qt::SvgMiterJoin is treated as LineFormat::LineJoin::Miter;
     * - pen width is always saved in pixels as a string with 2 digits after the decimal point (f.e. "1.25px");
     * - dashOffset is ingnored;
     * - brush is parsed with limitations. See FillFormat::setBrush().
     */
    LineFormat(const QPen &pen);

    LineFormat(const LineFormat &other);
    LineFormat &operator=(const LineFormat &rhs);
    ~LineFormat();

    /**
     * @brief returns line format converted to a Qt pen.
     * @return a QPen object.
     * @note Not all LineFormat parameters are supported:
     * - if set, line width is converted to pixels at 96 DPI with floating-type precision,
     * otherwise pen width is set to 0;
     * - fill format is converted with limitations. See FillFormat::toBrush().
     * - penAlignment is ignored;
     * - if not set, strokeType is converted as Qt::SolidLine;
     * - compoundLineType is ignored;
     * - tail and head parameters are ignored.
     */
    QPen toPen() const;
    /**
     * @brief sets line paramaters from a Qt pen.
     * @param pen
     * @note Not all QPen properties are supported:
     * - Qt::SvgMiterJoin is treated as LineFormat::LineJoin::Miter;
     * - pen width is always saved in pixels as a string with 2 digits after the decimal point (f.e. "1.25px");
     * - dashOffset is ingnored;
     * - brush is parsed with limitations. See FillFormat::setBrush().
     */
    void setPen(const QPen &pen);
    /**
     * @brief returns the line fill type.
     *
     * This is a convenience method, equivalent to `fill().type()`.
     */
    std::optional<FillFormat::FillType> type() const;
    /**
     * @brief sets the line fill type.
     * @param type A FillFormat::FillType enum value.
     *
     * This is a convenience method equivalent to `fill().setType(type)`.
     */
    void setType(FillFormat::FillType type);

    /**
     * @brief returns the line color.
     * @return A Color object. Invalid if the line format is invalid.
     *
     * This is a convenience method equivalent to `fill().color()`.
     */
    Color color() const;

    /**
     * @brief sets the line color irrespective of the line fill type.
     * @param color A Color object (rgb, hsl, scheme, system or preset
     * color).
     *
     * This is a convenience method equivalent to `fill().setColor(color)`.
     */
    void setColor(const Color &color);

    /**
     * @brief sets the line color irrespective of line fill type.
     * @param color the line color. This method assumes the color is RGB.
     *
     * This is a convenience method equivalent to `fill().setColor(color)`.
     */
    void setColor(const QColor &color); //assume color type is sRGB

    /**
     * @brief returns the fill parameters of the line.
     * @return a copy of the line fill.
     */
    FillFormat fill() const;
    /**
     * @brief sets the fill parameters to the line.
     * @param fill
     */
    void setFill(const FillFormat &fill);
    /**
     * @brief returns the fill parameters of the line.
     * @return a reference to the line fill.
     */
    FillFormat &fill();

    /**
     * @brief returns the line width in range [0..20116800 EMU] or [0..1584 pt]
     * @return valid Coordinate if the parameter was set, invalid Coordinate
     * otherwise.
     *
     * There is no default value for line width. If width is not set, it is
     * application-dependent.
     */
    Coordinate width() const;
    /**
     * @brief sets the line width in points.
     * @param widthInPt width in range [0..1584 pt]
     *
     * There is no default value for line width. If width is not set, it is
     * application-dependent.
     */
    void setWidth(double widthInPt);

    /**
     * @brief sets the line width in EMU (1 pt = 12700 EMU)
     * @param widthInEMU width in range [0..20116800 EMU]
     *
     * There is no default value for line width. If width is not set, it is
     * application-dependent.
     */
    void setWidthEMU(qint64 widthInEMU);
    /**
     * @brief returns the line's pen alignment with the path stroke.
     *
     * There is no default value for pen alignment. If pen alignment is not set,
     * it is application-dependent.
     */
    std::optional<PenAlignment> penAlignment() const;
    /**
     * @brief sets the line's pen alignment with the path stroke.
     * @param val A PenAlignment enum value.
     *
     * There is no default value for pen alignment. If pen alignment is not set,
     * it is application-dependent.
     */
    void setPenAlignment(PenAlignment val);
    /**
     * @brief returns the line's compound type.
     *
     * There is no default value for compound type. If compound type is not set,
     * it is application-dependent.
     */
    std::optional<CompoundLineType> compoundLineType() const;
    /**
     * @brief sets the line's compound type.
     * @param val A CompoundLineType enum value.
     *
     * There is no default value for compound type. If compound type is not set,
     * it is application-dependent.
     */
    void setCompoundLineType(CompoundLineType val);
    /**
     * @brief returns the line's stroke type.
     *
     * There is no default value for stroke type. If stroke type is not set,
     * it is application-dependent.
     */
    std::optional<StrokeType> strokeType() const;
    /**
     * @brief sets the line's stroke type.
     * @param val A StrokeType enum value.
     *
     * If @a val is StrokeType::Custom but no #dashPattern() was specified,
     * @a val is ignored. Use #setDashPattern() method instead.
     *
     * There is no default value for stroke type. If stroke type is not set,
     * it is application-dependent.
     */
    void setStrokeType(StrokeType val);
    /**
     * @brief returns the line's dash pattern.
     * @return A vector of widths of dashes and spaces by turns specified in
     * percents.
     */
    QVector<double> dashPattern() const;
    /**
     * @brief sets the line's dash pattern.
     * @param pattern A vector of widths of dashes and spaces by turns specified
     * in percents.
     *
     * A simple dash pattern may look like this:
     *
     * @code
     * QVector<double> customDashPattern {300,100};
     * series(0).setDashPattern(customDashPattern);
     * @endcode
     *
     * This method also sets stroke type to StrokeType::Custom.
     */
    void setDashPattern(QVector<double> pattern);
    /**
     * @brief returns the line cap style.
     *
     * There is no default value for line cap style. If line cap style is not
     * set, it is application-dependent.
     */
    std::optional<LineCap> lineCap() const;
    /**
     * @brief sets the line cap style.
     * @param val A LineCap enum value.
     *
     * There is no default value for line cap style. If line cap style is not
     * set, it is application-dependent.
     */
    void setLineCap(LineCap val);
    /**
     * @brief returns the line join style.
     *
     * There is no default value for line join style. If line join style is not
     * set, it is application-dependent.
     */
    std::optional<LineJoin> lineJoin() const;
    /**
     * @brief sets the line join style.
     * @param val A LineJoin enum value.
     *
     * There is no default value for line join style. If line join style is not
     * set, it is application-dependent.
     */
    void setLineJoin(LineJoin val);

    /**
     * @brief returns the amount by which lines are extended to form a miter
     * join - otherwise miter joins can extend infinitely far (for lines which
     * are almost parallel).
     * @return value in percents if the parameter is set. The value of 100
     * equals 100%.
     *
     * There is no default value for miter limit. If miter limit is not
     * set, it is application-dependent.
     */
    std::optional<double> miterLimit() const;
    /**
     * @brief sets the amount by which lines is extended to form a miter join -
     * otherwise miter joins can extend infinitely far (for lines which are
     * almost parallel).
     * @param val positive value in percents. The value of 100 equals 100%.
     *
     * There is no default value for miter limit. If miter limit is not
     * set, it is application-dependent.
     */
    void setMiterLimit(double val);
    /**
     * @brief returns the line end type.
     *
     * The default value is LineEndType::None.
     */
    std::optional<LineEndType> lineEndType();
    /**
     * @brief returns the line start type.
     *
     * The default value is LineEndType::None.
     */
    std::optional<LineEndType> lineStartType();
    /**
     * @brief sets the line end type.
     * @param val A LineEndType enum value.
     *
     * If no line end type is specified, LineEndType::None is assumed.
     */
    void setLineEndType(LineEndType val);
    /**
     * @brief sets the line start type.
     * @param val A LineEndType enum value.
     *
     * If no line start type is specified, LineEndType::None is assumed.
     */
    void setLineStartType(LineEndType val);
    /**
     * @brief returns the line end length.
     *
     * There is no default value for line end length. If line end length is not
     * set, it is application-dependent.
     */
    std::optional<LineEndSize> lineEndLength();
    /**
     * @brief returns the line start length.
     *
     * There is no default value for line start length. If line start length is
     * not set, it is application-dependent.
     */
    std::optional<LineEndSize> lineStartLength();
    /**
     * @brief sets the line end length.
     * @param val A LineEndSize enum value.
     *
     * There is no default value for line start length. If line start length is
     * not set, it is application-dependent.
     */
    void setLineEndLength(LineEndSize val);
    /**
     * @brief sets the line start length.
     * @param val A LineEndSize enum value.
     *
     * There is no default value for line start length. If line start length is
     * not set, it is application-dependent.
     */
    void setLineStartLength(LineEndSize val);
    /**
     * @brief returns the line end width.
     *
     * There is no default value for line end width. If line end width is
     * not set, it is application-dependent.
     */
    std::optional<LineEndSize> lineEndWidth();
    /**
     * @brief returns the line start width.
     *
     * There is no default value for line start width. If line start width is
     * not set, it is application-dependent.
     */
    std::optional<LineEndSize> lineStartWidth();
    /**
     * @brief sets the line end width.
     * @param val A LineEndSize enum value.
     *
     * There is no default value for line end width. If line end width is
     * not set, it is application-dependent.
     */
    void setLineEndWidth(LineEndSize val);
    /**
     * @brief sets the line start width.
     * @param val A LineEndSize enum value.
     *
     * There is no default value for line start width. If line start width is
     * not set, it is application-dependent.
     */
    void setLineStartWidth(LineEndSize val);
    /**
     * @brief returns whether the line format is valid.
     *
     * An invalid line format is also a default line format.
     */
    bool isValid() const;

    void write(QXmlStreamWriter &writer, const QString &name) const;
    void read(QXmlStreamReader &reader);

    bool operator == (const LineFormat &other) const;
    bool operator != (const LineFormat &other) const;

    operator QVariant() const;

    SERIALIZE_ENUM(StrokeType, {
        {StrokeType::Solid, "solid"},
        {StrokeType::SystemDash, "sysDash"},
        {StrokeType::SystemDot, "sysDot"},
        {StrokeType::Dash, "dash"},
        {StrokeType::Dot, "dot"},
        {StrokeType::DashDot, "dashDot"},
        {StrokeType::SystemDashDot, "sysDashDot"},
        {StrokeType::LongDash, "lgDash"},
        {StrokeType::LongDashDot, "lgDashDot"},
        {StrokeType::LongDashDotDot, "lgDashDotDot"},
        {StrokeType::SystemDashDotDot, "sysDashDotDot"},
    });
    SERIALIZE_ENUM(LineJoin, {
        {LineJoin::Round, "round"},
        {LineJoin::Bevel, "bevel"},
        {LineJoin::Miter, "miter"},
    });

    SERIALIZE_ENUM(LineCap,  {
        {LineCap::Square, "sq"},
        {LineCap::Round, "rnd"},
        {LineCap::Flat, "flat"},
    });

    SERIALIZE_ENUM(LineEndType,  {
        {LineEndType::None, "none"},
        {LineEndType::Triangle, "triangle"},
        {LineEndType::Stealth, "stealth"},
        {LineEndType::Oval, "oval"},
        {LineEndType::Arrow, "arrow"},
        {LineEndType::Diamond, "diamond"},
    });

    SERIALIZE_ENUM(LineEndSize,  {
        {LineEndSize::Small, "sm"},
        {LineEndSize::Medium, "med"},
        {LineEndSize::Large, "lg"},
    });

    SERIALIZE_ENUM(PenAlignment,
    {
        {PenAlignment::Center, "ctr"},
        {PenAlignment::Inset, "in"},
    });
    SERIALIZE_ENUM(CompoundLineType, {
        {CompoundLineType::Single, "sng"},
        {CompoundLineType::Double, "dbl"},
        {CompoundLineType::ThickThin, "thickThin"},
        {CompoundLineType::ThinThick, "thinThick"},
        {CompoundLineType::Triple, "tri"}
    });
private:
    void readDashPattern(QXmlStreamReader &reader);
#ifndef QT_NO_DEBUG_STREAM
    friend QDebug operator<<(QDebug, const LineFormat &f);
#endif
    QSharedDataPointer<LineFormatPrivate> d;
};

#ifndef QT_NO_DEBUG_STREAM
QDebug operator<<(QDebug dbg, const LineFormat &f);
#endif

}

Q_DECLARE_METATYPE(QXlsx::LineFormat)
Q_DECLARE_TYPEINFO(QXlsx::LineFormat, Q_MOVABLE_TYPE);

#endif
