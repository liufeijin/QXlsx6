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
 * @brief The LineFormat class represents stroke parameters that are used to draw
 * borders of shapes, axes, grid lines etc. on a chart.
 *
 * The class is _implicitly shareable_: the deep copy occurs only in the non-const methods.
 */
class QXLSX_EXPORT LineFormat
{
public:
    enum class CompoundLineType {
        Single,
        Double,
        ThickThin,
        ThinThick,
        Triple,
    };
    /**
     * @brief The StrokeType enum  represents line dash values
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

    enum class LineJoin {
        Round,
        Bevel,
        Miter
    };

    enum class LineCap {
        Square,
        Round,
        Flat,
    };

    enum class LineEndType {
        None,
        Triangle,
        Stealth,
        Oval,
        Arrow,
        Diamond
    };

    enum class LineEndSize {
        Small,
        Medium,
        Large
    };

    enum class PenAlignment
    {
        Center,
        Inset
    };
    /**
     * @brief creates an invalid line format.
     */
    LineFormat(); //no fill
    /**
     * @brief creates a new line format with specified fill, width and color.
     * @param fill line fill.
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

    FillFormat::FillType type() const;
    void setType(FillFormat::FillType type);

    /**
     * @brief color
     * @return line color.
     */
    Color color() const;

    /**
     * @brief sets line color irrespective of line fill type.
     * @param color color specification (rgb, hsl, scheme, system or preset color)
     */
    void setColor(const Color &color);

    /**
     * @brief sets line color irrespective of line fill type.
     * @param color the line color. This method assumes the color is RGB.
     */
    void setColor(const QColor &color); //assume color type is sRGB

    /**
     * @brief returns fill parameters of the line.
     * @return a copy of the line fill.
     */
    FillFormat fill() const;
    /**
     * @brief sets fill parameters to the line.
     * @param fill
     */
    void setFill(const FillFormat &fill);
    /**
     * @brief returns fill parameters of the line.
     * @return a reference to the line fill.
     */
    FillFormat &fill();

    /**
     * @brief width returns width in range [0..20116800 EMU] or [0..1584 pt]
     * @return valid Coordinate if the parameter was set, invalid Coordinate otherwise.
     */
    Coordinate width() const;
    /**
     * @brief setWidth sets line width in points
     * @param widthInPt width in range [0..1584 pt]
     */
    void setWidth(double widthInPt);

    /**
     * @brief setWidth sets line width in EMU (1 pt = 12700 EMU)
     * @param widthInEMU width in range [0..20116800 EMU]
     */
    void setWidthEMU(qint64 widthInEMU);

    std::optional<PenAlignment> penAlignment() const;
    void setPenAlignment(PenAlignment val);

    std::optional<CompoundLineType> compoundLineType() const;
    void setCompoundLineType(CompoundLineType val);

    std::optional<StrokeType> strokeType() const;
    void setStrokeType(StrokeType val);
    QVector<double> dashPattern() const;
    void setDashPattern(QVector<double> pattern);

    LineCap lineCap() const;
    void setLineCap(LineCap val);

    std::optional<LineJoin> lineJoin() const;
    void setLineJoin(LineJoin val);

    /**
     * @brief returns the amount by which lines is extended to form a miter join -
     * otherwise miter joins can extend infinitely far (for lines which are almost parallel).
     * @return value in percents if the parameter is set. The value of 100 equals 100%.
     */
    std::optional<double> miterLimit() const;
    /**
     * @brief sets the amount by which lines is extended to form a miter join -
     * otherwise miter joins can extend infinitely far (for lines which are almost parallel).
     * @param val positive value in percents. The value of 100 equals 100%.
     */
    void setMiterLimit(double val);

    std::optional<LineEndType> lineEndType();
    std::optional<LineEndType> lineStartType();
    void setLineEndType(LineEndType val);
    void setLineStartType(LineEndType val);

    std::optional<LineEndSize> lineEndLength();
    std::optional<LineEndSize> lineStartLength();
    void setLineEndLength(LineEndSize val);
    void setLineStartLength(LineEndSize val);

    std::optional<LineEndSize> lineEndWidth();
    std::optional<LineEndSize> lineStartWidth();
    void setLineEndWidth(LineEndSize val);
    void setLineStartWidth(LineEndSize val);

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
