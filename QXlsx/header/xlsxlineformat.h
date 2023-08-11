// xlsxformat.h

#ifndef QXLSX_LINEFORMAT_H
#define QXLSX_LINEFORMAT_H

#include <QFont>
#include <QColor>
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

QT_BEGIN_NAMESPACE_XLSX

class Worksheet;
class WorksheetPrivate;
class RichStringPrivate;
class SharedStrings;
class LineFormatPrivate;

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

    enum class StrokeType {
        Solid,
        Dot,
        RoundDot,
        Dash,
        DashDot,
        LongDash,
        LongDashDot,
        LongDashDotDot,
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

    LineFormat(); //no fill
    LineFormat(FillFormat::FillType fill, double widthInPt, QColor color); // quick create
    LineFormat(FillFormat::FillType fill, qint64 widthInEMU, QColor color); // quick create
    LineFormat(const LineFormat &other);
    LineFormat &operator=(const LineFormat &rhs);
    ~LineFormat();

    FillFormat::FillType type() const;
    void setType(FillFormat::FillType type);

    /**
     * @brief color
     * @return line color if line fill type is SolidFill,
     * invalid color otherwise;
     */
    Color color() const;

    /**
     * @brief setColor overload, sets line color irrespective of line fill type
     * @param color color specification (rgb, hsl, scheme, system or preset color)
     */
    void setColor(const Color &color);

    /**
     * @brief setColor overload, sets line color irrespective of line fill type
     * @param color the line color
     */
    void setColor(const QColor &color); //assume color type is sRGB

    FillFormat fill() const;
    void setFill(const FillFormat &fill);
    FillFormat &fill();

    /**
     * @brief width returns width in range [0..20116800 EMU] or [0..1584 pt]
     * @return optional value
     */
    Coordinate width() const;
    /**
     * @brief setWidth sets line width in points
     * @param widthInPt width in range [0..1584 pt]
     */
    void setWidth(double widthInPt);

    /**
     * @brief setWidth sets line width in EMU
     * @param widthInEMU width in range [0..20116800 EMU]
     */
    void setWidthEMU(qint64 widthInEMU);

    std::optional<PenAlignment> penAlignment() const;
    void setPenAlignment(PenAlignment val);

    std::optional<CompoundLineType> compoundLineType() const;
    void setCompoundLineType(CompoundLineType val);

    std::optional<StrokeType> strokeType() const;
    void setStrokeType(StrokeType val);

    LineCap lineCap() const;
    void setLineCap(LineCap val);

    std::optional<LineJoin> lineJoin() const;
    void setLineJoin(LineJoin val);

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

private:
    SERIALIZE_ENUM(CompoundLineType, {
        {CompoundLineType::Single, "sng"},
        {CompoundLineType::Double, "dbl"},
        {CompoundLineType::ThickThin, "thickThin"},
        {CompoundLineType::ThinThick, "thinThick"},
        {CompoundLineType::Triple, "tri"}
    });

    SERIALIZE_ENUM(StrokeType, {
        {StrokeType::Solid, "solid"},
        {StrokeType::Dot, "sysDash"},
        {StrokeType::RoundDot, "sysDot"},
        {StrokeType::Dash, "dash"},
        {StrokeType::DashDot, "dashDot"},
        {StrokeType::LongDash, "lgDash"},
        {StrokeType::LongDashDot, "lgDashDot"},
        {StrokeType::LongDashDotDot, "lgDashDotDot"},
    });

    friend QDebug operator<<(QDebug, const LineFormat &f);

    QSharedDataPointer<LineFormatPrivate> d;
};

QDebug operator<<(QDebug dbg, const LineFormat &f);

QT_END_NAMESPACE_XLSX

Q_DECLARE_METATYPE(QXlsx::LineFormat)

#endif
