#ifndef XLSXFILLFORMAT_H
#define XLSXFILLFORMAT_H

#include <QFont>
#include <QColor>
#include <QByteArray>
#include <QList>
#include <QVariant>
#include <QBrush>
#include <QXmlStreamWriter>
#include <QXmlStreamReader>
#include <QSharedData>

#include "xlsxglobal.h"
#include "xlsxcolor.h"
#include "xlsxmain.h"
#include "xlsxutility_p.h"

namespace QXlsx {

class FillFormatPrivate;

/**
 * @brief Sets fill properties for lines, shapes etc.
 */
class QXLSX_EXPORT FillFormat
{
public:
    /**
     * @brief The FillType enum specifies the fill type for lines and shapes.
     *
     */
    enum class FillType {
        NoFill, /**< @brief line or shape is not filled (not drawn) */
        SolidFill, /**< @brief line or shape has one color */
        GradientFill, /**< @brief line or shape is drawn with gradient */
        BlipFill, /**< @brief shape is filled with a picture (not applicable to lines) */
        PatternFill, /**< @brief line or shape is filled with a predefined pattern */
        GroupFill  /**< @brief shape inherits the fill properties of a group (not applicable to lines)*/
    };
    enum class PathShadeType
    {
        Shape,
        Circle,
        Rectangle
    };
    enum class TileFlipMode
    {
        None,
        X,
        Y,
        XY
    };
    enum class PatternType
    {
        Percent5,
        Percent10,
        Percent20,
        Percent25,
        Percent30,
        Percent40,
        Percent50,
        Percent60,
        Percent70,
        Percent75,
        Percent80,
        Percent90,
        Horizontal,
        Vertical,
        LightHorizontal,
        LightVertical,
        DarkHorizontal,
        DarkVertical,
        NarrowHorizontal,
        NarrowVertical,
        DashedHorizontal,
        DashedVertical,
        Cross,
        DownwardDiagonal,
        UpwardDiagonal,
        LightDownwardDiagonal,
        LightUpwardDiagonal,
        DarkDownwardDiagonal,
        DarkUpwardDiagonal,
        WideDownwardDiagonal,
        WideUpwardDiagonal,
        DashedDownwardDiagonal,
        DashedUpwardDiagonal,
        DiagonalCross,
        SmallCheckerBoard,
        LargeCheckerBoard,
        SmallGrid,
        LargeGrid,
        DottedGrid,
        SmallConfetti,
        LargeConfetti,
        HorizontalBrick,
        DiagonalBrick,
        SolidDiamond,
        OpenDiamond,
        DottedDiamond,
        Plaid,
        Sphere,
        Weave,
        Divot,
        Shingle,
        Wave,
        Trellis,
        ZigZag,
    };

    FillFormat();
    explicit FillFormat(FillType type);
    FillFormat(const QBrush &brush);
    FillFormat(const FillFormat &other);
    FillFormat &operator=(const FillFormat &rhs);
    ~FillFormat();

    FillType type() const;
    void setType(FillType type);

    void setBrush(const QBrush &brush);
    QBrush toBrush() const;

    /* Solid fill properties */
    /**
     * @brief returns the solid fill color
     * @return
     */
    Color color() const;
    /**
     * @brief sets the color for solid fill
     * @param color
     */
    void setColor(const Color &color);
    /**
     * @brief sets an rgb color for solid fill
     * @param color
     */
    void setColor(const QColor &color);

    /* Gradient fill properties */

    /**
     * @brief returns the gradients list for the gradient fill
     * @return QMap, where keys are color stops in percents, values are gradient colors.
     */
    QMap<double, Color> gradientList() const;
    /**
     * @brief setGradientList sets the gradient list for the gradient fill
     * @param list a map with keys being the color stops (in percents), values being the colors.
     */
    void setGradientList(const QMap<double, Color> &list);
    /**
     * @brief adds a gradient stop to the gradient list.
     * @param stop percentage value (100.0 is 100%).
     * @param color Color.
     */
    void addGradientStop(double stop, const Color &color);
    /**
     * @brief returns the direction of color change for the linear gradient.
     * @return valid Angle if the parameter is set.
     *
     * The direction of color change is measured clockwise starting from the horizontal.
     */
    Angle linearShadeAngle() const;
    /**
     * @brief sets the direction of color change for the linear gradient.
     * @param val the direction measured clockwise starting from the horizontal.
     */
    //TODO: test and check clockwise or counterclockwise
    void setLinearShadeAngle(Angle val);
    /**
     * @brief returns whether the linear gradient angle scales with the fill region.
     * If set to true, the linearShadeAngle() is scaled to the shape's fill region.
     * For example, gradient with an angle of 45° (i.e. vector (1, -1)) in a shape with width of 300 and height of 200
     * will be scaled to vector(300,-200), that is an angle of 33.69°.
     * @return
     */
    std::optional<bool> linearShadeScaled() const;
    /**
     * @brief sets whether the linear gradient angle scales with the fill region.
     * @param scaled if true, then the linearShadeAngle() will be scaled to the shape's fill region.
     * For example, gradient with an angle of 45° (i.e. vector (1, -1)) in a shape with width of 300 and height of 200
     * will be scaled to vector(300,-200), that is an angle of 33.69°.
     * If false, the gradient angle is independent of the shape's fill region.
     */
    void setLinearShadeScaled(bool scaled);

    /**
     * @brief returns the type of the path gradient.
     * @return
     */
    std::optional<PathShadeType> pathShadeType() const;
    void setPathShadeType(PathShadeType pathShadeType);

    /**
     * @brief returns the "focus" rectangle for the center shade, specified
     * relative to the fill tile rectangle. The center shade fills the entire
     * tile except for the margins specified by pathShadeRect.
     * @return
     */
    std::optional<RelativeRect> pathShadeRect() const;
    /**
     * @brief sets the "focus" rectangle for the center shade, specified
     * relative to the fill tile rectangle. The center shade fills the entire
     * tile except for the margins specified by rect.
     * @param rect
     */
    void setPathShadeRect(RelativeRect rect);

    /**
     * @brief returns a rectangular region of the shape to which the gradient
     * is applied.  This region is then tiled across the remaining area of the
     * shape to complete the fill.  The tile rectangle is defined by percentage
     * offsets from the sides of the shape's bounding box.

     * @return valid RelativeRect if the parameter is set, nullopt otherwise.
     */
    std::optional<RelativeRect> tileRect() const;
    /**
     * @brief sets a rectangular region of the shape to which the gradient
     * is applied. This region is then tiled across the remaining area of the
     * shape to complete the fill. The tile rectangle is defined by percentage
     * offsets from the sides of the shape's bounding box.
     *
     * This parameter is applicable to both linear and path gradients.
     *
     * @param rect
     */
    void setTileRect(RelativeRect rect);
    /**
     * @brief tileFlipMode
     * @return
     */
    std::optional<FillFormat::TileFlipMode> tileFlipMode() const;
    void setTileFlipMode(TileFlipMode tileFlipMode);

    std::optional<bool> rotateWithShape() const;
    void setRotateWithShape(bool val);

    /* Pattern fill properties */

    Color foregroundColor() const;
    void setForegroundColor(const Color &color);

    Color backgroundColor() const;
    void setBackgroundColor(const Color &color);

    std::optional<PatternType> patternType();
    void setPatternType(PatternType patternType);

    bool isValid() const;

    void write(QXmlStreamWriter &writer) const;
    void read(QXmlStreamReader &reader);

    bool operator==(const FillFormat &other) const;
    bool operator!=(const FillFormat &other) const;

    operator QVariant() const;

private:
    SERIALIZE_ENUM(FillType, {
        {FillType::NoFill, "noFill"},
        {FillType::SolidFill, "solidFill"},
        {FillType::GradientFill, "gradFill"},
        {FillType::BlipFill, "blipFill"},
        {FillType::PatternFill, "pattFill"},
        {FillType::GroupFill, "grpFill"},
    });
    SERIALIZE_ENUM(PathShadeType, {
        {PathShadeType::Shape, "shape"},
        {PathShadeType::Circle, "circle"},
        {PathShadeType::Rectangle, "rect"},
    });
    SERIALIZE_ENUM(TileFlipMode, {
        {TileFlipMode::None, "none"},
        {TileFlipMode::X, "x"},
        {TileFlipMode::Y, "y"},
        {TileFlipMode::XY, "xy"},
    });
    SERIALIZE_ENUM(PatternType, {
        {PatternType::Percent5, "pct5"},
        {PatternType::Percent10, "pct10"},
        {PatternType::Percent20, "pct20"},
        {PatternType::Percent25, "pct25"},
        {PatternType::Percent30, "pct30"},
        {PatternType::Percent40, "pct40"},
        {PatternType::Percent50, "pct50"},
        {PatternType::Percent60, "pct60"},
        {PatternType::Percent70, "pct70"},
        {PatternType::Percent75, "pct75"},
        {PatternType::Percent80, "pct80"},
        {PatternType::Percent90, "pct90"},
        {PatternType::Horizontal, "horz"},
        {PatternType::Vertical, "vert"},
        {PatternType::LightHorizontal, "ltHorz"},
        {PatternType::LightVertical, "ltVert"},
        {PatternType::DarkHorizontal, "dkHorz"},
        {PatternType::DarkVertical, "dkVert"},
        {PatternType::NarrowHorizontal, "narHorz"},
        {PatternType::NarrowVertical, "narVert"},
        {PatternType::DashedHorizontal, "dashHorz"},
        {PatternType::DashedVertical, "dashVert"},
        {PatternType::Cross, "cross"},
        {PatternType::DownwardDiagonal, "dnDiag"},
        {PatternType::UpwardDiagonal, "upDiag"},
        {PatternType::LightDownwardDiagonal, "ltDnDiag"},
        {PatternType::LightUpwardDiagonal, "ltUpDiag"},
        {PatternType::DarkDownwardDiagonal, "dkDnDiag"},
        {PatternType::DarkUpwardDiagonal, "dkUpDiag"},
        {PatternType::WideDownwardDiagonal, "wdDnDiag"},
        {PatternType::WideUpwardDiagonal, "wdUpDiag"},
        {PatternType::DashedDownwardDiagonal, "dashDnDiag"},
        {PatternType::DashedUpwardDiagonal, "dashUpDiag"},
        {PatternType::DiagonalCross, "diagCross"},
        {PatternType::SmallCheckerBoard, "smCheck"},
        {PatternType::LargeCheckerBoard, "lgCheck"},
        {PatternType::SmallGrid, "smGrid"},
        {PatternType::LargeGrid, "lgGrid"},
        {PatternType::DottedGrid, "dotGrid"},
        {PatternType::SmallConfetti, "smConfetti"},
        {PatternType::LargeConfetti, "lgConfetti"},
        {PatternType::HorizontalBrick, "horzBrick"},
        {PatternType::DiagonalBrick, "diagBrick"},
        {PatternType::SolidDiamond, "solidDmnd"},
        {PatternType::OpenDiamond, "openDmnd"},
        {PatternType::DottedDiamond, "dotDmnd"},
        {PatternType::Plaid, "plaid"},
        {PatternType::Sphere, "sphere"},
        {PatternType::Weave, "weave"},
        {PatternType::Divot, "divot"},
        {PatternType::Shingle, "shingle"},
        {PatternType::Wave, "wave"},
        {PatternType::Trellis, "trellis"},
        {PatternType::ZigZag, "zigZag"},
    });
    friend QDebug operator<<(QDebug, const FillFormat &f);
    QSharedDataPointer<FillFormatPrivate> d;
    void readNoFill(QXmlStreamReader &reader);
    void readSolidFill(QXmlStreamReader &reader);
    void readGradientFill(QXmlStreamReader &reader);
    void readPatternFill(QXmlStreamReader &reader);
    void readGroupFill(QXmlStreamReader &reader);

    void writeNoFill(QXmlStreamWriter &writer) const;
    void writeSolidFill(QXmlStreamWriter &writer) const;
    void writeGradientFill(QXmlStreamWriter &writer) const;
    void writePatternFill(QXmlStreamWriter &writer) const;
    void writeGroupFill(QXmlStreamWriter &writer) const;

    void readGradientList(QXmlStreamReader &reader);
    void writeGradientList(QXmlStreamWriter &writer) const;
};

QDebug operator<<(QDebug dbg, const FillFormat &f);

}

Q_DECLARE_TYPEINFO(QXlsx::FillFormat, Q_MOVABLE_TYPE);
Q_DECLARE_METATYPE(QXlsx::FillFormat)

#endif // XLSXFillformat_H
