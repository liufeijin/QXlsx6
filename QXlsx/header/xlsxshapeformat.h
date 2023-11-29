#ifndef XLSXSHAPEPROPERTIES_H
#define XLSXSHAPEPROPERTIES_H

#include <QFont>
#include <QColor>
#include <QByteArray>
#include <QList>
#include <QVariant>
#include <QXmlStreamWriter>
#include <QXmlStreamReader>
#include <QSharedData>
#include <QDebug>

#include "xlsxglobal.h"
#include "xlsxmain.h"
#include "xlsxfillformat.h"
#include "xlsxlineformat.h"
#include "xlsxeffect.h"

namespace QXlsx {

class ShapePrivate;

/**
 * @brief The ShapeFormat class represents parameters of a shape on a chart,
 * a part of a chart etc.
 *
 * Each shape has:
 *
 * - a border, which is accessible via #line(), #setLine() methods.
 * - a fill, which is accessible via #fill(), #setFill() methods.
 * - an effects list (#effectList(), #setEffectList()).
 * - 3D parameters if needed (#scene3D(), #setScene3D(), #shape3D(),
 * #setShape3D()).
 * - an additional transformation (#transform2D(), #setTransform2D()).
 * - a preset geometry specified with #presetGeometry() and
 * #setPresetGeometry() (custom geometry is currently not supported).
 *
 * In most of the cases there is no need to specify the geometry of the shape,
 * as this class is mostly used to set the chart parts line and fill.
 *
 * Adding user-specified shapes on the chart canvas is currently not
 * implemented.
 *
 * See [Charts](../../examples/Charts) and [Chartsheets](../../examples/Chartsheets)
 * examples.
 *
 */
class QXLSX_EXPORT ShapeFormat
{
public:
    /**
     * @brief The BlackWhiteMode enum specifies how the shape fill should be
     * rendered in the black-and-white mode.
     */
    enum class BlackWhiteMode {
        Color, /**< Fill is rendered with normal coloring. */
        Auto, /**< Fill is rendered with automatic coloring. */
        Gray, /**< Fill is rendered with gray coloring. */
        LightGray, /**< Fill is rendered with light gray coloring. */
        InverseGray, /**< Fill is rendered with inverse gray coloring. */
        GrayWhite, /**< Fill is rendered with gray and white coloring. */
        BlackGray, /**< Fill is rendered with black and gray coloring. */
        BlackWhite, /**< Fill is rendered with black and white coloring. */
        Black, /**< Fill is rendered with black only coloring. */
        White, /**< Fill is rendered with white only coloring. */
        Hidden, /**< Fill is rendered with hidden coloring. */
    };

    /**
     * @brief creates an invalid shape format with default line, shape and fill
     * parameters.
     */
    ShapeFormat();

    ShapeFormat(const ShapeFormat &other);
    ~ShapeFormat();
    ShapeFormat &operator=(const ShapeFormat &other);
    /**
     * @brief returns how to interpret color information contained within a
     * shape to achieve a color, black and white, or grayscale rendering of the
     * shape.
     *
     * There's no default value for this parameter.
     */
    std::optional<ShapeFormat::BlackWhiteMode> blackWhiteMode() const;
    /**
     * @brief sets how to interpret color information contained within a
     * shape to achieve a color, black and white, or grayscale rendering of the
     * shape.
     * @param val A BlackWhiteMode enum value.
     *
     * There's no default value for this parameter.
     */
    void setBlackWhiteMode(ShapeFormat::BlackWhiteMode val);
    /**
     * @brief returns the 2D transform parameters of the shape - offset,
     * rotation, flip etc.
     * @return A Transform2D object.
     */
    std::optional<Transform2D> transform2D() const; //TODO: convert to implicitly shareable and remove optional.
    /**
     * @brief sets the 2D transform parameters of the shape - offset,
     * rotation, flip etc.
     * @param val A Transform2D object.
     */
    void setTransform2D(Transform2D val);
    /**
     * @brief sets the rotation angle of the shape.
     * @param angle Angle in degrees or in fractions of degree.
     *
     * This is a convenience method, equivalent to
     *
     * @code
     * auto transform = shape().transform2D().value_or(Transform2D());
     * transform.rotation = angle;
     * shape().setTransform2D(transform);
     * @endcode
     */
    void setRotation(const Angle angle);
    /**
     * @brief returns the geometry of the shape.
     * @return A PresetGeometry2D object.
     */
    std::optional<PresetGeometry2D> presetGeometry() const; //TODO: convert to implicitly shareable and remove optional.
    /**
     * @brief sets the geometry of the shape.
     * @param val A PresetGeometry2D object.
     */
    void setPresetGeometry(PresetGeometry2D val);
    /**
     * @brief sets the geometry of the shape.
     * @param shapeType A ShapeType enum value.
     *
     * If you need to specify the additional geometry parameters, use
     * #setPresetGeometry(PresetGeometry2D val).
     */
    void setPresetGeometry(ShapeType shapeType);
    /**
     * @brief returns the fill parameters of the shape.
     */
    FillFormat &fill();
    /**
     * @brief returns the fill parameters of the shape.
     */
    FillFormat fill() const;
    /**
     * @brief sets the fill parameters of the shape.
     * @param val A FillFormat object.
     */
    void setFill(const FillFormat &val);
    /**
     * @brief returns the border parameters of the shape.
     */
    LineFormat line() const;
    /**
     * @brief returns the border parameters of the shape.
     */
    LineFormat &line();
    /**
     * @brief sets the border parameters of the shape.
     * @param line A LineFormat object.
     */
    void setLine(const LineFormat &line);
    /**
     * @brief returns the effects of the shape.
     */
    Effect effectList() const;
    /**
     * @brief sets the effects of the shape.
     * @param effectList An Effect object.
     */
    void setEffectList(const Effect &effectList);

    std::optional<Scene3D> scene3D() const;
    void setScene3D(const Scene3D scene3D);

    std::optional<Shape3D> shape3D() const;
    void setShape3D(const Shape3D &shape3D);

    bool isValid() const;

    void write(QXmlStreamWriter &writer, const QString &name) const;
    void read(QXmlStreamReader &reader);

    bool operator == (const ShapeFormat &other) const;
    bool operator != (const ShapeFormat &other) const;

private:
    friend class Chart;
    friend class Labels;

    QList<std::reference_wrapper<FillFormat>> fills();

    SERIALIZE_ENUM(BlackWhiteMode, {
        {BlackWhiteMode::Color, "clr"},
        {BlackWhiteMode::Auto, "auto"},
        {BlackWhiteMode::Gray, "gray"},
        {BlackWhiteMode::LightGray, "ltGray"},
        {BlackWhiteMode::InverseGray, "invGray"},
        {BlackWhiteMode::GrayWhite, "grayWhite"},
        {BlackWhiteMode::BlackGray, "blackGray"},
        {BlackWhiteMode::BlackWhite, "blackWhite"},
        {BlackWhiteMode::Black, "black"},
        {BlackWhiteMode::White, "white"},
        {BlackWhiteMode::Hidden, "hidden"},
    });
    friend QDebug operator<<(QDebug, const ShapeFormat &f);

    QSharedDataPointer<ShapePrivate> d;
};

QDebug operator<<(QDebug dbg, const ShapeFormat &f);

}

Q_DECLARE_METATYPE(QXlsx::ShapeFormat)
Q_DECLARE_TYPEINFO(QXlsx::ShapeFormat, Q_MOVABLE_TYPE);

#endif // XLSXSHAPEPROPERTIES_H
