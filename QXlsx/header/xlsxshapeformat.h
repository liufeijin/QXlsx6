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

class QXLSX_EXPORT ShapeFormat
{
public:
    enum class BlackWhiteMode {
        Clear, // "clr"
        Auto, // "auto"
        Gray, // "gray"
        LightGray, // "ltGray"
        InverseGray, // "invGray"
        GrayWhite, // "grayWhite"
        BlackGray, // "blackGray"
        BlackWhite, // "blackWhite"
        Black, // "black"
        White, // "white"
        Hidden, // "hidden
    };

    /**
     * @brief creates a rectangular shape with default line and fill
     */
    ShapeFormat();
    //TODO: add methods to create a shape with line, fill, shape
    ShapeFormat(const ShapeFormat &other);
    ~ShapeFormat();
    ShapeFormat &operator=(const ShapeFormat &other);

    std::optional<ShapeFormat::BlackWhiteMode> blackWhiteMode() const;
    void setBlackWhiteMode(ShapeFormat::BlackWhiteMode val);

    std::optional<Transform2D> transform2D() const;
    void setTransform2D(Transform2D val);

    std::optional<PresetGeometry2D> presetGeometry() const;
    void setPresetGeometry(PresetGeometry2D val);
    void setPresetGeometry(ShapeType shapeType);

    FillFormat &fill();
    FillFormat fill() const;
    void setFill(const FillFormat &val);

    LineFormat line() const;
    LineFormat &line();
    void setLine(const LineFormat &line);

    Effect effectList() const;
    void setEffectList(const Effect &effectList);

    std::optional<Scene3D> scene3D() const;
    void setScene3D(const Scene3D scene3D);

    std::optional<Shape3D> shape3D() const;
    void setShape3D(const Shape3D &shape3D);

    bool isValid() const;

    void write(QXmlStreamWriter &writer, const QString &name) const;
    void read(QXmlStreamReader &reader);
    QList<std::reference_wrapper<FillFormat>> fills();

    bool operator == (const ShapeFormat &other) const;
    bool operator != (const ShapeFormat &other) const;

private:
    SERIALIZE_ENUM(BlackWhiteMode, {
        {BlackWhiteMode::Clear, "clr"},
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
