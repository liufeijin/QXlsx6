#ifndef XLSXEFFECT_H
#define XLSXEFFECT_H

#include "xlsxglobal.h"

#include <QSharedData>
#include <QXmlStreamReader>
#include <QXmlStreamWriter>
#include <QDebug>

#include "xlsxutility_p.h"
#include "xlsxmain.h"
#include "xlsxfillformat.h"

namespace QXlsx {

class EffectPrivate;

/**
 * @brief The Effect class represents effects for shapes.
 */
class Effect
{
public:
    enum class Type
    {
        List,
        DAG
    };
    enum class FillBlendMode
    {
        Overlay,
        Multiply,
        Screen,
        Darken,
        Lighten
    };
    enum class Alignment
    {
        AlignTopLeft,
        AlignTop,
        AlignTopRight,
        AlignRight,
        AlignCenter,
        AlignLeft,
        AlignBottomLeft,
        AlignBottom,
        AlignBottomRight,
    };

    Effect();
    explicit Effect(Type type);
    Effect(const Effect &other);
    ~Effect();
    Effect &operator=(const Effect &other);

    bool operator == (const Effect &other) const;
    bool operator != (const Effect &other) const;

    bool isValid() const;

    /**
     * @brief type returns the type of the effect list
     * @return List if it is a simple effect list with fixed order of applying,
     * DAG if it is a directed acyclic graph of effects (not implemented)
     */
    Type type() const;
    /**
     * @brief setType sets the type of the effect list
     * @param type List if it is a simple effect list with fixed order of applying,
     * DAG if it is a directed acyclic graph of effects (not implemented)
     */
    void setType(Type type);

    /**
     * @brief blurRadius returns the radius of blur effect applied to the entire shape
     * @return
     */
    Coordinate blurRadius() const;
    /**
     * @brief setBlurRadius sets the radius of blur effect applied to the entire shape
     * @param newBlurRadius
     */
    void setBlurRadius(const Coordinate &newBlurRadius);
    /**
     * @brief blurGrow returns whether the bounds of the object should be grown as a result of the blurring
     * @return
     */
    std::optional<bool> blurGrow() const;
    /**
     * @brief setBlurGrow specifies whether the bounds of the object should be grown as a result of the blurring
     * @param newBlurGrow
     */
    void setBlurGrow(bool newBlurGrow);
    /**
     * @brief fillOverlay returns the additional fill to the object
     * @return
     */
    FillFormat fillOverlay() const;
    /**
     * @brief setFillOverlay specifies an additional fill to the object
     * @param newFillOverlay
     */
    void setFillOverlay(const FillFormat &newFillOverlay);
    /**
     * @brief fillBlendMode returns how to blend the additional fill with the base effect
     * @return
     */
    Effect::FillBlendMode fillBlendMode() const;
    /**
     * @brief setFillBlendMode specifies how to blend the additional fill with the base effect
     * @param newFillBlendMode
     */
    void setFillBlendMode(FillBlendMode newFillBlendMode);
    /**
     * @brief glowColor returns the color for the glow effect
     * @return
     */
    Color glowColor() const;
    /**
     * @brief setGlowColor specifies a color for the glow effect
     * @param newGlowColor
     */
    void setGlowColor(const Color &newGlowColor);
    /**
     * @brief glowRadius returns the radius for the glow effect
     * @return
     */
    Coordinate glowRadius() const;
    /**
     * @brief setGlowRadius specifies the radius for the glow effect
     * @param newGlowRadius
     */
    void setGlowRadius(const Coordinate &newGlowRadius);

    Color innerShadowColor() const;
    void setInnerShadowColor(const Color &newInnerShadowColor);
    Coordinate innerShadowBlurRadius() const;
    void setInnerShadowBlurRadius(const Coordinate &newInnerShadowBlurRadius);
    Coordinate innerShadowOffset() const;
    void setInnerShadowOffset(const Coordinate &newInnerShadowOffset);
    std::optional<Angle> innerShadowDirection() const;
    void setInnerShadowDirection(Angle newInnerShadowDirection);
    Color outerShadowColor() const;
    void setOuterShadowColor(const Color &newOuterShadowColor);
    Coordinate outerShadowBlurRadius() const;
    void setOuterShadowBlurRadius(const Coordinate &newOuterShadowBlurRadius);
    Coordinate outerShadowOffset() const;
    void setOuterShadowOffset(const Coordinate &newOuterShadowOffset);
    std::optional<Angle> outerShadowDirection() const;
    void setOuterShadowDirection(Angle newOuterShadowDirection);
    std::optional<double> outerShadowHorizontalScalingFactor() const;
    void setOuterShadowHorizontalScalingFactor(double newOuterShadowHorizontalScalingFactor);
    std::optional<double> outerShadowVerticalScalingFactor() const;
    void setOuterShadowVerticalScalingFactor(double newOuterShadowVerticalScalingFactor);
    std::optional<Angle> outerShadowHorizontalSkewFactor() const;
    void setOuterShadowHorizontalSkewFactor(Angle newOuterShadowHorizontalSkewFactor);
    std::optional<Angle> outerShadowVerticalSkewFactor() const;
    void setOuterShadowVerticalSkewFactor(Angle newOuterShadowVerticalSkewFactor);
    std::optional<bool> outerShadowRotateWithShape() const;
    void setOuterShadowRotateWithShape(bool newOuterShadowRotateWithShape);
    std::optional<Effect::Alignment> outerShadowAlignment() const;
    void setOuterShadowAlignment(Alignment newOuterShadowAlignment);
    Color presetShadowColor() const;
    void setPresetShadowColor(const Color &newPresetShadowColor);
    Coordinate presetShadowOffset() const;
    void setPresetShadowOffset(const Coordinate &newPresetShadowOffset);
    std::optional<Angle> presetShadowDirection() const;
    void setPresetShadowDirection(Angle newPresetShadowDirection);
    /**
     * @brief presetShadow returns the preset shadow type from the range [1..20]
     * @return
     */
    std::optional<int> presetShadow() const;
    /**
     * @brief setPresetShadow specifies the preset shadow type
     * @param newPresetShadow an int from the range [1..20]
     */
    void setPresetShadow(int newPresetShadow);
    Coordinate reflectionBlurRadius() const;
    void setReflectionBlurRadius(const Coordinate &newReflectionBlurRadius);
    std::optional<double> reflectionStartOpacity() const;
    void setReflectionStartOpacity(double newReflectionStartOpacity);
    std::optional<double> reflectionStartPosition() const;
    void setReflectionStartPosition(double newReflectionStartPosition);
    std::optional<double> reflectionEndOpacity() const;
    void setReflectionEndOpacity(double newReflectionEndOpacity);
    std::optional<double> reflectionEndPosition() const;
    void setReflectionEndPosition(double newReflectionEndPosition);
    Coordinate reflectionShadowOffset() const;
    void setReflectionShadowOffset(const Coordinate &newReflectionShadowOffset);
    std::optional<Angle> reflectionGradientDirection() const;
    void setReflectionGradientDirection(Angle newReflectionGradientDirection);
    std::optional<Angle> reflectionOffsetDirection() const;
    void setReflectionOffsetDirection(Angle newReflectionOffsetDirection);
    std::optional<double> reflectionHorizontalScalingFactor() const;
    void setReflectionHorizontalScalingFactor(double newReflectionHorizontalScalingFactor);
    std::optional<double> reflectionVerticalScalingFactor() const;
    void setReflectionVerticalScalingFactor(double newReflectionVerticalScalingFactor);
    std::optional<Angle> reflectionHorizontalSkewFactor() const;
    void setReflectionHorizontalSkewFactor(Angle newReflectionHorizontalSkewFactor);
    std::optional<Angle> reflectionVerticalSkewFactor() const;
    void setReflectionVerticalSkewFactor(Angle newReflectionVerticalSkewFactor);
    std::optional<Effect::Alignment> reflectionShadowAlignment() const;
    void setReflectionShadowAlignment(Alignment newReflectionShadowAlignment);
    std::optional<bool> reflectionRotateWithShape() const;
    void setReflectionRotateWithShape(bool newReflectionRotateWithShape);
    Coordinate softEdgesBlurRadius() const;
    void setSoftEdgesBlurRadius(const Coordinate &newSoftEdgesBlurRadius);

    void read(QXmlStreamReader &reader);
    void write(QXmlStreamWriter &writer) const;
private:
    SERIALIZE_ENUM(FillBlendMode,
    {
        {FillBlendMode::Overlay, "over"},
        {FillBlendMode::Multiply, "mult"},
        {FillBlendMode::Screen, "screen"},
        {FillBlendMode::Darken, "darken"},
        {FillBlendMode::Lighten, "lighten"}
    });
    SERIALIZE_ENUM(Alignment,
    {
        {Alignment::AlignTopLeft, "tl"},
        {Alignment::AlignTop, "t"},
        {Alignment::AlignTopRight, "tr"},
        {Alignment::AlignRight, "r"},
        {Alignment::AlignCenter, "ctr"},
        {Alignment::AlignLeft, "l"},
        {Alignment::AlignBottomLeft, "bl"},
        {Alignment::AlignBottom, "b"},
        {Alignment::AlignBottomRight, "br"},
    });

    void readEffectList(QXmlStreamReader &reader);
    void writeEffectList(QXmlStreamWriter &writer) const;
    QSharedDataPointer<EffectPrivate> d;
    friend QDebug operator<<(QDebug dbg, const Effect &e);
};

QDebug operator<<(QDebug dbg, const Effect &e);


}

#endif // XLSXEFFECT_H
