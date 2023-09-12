#include "xlsxeffect.h"

namespace QXlsx {

class EffectPrivate : public QSharedData
{
public:
    EffectPrivate();
    EffectPrivate(const EffectPrivate &other);
    ~EffectPrivate();

    bool operator == (const EffectPrivate &other) const;

    Effect::Type type = Effect::Type::List; //list by default, DAG if setType()
    Coordinate blurRadius;
    std::optional<bool> blurGrow;

    FillFormat fillOverlay;
    Effect::FillBlendMode fillBlendMode;

    Color glowColor;
    Coordinate glowRadius;

    Color innerShadowColor;
    Coordinate innerShadowBlurRadius;
    Coordinate innerShadowOffset;
    Angle innerShadowDirection;


    Color outerShadowColor;
    Coordinate outerShadowBlurRadius;
    Coordinate outerShadowOffset;
    Angle outerShadowDirection;
    std::optional<double> outerShadowHorizontalScalingFactor; //in %
    std::optional<double> outerShadowVerticalScalingFactor; //in %
    Angle outerShadowHorizontalSkewFactor;
    Angle outerShadowVerticalSkewFactor;
    std::optional<bool> outerShadowRotateWithShape;
    std::optional<FillFormat::Alignment> outerShadowAlignment;

    Color presetShadowColor;
    Coordinate presetShadowOffset;
    Angle presetShadowDirection;
    std::optional<int> presetShadow;

    //reflection properties
    Coordinate reflectionBlurRadius;
    std::optional<double> reflectionStartOpacity;
    std::optional<double> reflectionStartPosition;
    std::optional<double> reflectionEndOpacity;
    std::optional<double> reflectionEndPosition;
    Coordinate reflectionShadowOffset;
    Angle reflectionGradientDirection;
    Angle reflectionOffsetDirection;
    std::optional<double> reflectionHorizontalScalingFactor; //in %
    std::optional<double> reflectionVerticalScalingFactor; //in %
    Angle reflectionHorizontalSkewFactor;
    Angle reflectionVerticalSkewFactor;
    std::optional<FillFormat::Alignment> reflectionShadowAlignment;
    std::optional<bool> reflectionRotateWithShape;

    Coordinate softEdgesBlurRadius;
};

Effect::Effect()
{

}

Effect::Effect(Type type)
{
    d = new EffectPrivate;
    d->type = type;
}

Effect::Effect(const Effect &other) : d(other.d)
{

}

Effect::~Effect()
{

}

Effect &QXlsx::Effect::operator=(const QXlsx::Effect &other)
{
    if (*this != other)
        d = other.d;
    return *this;
}

bool Effect::operator ==(const QXlsx::Effect &other) const
{
    if (d == other.d) return true;
    if (!d || !other.d) return false;
    return *this->d.constData() == *other.d.constData();
}

bool Effect::operator !=(const QXlsx::Effect &other) const
{
    return !(operator==(other));
}

bool Effect::isValid() const
{
    if (d) return true;
    return false;
}

void Effect::read(QXmlStreamReader &reader)
{
    const auto &name = reader.name();
    if (name == QLatin1String("effectLst")) {
        d = new EffectPrivate;
        d->type = Type::List;
        readEffectList(reader);
    }
    else if (name == QLatin1String("effectDag")) {
        //TODO: effects tree
//        d = new EffectPrivate;
//        d->type = Type::DAG;
    }
}

void Effect::write(QXmlStreamWriter &writer) const
{
    if (!d) return;

    switch (d->type) {
    case Type::List: writeEffectList(writer); break;
    case Type::DAG:
        break;
    }
}

void Effect::readEffectList(QXmlStreamReader &reader)
{
    const auto &name = reader.name();
    while (!reader.atEnd()) {
        auto token = reader.readNext();

        if (token == QXmlStreamReader::StartElement) {
            const auto &a = reader.attributes();
            if (reader.name() == QLatin1String("blur")) {
                parseAttribute(a, QLatin1String("rad"), d->blurRadius);
                parseAttributeBool(a, QLatin1String("grow"), d->blurGrow);
            }
            else if (reader.name() == QLatin1String("fillOverlay")) {
                FillBlendMode t;
                fromString(a.value(QLatin1String("blend")).toString(), t);
                d->fillBlendMode = t;
                reader.readNextStartElement();
                d->fillOverlay.read(reader);
            }
            else if (reader.name() == QLatin1String("glow")) {
                parseAttribute(a, QLatin1String("rad"), d->glowRadius);
                reader.readNextStartElement();
                d->glowColor.read(reader);
            }
            else if (reader.name() == QLatin1String("innerShdw")) {
                parseAttribute(a, QLatin1String("dir"), d->innerShadowDirection);
                parseAttribute(a, QLatin1String("dist"), d->innerShadowOffset);
                parseAttribute(a, QLatin1String("blurRad"), d->innerShadowBlurRadius);
                reader.readNextStartElement();
                d->innerShadowColor.read(reader);
            }
            else if (reader.name() == QLatin1String("outerShdw")) {
                parseAttribute(a, QLatin1String("dir"), d->outerShadowDirection);
                parseAttribute(a, QLatin1String("dist"), d->outerShadowOffset);
                parseAttribute(a, QLatin1String("blurRad"), d->outerShadowBlurRadius);
                parseAttributePercent(a, QLatin1String("sx"), d->outerShadowHorizontalScalingFactor);
                parseAttributePercent(a, QLatin1String("sy"), d->outerShadowVerticalScalingFactor);
                parseAttribute(a, QLatin1String("kx"), d->outerShadowHorizontalSkewFactor);
                parseAttribute(a, QLatin1String("ky"), d->outerShadowVerticalSkewFactor);
                parseAttributeBool(a, QLatin1String("rotWithShape"), d->outerShadowRotateWithShape);
                if (a.hasAttribute(QLatin1String("algn"))) {
                    FillFormat::Alignment t;
                    FillFormat::fromString(a.value(QLatin1String("algn")).toString(), t);
                    d->outerShadowAlignment = t;
                }
                if (reader.readNextStartElement())
                    d->outerShadowColor.read(reader);
            }
            else if (reader.name() == QLatin1String("prstShdw")) {
                parseAttribute(a, QLatin1String("dir"), d->presetShadowDirection);
                parseAttribute(a, QLatin1String("dist"), d->presetShadowOffset);
                parseAttributeInt(a, QLatin1String("prst"), d->presetShadow);
                reader.readNextStartElement();
                d->presetShadowColor.read(reader);
            }
            else if (reader.name() == QLatin1String("reflection")) {
                parseAttribute(a, QLatin1String("blurRad"), d->reflectionBlurRadius);
                parseAttributePercent(a, QLatin1String("stA"), d->reflectionStartOpacity);
                parseAttributePercent(a, QLatin1String("stPos"), d->reflectionStartPosition);
                parseAttributePercent(a, QLatin1String("endA"), d->reflectionStartOpacity);
                parseAttributePercent(a, QLatin1String("endPos"), d->reflectionStartPosition);
                parseAttribute(a, QLatin1String("dist"), d->reflectionShadowOffset);
                parseAttribute(a, QLatin1String("dir"), d->reflectionGradientDirection);
                parseAttribute(a, QLatin1String("fadeDir"), d->reflectionOffsetDirection);
                parseAttributePercent(a, QLatin1String("sx"), d->reflectionHorizontalScalingFactor);
                parseAttributePercent(a, QLatin1String("sy"), d->reflectionVerticalScalingFactor);
                parseAttribute(a, QLatin1String("kx"), d->reflectionHorizontalSkewFactor);
                parseAttribute(a, QLatin1String("ky"), d->reflectionVerticalSkewFactor);
                if (a.hasAttribute(QLatin1String("algn"))) {
                    FillFormat::Alignment t;
                    FillFormat::fromString(a.value(QLatin1String("algn")).toString(), t);
                    d->reflectionShadowAlignment = t;
                }
                parseAttributeBool(a, QLatin1String("rotWithShape"), d->reflectionRotateWithShape);
            }
            else if (reader.name() == QLatin1String("softEdge")) {
                parseAttribute(a, QLatin1String("rad"), d->softEdgesBlurRadius);
            }
            else reader.skipCurrentElement();
        }
        else if (token == QXmlStreamReader::EndElement && reader.name() == name)
            break;
    }
}

void Effect::writeEffectList(QXmlStreamWriter &writer) const
{
    writer.writeStartElement(QLatin1String("a:effectLst"));
//        <xsd:element name="outerShdw" type="CT_OuterShadowEffect" minOccurs="0" maxOccurs="1"/>
//        <xsd:element name="prstShdw" type="CT_PresetShadowEffect" minOccurs="0" maxOccurs="1"/>
//        <xsd:element name="reflection" type="CT_ReflectionEffect" minOccurs="0" maxOccurs="1"/>
//        <xsd:element name="softEdge" type="CT_SoftEdgesEffect" minOccurs="0" maxOccurs="1"/>
    if (d->blurGrow.has_value() || d->blurRadius.isValid()) {
        writer.writeEmptyElement(QLatin1String("a:blur"));
        if (d->blurGrow.has_value())
            writer.writeAttribute(QLatin1String("grow"), toST_Boolean(d->blurGrow.value()));
        if (d->blurRadius.isValid())
            writer.writeAttribute(QLatin1String("rad"), d->blurRadius.toString());
    }
    if (d->fillOverlay.isValid()) {
        writer.writeStartElement(QLatin1String("a:fillOverlay"));
        QString s;
        toString(d->fillBlendMode, s);
        writer.writeAttribute(QLatin1String("blend"), s);
        d->fillOverlay.write(writer);
        writer.writeEndElement();
    }
    if (d->glowColor.isValid() || d->glowRadius.isValid()) {
        writer.writeStartElement(QLatin1String("a:glow"));
        if (d->glowRadius.isValid())
            writer.writeAttribute(QLatin1String("rad"), d->glowRadius.toString());
        if (d->glowColor.isValid()) d->glowColor.write(writer);
        writer.writeEndElement();
    }
    if (d->innerShadowBlurRadius.isValid() || d->innerShadowColor.isValid()
        || d->innerShadowDirection.isValid() || d->innerShadowOffset.isValid()) {
        writer.writeStartElement(QLatin1String("a:innerShdw"));
        if (d->innerShadowDirection.isValid())
            writer.writeAttribute(QLatin1String("dir"), d->innerShadowDirection.toString());
        if (d->innerShadowBlurRadius.isValid())
            writer.writeAttribute(QLatin1String("blurRad"), d->innerShadowBlurRadius.toString());
        if (d->innerShadowOffset.isValid())
            writer.writeAttribute(QLatin1String("dist"), d->innerShadowOffset.toString());
        if (d->innerShadowColor.isValid()) d->innerShadowColor.write(writer);
        writer.writeEndElement();
    }
    if (d->outerShadowAlignment.has_value() || d->outerShadowBlurRadius.isValid()
        || d->outerShadowColor.isValid() || d->outerShadowDirection.isValid()
        || d->outerShadowHorizontalScalingFactor.has_value() || d->outerShadowVerticalScalingFactor.has_value()
        || d->outerShadowHorizontalSkewFactor.isValid() || d->outerShadowVerticalSkewFactor.isValid()
        || d->outerShadowOffset.isValid() || d->outerShadowRotateWithShape.has_value()) {
        writer.writeStartElement(QLatin1String("a:outerShdw"));
        if (d->outerShadowDirection.isValid())
            writer.writeAttribute(QLatin1String("dir"), d->outerShadowDirection.toString());
        if (d->outerShadowBlurRadius.isValid())
            writer.writeAttribute(QLatin1String("blurRad"), d->outerShadowBlurRadius.toString());
        if (d->outerShadowOffset.isValid())
            writer.writeAttribute(QLatin1String("dist"), d->outerShadowOffset.toString());
        if (d->outerShadowAlignment.has_value()) {
            QString s; FillFormat::toString(d->outerShadowAlignment.value(), s);
            writer.writeAttribute(QLatin1String("algn"), s);
        }
        if (d->outerShadowRotateWithShape.has_value())
            writer.writeAttribute(QLatin1String("rotWithShape"), toST_Boolean(d->outerShadowRotateWithShape.value()));
        writeAttributePercent(writer, QLatin1String("sx"), d->outerShadowHorizontalScalingFactor);
        writeAttributePercent(writer, QLatin1String("sy"), d->outerShadowVerticalScalingFactor);
        if (d->outerShadowHorizontalSkewFactor.isValid())
            writer.writeAttribute(QLatin1String("kx"), d->outerShadowHorizontalSkewFactor.toString());
        if (d->outerShadowVerticalSkewFactor.isValid())
            writer.writeAttribute(QLatin1String("ky"), d->outerShadowVerticalSkewFactor.toString());
        if (d->outerShadowColor.isValid()) d->outerShadowColor.write(writer);
        writer.writeEndElement();
    }
    if (d->presetShadow.has_value() || d->presetShadowColor.isValid()
        || d->presetShadowDirection.isValid() || d->presetShadowOffset.isValid()) {
        writer.writeStartElement(QLatin1String("a:prstShdw"));
        if (d->presetShadowDirection.isValid())
            writer.writeAttribute(QLatin1String("dir"), d->presetShadowDirection.toString());
        if (d->presetShadowOffset.isValid())
            writer.writeAttribute(QLatin1String("dist"), d->presetShadowOffset.toString());
        if (d->presetShadow.has_value()) {
            QString s = QString("shdw%1").arg(d->presetShadow.value());
            writer.writeAttribute(QLatin1String("prst"), s);
        }
        if (d->presetShadowColor.isValid()) d->presetShadowColor.write(writer);
        writer.writeEndElement();
    }
    if (d->reflectionBlurRadius.isValid() || d->reflectionEndOpacity.has_value()
        || d->reflectionStartOpacity.has_value() || d->reflectionStartPosition.has_value()
        || d->reflectionEndPosition.has_value() || d->reflectionGradientDirection.isValid()
        || d->reflectionOffsetDirection.isValid() || d->reflectionShadowOffset.isValid()
        || d->reflectionHorizontalScalingFactor.has_value() || d->reflectionHorizontalSkewFactor.isValid()
        || d->reflectionVerticalScalingFactor.has_value() || d->reflectionVerticalSkewFactor.isValid()
        || d->reflectionRotateWithShape.has_value() || d->reflectionShadowAlignment.has_value()) {
        writer.writeEmptyElement(QLatin1String("a:reflection"));
        if (d->reflectionGradientDirection.isValid())
            writer.writeAttribute(QLatin1String("dir"), d->reflectionGradientDirection.toString());
        writeAttributePercent(writer, QLatin1String("endA"), d->reflectionEndOpacity);
        writeAttributePercent(writer, QLatin1String("stA"), d->reflectionStartOpacity);
        writeAttributePercent(writer, QLatin1String("stPos"), d->reflectionStartPosition);
        writeAttributePercent(writer, QLatin1String("endPos"), d->reflectionEndPosition);
        if (d->reflectionOffsetDirection.isValid())
            writer.writeAttribute(QLatin1String("fadeDir"), d->reflectionOffsetDirection.toString());
        if (d->reflectionBlurRadius.isValid())
            writer.writeAttribute(QLatin1String("blurRad"), d->reflectionBlurRadius.toString());
        if (d->reflectionShadowOffset.isValid())
            writer.writeAttribute(QLatin1String("dist"), d->reflectionShadowOffset.toString());
        if (d->reflectionShadowAlignment.has_value()) {
            QString s; FillFormat::toString(d->reflectionShadowAlignment.value(), s);
            writer.writeAttribute(QLatin1String("algn"), s);
        }
        if (d->reflectionRotateWithShape.has_value())
            writer.writeAttribute(QLatin1String("rotWithShape"), toST_Boolean(d->reflectionRotateWithShape.value()));
        writeAttributePercent(writer, QLatin1String("sx"), d->reflectionHorizontalScalingFactor);
        writeAttributePercent(writer, QLatin1String("sy"), d->reflectionVerticalScalingFactor);
        if (d->reflectionHorizontalSkewFactor.isValid())
            writer.writeAttribute(QLatin1String("kx"), d->reflectionHorizontalSkewFactor.toString());
        if (d->reflectionVerticalSkewFactor.isValid())
            writer.writeAttribute(QLatin1String("ky"), d->reflectionVerticalSkewFactor.toString());

    }
    if (d->softEdgesBlurRadius.isValid()) {
        writer.writeEmptyElement(QLatin1String("a:softEdge"));
        writer.writeAttribute(QLatin1String("rad"), d->softEdgesBlurRadius.toString());
    }

    writer.writeEndElement();
}

EffectPrivate::EffectPrivate()
{

}

EffectPrivate::EffectPrivate(const EffectPrivate &other) : QSharedData(other),
    type(other.type),
    blurRadius(other.blurRadius),
    blurGrow(other.blurGrow),
    fillOverlay(other.fillOverlay),
    fillBlendMode(other.fillBlendMode),
    glowColor(other.glowColor),
    glowRadius(other.glowRadius),
    innerShadowColor(other.innerShadowColor),
    innerShadowBlurRadius(other.innerShadowBlurRadius),
    innerShadowOffset(other.innerShadowOffset),
    innerShadowDirection(other.innerShadowDirection),
    outerShadowColor(other.outerShadowColor),
    outerShadowBlurRadius(other.outerShadowBlurRadius),
    outerShadowOffset(other.outerShadowOffset),
    outerShadowDirection(other.outerShadowDirection),
    outerShadowHorizontalScalingFactor(other.outerShadowHorizontalScalingFactor), //in %
    outerShadowVerticalScalingFactor(other.outerShadowVerticalScalingFactor), //in %
    outerShadowHorizontalSkewFactor(other.outerShadowHorizontalSkewFactor),
    outerShadowVerticalSkewFactor(other.outerShadowVerticalSkewFactor),
    outerShadowRotateWithShape(other.outerShadowRotateWithShape),
    outerShadowAlignment(other.outerShadowAlignment),
    presetShadowColor(other.presetShadowColor),
    presetShadowOffset(other.presetShadowOffset),
    presetShadowDirection(other.presetShadowDirection),
    presetShadow(other.presetShadow),
    reflectionBlurRadius(other.reflectionBlurRadius),
    reflectionStartOpacity(other.reflectionStartOpacity),
    reflectionStartPosition(other.reflectionStartPosition),
    reflectionEndOpacity(other.reflectionEndOpacity),
    reflectionEndPosition(other.reflectionEndPosition),
    reflectionShadowOffset(other.reflectionShadowOffset),
    reflectionGradientDirection(other.reflectionGradientDirection),
    reflectionOffsetDirection(other.reflectionOffsetDirection),
    reflectionHorizontalScalingFactor(other.reflectionHorizontalScalingFactor), //in %
    reflectionVerticalScalingFactor(other.reflectionVerticalScalingFactor), //in %
    reflectionHorizontalSkewFactor(other.reflectionHorizontalSkewFactor),
    reflectionVerticalSkewFactor(other.reflectionVerticalSkewFactor),
    reflectionShadowAlignment(other.reflectionShadowAlignment),
    reflectionRotateWithShape(other.reflectionRotateWithShape),
    softEdgesBlurRadius(other.softEdgesBlurRadius)
{

}

EffectPrivate::~EffectPrivate()
{

}

bool EffectPrivate::operator ==(const EffectPrivate &other) const
{
    if (type != other.type) return false;
    if (blurRadius != other.blurRadius) return false;
    if (blurGrow != other.blurGrow) return false;
    if (fillOverlay != other.fillOverlay) return false;
    if (fillBlendMode != other.fillBlendMode) return false;
    if (glowColor != other.glowColor) return false;
    if (glowRadius != other.glowRadius) return false;
    if (innerShadowColor != other.innerShadowColor) return false;
    if (innerShadowBlurRadius != other.innerShadowBlurRadius) return false;
    if (innerShadowOffset != other.innerShadowOffset) return false;
    if (innerShadowDirection != other.innerShadowDirection) return false;
    if (outerShadowColor != other.outerShadowColor) return false;
    if (outerShadowBlurRadius != other.outerShadowBlurRadius) return false;
    if (outerShadowOffset != other.outerShadowOffset) return false;
    if (outerShadowDirection != other.outerShadowDirection) return false;
    if (outerShadowHorizontalScalingFactor != other.outerShadowHorizontalScalingFactor) return false; //in %
    if (outerShadowVerticalScalingFactor != other.outerShadowVerticalScalingFactor) return false; //in %
    if (outerShadowHorizontalSkewFactor != other.outerShadowHorizontalSkewFactor) return false;
    if (outerShadowVerticalSkewFactor != other.outerShadowVerticalSkewFactor) return false;
    if (outerShadowRotateWithShape != other.outerShadowRotateWithShape) return false;
    if (outerShadowAlignment != other.outerShadowAlignment) return false;
    if (presetShadowColor != other.presetShadowColor) return false;
    if (presetShadowOffset != other.presetShadowOffset) return false;
    if (presetShadowDirection != other.presetShadowDirection) return false;
    if (presetShadow != other.presetShadow) return false;
    if (reflectionBlurRadius != other.reflectionBlurRadius) return false;
    if (reflectionStartOpacity != other.reflectionStartOpacity) return false;
    if (reflectionStartPosition != other.reflectionStartPosition) return false;
    if (reflectionEndOpacity != other.reflectionEndOpacity) return false;
    if (reflectionEndPosition != other.reflectionEndPosition) return false;
    if (reflectionShadowOffset != other.reflectionShadowOffset) return false;
    if (reflectionGradientDirection != other.reflectionGradientDirection) return false;
    if (reflectionOffsetDirection != other.reflectionOffsetDirection) return false;
    if (reflectionHorizontalScalingFactor != other.reflectionHorizontalScalingFactor) return false; //in %
    if (reflectionVerticalScalingFactor != other.reflectionVerticalScalingFactor) return false; //in %
    if (reflectionHorizontalSkewFactor != other.reflectionHorizontalSkewFactor) return false;
    if (reflectionVerticalSkewFactor != other.reflectionVerticalSkewFactor) return false;
    if (reflectionShadowAlignment != other.reflectionShadowAlignment) return false;
    if (reflectionRotateWithShape != other.reflectionRotateWithShape) return false;
    if (softEdgesBlurRadius != other.softEdgesBlurRadius) return false;
    return true;
}

Coordinate Effect::blurRadius() const
{
    if (d)
        return d->blurRadius;
    return Coordinate();
}

void Effect::setBlurRadius(const Coordinate &newBlurRadius)
{
    if (!d) d = new EffectPrivate;
    d->blurRadius = newBlurRadius;
}

std::optional<bool> Effect::blurGrow() const
{
    if (d)
        return d->blurGrow;
    return {};
}

void Effect::setBlurGrow(bool newBlurGrow)
{
    if (!d) d = new EffectPrivate;
    d->blurGrow = newBlurGrow;
}

FillFormat &Effect::fillOverlay()
{
    if (!d) d = new EffectPrivate;
    return d->fillOverlay;
}

FillFormat Effect::fillOverlay() const
{
    if (d)
        return d->fillOverlay;
    return {};
}

void Effect::setFillOverlay(const FillFormat &newFillOverlay)
{
    if (!d) d = new EffectPrivate;
    d->fillOverlay = newFillOverlay;
}

Effect::FillBlendMode Effect::fillBlendMode() const
{
    if (d)
        return d->fillBlendMode;
    return Effect::FillBlendMode::Overlay;
}

void Effect::setFillBlendMode(Effect::FillBlendMode newFillBlendMode)
{
    if (!d) d = new EffectPrivate;
    d->fillBlendMode = newFillBlendMode;
}

Color Effect::glowColor() const
{
    if (d)
        return d->glowColor;
    return Color();
}

void Effect::setGlowColor(const Color &newGlowColor)
{
    if (!d) d = new EffectPrivate;
    d->glowColor = newGlowColor;
}

Coordinate Effect::glowRadius() const
{
    if (d)
        return d->glowRadius;
    return Coordinate();
}

void Effect::setGlowRadius(const Coordinate &newGlowRadius)
{
    if (!d) d = new EffectPrivate;
    d->glowRadius = newGlowRadius;
}

Color Effect::innerShadowColor() const
{
    if (d)
        return d->innerShadowColor;
    return Color();
}

void Effect::setInnerShadowColor(const Color &newInnerShadowColor)
{
    if (!d) d = new EffectPrivate;
    d->innerShadowColor = newInnerShadowColor;
}

Coordinate Effect::innerShadowBlurRadius() const
{
    if (d)
        return d->innerShadowBlurRadius;
    return Coordinate();
}

void Effect::setInnerShadowBlurRadius(const Coordinate &newInnerShadowBlurRadius)
{
    if (!d) d = new EffectPrivate;
    d->innerShadowBlurRadius = newInnerShadowBlurRadius;
}

Coordinate Effect::innerShadowOffset() const
{
    if (d)
        return d->innerShadowOffset;
    return Coordinate();
}

void Effect::setInnerShadowOffset(const Coordinate &newInnerShadowOffset)
{
    if (!d) d = new EffectPrivate;
    d->innerShadowOffset = newInnerShadowOffset;
}

Angle Effect::innerShadowDirection() const
{
    if (d)
        return d->innerShadowDirection;
    return {};
}

void Effect::setInnerShadowDirection(Angle newInnerShadowDirection)
{
    if (!d) d = new EffectPrivate;
    d->innerShadowDirection = newInnerShadowDirection;
}

Color Effect::outerShadowColor() const
{
    if (d)
        return d->outerShadowColor;
    return Color();
}

void Effect::setOuterShadowColor(const Color &newOuterShadowColor)
{
    if (!d) d = new EffectPrivate;
    d->outerShadowColor = newOuterShadowColor;
}

Coordinate Effect::outerShadowBlurRadius() const
{
    if (d)
        return d->outerShadowBlurRadius;
    return Coordinate();
}

void Effect::setOuterShadowBlurRadius(const Coordinate &newOuterShadowBlurRadius)
{
    if (!d) d = new EffectPrivate;
    d->outerShadowBlurRadius = newOuterShadowBlurRadius;
}

Coordinate Effect::outerShadowOffset() const
{
    if (d)
        return d->outerShadowOffset;
    return Coordinate();
}

void Effect::setOuterShadowOffset(const Coordinate &newOuterShadowOffset)
{
    if (!d) d = new EffectPrivate;
    d->outerShadowOffset = newOuterShadowOffset;
}

Angle Effect::outerShadowDirection() const
{
    if (d)
        return d->outerShadowDirection;
    return {};
}

void Effect::setOuterShadowDirection(Angle newOuterShadowDirection)
{
    if (!d) d = new EffectPrivate;
    d->outerShadowDirection = newOuterShadowDirection;
}

std::optional<double> Effect::outerShadowHorizontalScalingFactor() const
{
    if (d)
        return d->outerShadowHorizontalScalingFactor;
    return {};
}

void Effect::setOuterShadowHorizontalScalingFactor(double newOuterShadowHorizontalScalingFactor)
{
    if (!d) d = new EffectPrivate;
    d->outerShadowHorizontalScalingFactor = newOuterShadowHorizontalScalingFactor;
}

std::optional<double> Effect::outerShadowVerticalScalingFactor() const
{
    if (d)
        return d->outerShadowVerticalScalingFactor;
    return {};
}

void Effect::setOuterShadowVerticalScalingFactor(double newOuterShadowVerticalScalingFactor)
{
    if (!d) d = new EffectPrivate;
    d->outerShadowVerticalScalingFactor = newOuterShadowVerticalScalingFactor;
}

Angle Effect::outerShadowHorizontalSkewFactor() const
{
    if (d)
        return d->outerShadowHorizontalSkewFactor;
    return {};
}

void Effect::setOuterShadowHorizontalSkewFactor(Angle newOuterShadowHorizontalSkewFactor)
{
    if (!d) d = new EffectPrivate;
    d->outerShadowHorizontalSkewFactor = newOuterShadowHorizontalSkewFactor;
}

Angle Effect::outerShadowVerticalSkewFactor() const
{
    if (d)
        return d->outerShadowVerticalSkewFactor;
    return {};
}

void Effect::setOuterShadowVerticalSkewFactor(Angle newOuterShadowVerticalSkewFactor)
{
    if (!d) d = new EffectPrivate;
    d->outerShadowVerticalSkewFactor = newOuterShadowVerticalSkewFactor;
}

std::optional<bool> Effect::outerShadowRotateWithShape() const
{
    if (d)
        return d->outerShadowRotateWithShape;
    return {};
}

void Effect::setOuterShadowRotateWithShape(bool newOuterShadowRotateWithShape)
{
    if (!d) d = new EffectPrivate;
    d->outerShadowRotateWithShape = newOuterShadowRotateWithShape;
}

std::optional<FillFormat::Alignment> Effect::outerShadowAlignment() const
{
    if (d)
        return d->outerShadowAlignment;
    return {};
}

void Effect::setOuterShadowAlignment(FillFormat::Alignment newOuterShadowAlignment)
{
    if (!d) d = new EffectPrivate;
    d->outerShadowAlignment = newOuterShadowAlignment;
}

Color Effect::presetShadowColor() const
{
    if (d)
        return d->presetShadowColor;
    return Color();
}

void Effect::setPresetShadowColor(const Color &newPresetShadowColor)
{
    if (!d) d = new EffectPrivate;
    d->presetShadowColor = newPresetShadowColor;
}

Coordinate Effect::presetShadowOffset() const
{
    if (d)
        return d->presetShadowOffset;
    return Coordinate();
}

void Effect::setPresetShadowOffset(const Coordinate &newPresetShadowOffset)
{
    if (!d) d = new EffectPrivate;
    d->presetShadowOffset = newPresetShadowOffset;
}

Angle Effect::presetShadowDirection() const
{
    if (d)
        return d->presetShadowDirection;
    return {};
}

void Effect::setPresetShadowDirection(Angle newPresetShadowDirection)
{
    if (!d) d = new EffectPrivate;
    d->presetShadowDirection = newPresetShadowDirection;
}

std::optional<int> Effect::presetShadow() const
{
    if (d)
        return d->presetShadow;
    return {};
}

void Effect::setPresetShadow(int newPresetShadow)
{
    if (newPresetShadow < 1 || newPresetShadow > 20) return;
    if (!d) d = new EffectPrivate;
    d->presetShadow = newPresetShadow;
}

Coordinate Effect::reflectionBlurRadius() const
{
    if (d)
        return d->reflectionBlurRadius;
    return Coordinate();
}

void Effect::setReflectionBlurRadius(const Coordinate &newReflectionBlurRadius)
{
    if (!d) d = new EffectPrivate;
    d->reflectionBlurRadius = newReflectionBlurRadius;
}

std::optional<double> Effect::reflectionStartOpacity() const
{
    if (d)
        return d->reflectionStartOpacity;
    return {};
}

void Effect::setReflectionStartOpacity(double newReflectionStartOpacity)
{
    if (!d) d = new EffectPrivate;
    d->reflectionStartOpacity = newReflectionStartOpacity;
}

std::optional<double> Effect::reflectionStartPosition() const
{
    if (d)
        return d->reflectionStartPosition;
    return {};
}

void Effect::setReflectionStartPosition(double newReflectionStartPosition)
{
    if (!d) d = new EffectPrivate;
    d->reflectionStartPosition = newReflectionStartPosition;
}

std::optional<double> Effect::reflectionEndOpacity() const
{
    if (d)
        return d->reflectionEndOpacity;
    return {};
}

void Effect::setReflectionEndOpacity(double newReflectionEndOpacity)
{
    if (!d) d = new EffectPrivate;
    d->reflectionEndOpacity = newReflectionEndOpacity;
}

std::optional<double> Effect::reflectionEndPosition() const
{
    if (d)
        return d->reflectionEndPosition;
    return {};
}

void Effect::setReflectionEndPosition(double newReflectionEndPosition)
{
    if (!d) d = new EffectPrivate;
    d->reflectionEndPosition = newReflectionEndPosition;
}

Coordinate Effect::reflectionShadowOffset() const
{
    if (d)
        return d->reflectionShadowOffset;
    return Coordinate();
}

void Effect::setReflectionShadowOffset(const Coordinate &newReflectionShadowOffset)
{
    if (!d) d = new EffectPrivate;
    d->reflectionShadowOffset = newReflectionShadowOffset;
}

Angle Effect::reflectionGradientDirection() const
{
    if (d)
        return d->reflectionGradientDirection;
    return {};
}

void Effect::setReflectionGradientDirection(Angle newReflectionGradientDirection)
{
    if (!d) d = new EffectPrivate;
    d->reflectionGradientDirection = newReflectionGradientDirection;
}

Angle Effect::reflectionOffsetDirection() const
{
    if (d)
        return d->reflectionOffsetDirection;
    return {};
}

void Effect::setReflectionOffsetDirection(Angle newReflectionOffsetDirection)
{
    if (!d) d = new EffectPrivate;
    d->reflectionOffsetDirection = newReflectionOffsetDirection;
}

std::optional<double> Effect::reflectionHorizontalScalingFactor() const
{
    if (d)
        return d->reflectionHorizontalScalingFactor;
    return {};
}

void Effect::setReflectionHorizontalScalingFactor(double newReflectionHorizontalScalingFactor)
{
    if (!d) d = new EffectPrivate;
    d->reflectionHorizontalScalingFactor = newReflectionHorizontalScalingFactor;
}

std::optional<double> Effect::reflectionVerticalScalingFactor() const
{
    if (d)
        return d->reflectionVerticalScalingFactor;
    return {};
}

void Effect::setReflectionVerticalScalingFactor(double newReflectionVerticalScalingFactor)
{
    if (!d) d = new EffectPrivate;
    d->reflectionVerticalScalingFactor = newReflectionVerticalScalingFactor;
}

Angle Effect::reflectionHorizontalSkewFactor() const
{
    if (d)
        return d->reflectionHorizontalSkewFactor;
    return {};
}

void Effect::setReflectionHorizontalSkewFactor(Angle newReflectionHorizontalSkewFactor)
{
    if (!d) d = new EffectPrivate;
    d->reflectionHorizontalSkewFactor = newReflectionHorizontalSkewFactor;
}

Angle Effect::reflectionVerticalSkewFactor() const
{
    if (d)
        return d->reflectionVerticalSkewFactor;
    return {};
}

void Effect::setReflectionVerticalSkewFactor(Angle newReflectionVerticalSkewFactor)
{
    if (!d) d = new EffectPrivate;
    d->reflectionVerticalSkewFactor = newReflectionVerticalSkewFactor;
}

std::optional<FillFormat::Alignment> Effect::reflectionShadowAlignment() const
{
    if (d)
        return d->reflectionShadowAlignment;
    return {};
}

void Effect::setReflectionShadowAlignment(FillFormat::Alignment newReflectionShadowAlignment)
{
    if (!d) d = new EffectPrivate;
    d->reflectionShadowAlignment = newReflectionShadowAlignment;
}

std::optional<bool> Effect::reflectionRotateWithShape() const
{
    if (d)
        return d->reflectionRotateWithShape;
    return {};
}

void Effect::setReflectionRotateWithShape(bool newReflectionRotateWithShape)
{
    if (!d) d = new EffectPrivate;
    d->reflectionRotateWithShape = newReflectionRotateWithShape;
}

Coordinate Effect:: softEdgesBlurRadius() const
{
    if (d)
        return d->softEdgesBlurRadius;
    return Coordinate();
}

void Effect::setSoftEdgesBlurRadius(const Coordinate &newSoftEdgesBlurRadius)
{
    if (!d) d = new EffectPrivate;
    d->softEdgesBlurRadius = newSoftEdgesBlurRadius;
}

QDebug operator<<(QDebug dbg, const Effect &e)
{
    QDebugStateSaver saver(dbg);
    dbg.setAutoInsertSpaces(false);

    dbg << "QXlsx::Effect(";
    if (e.isValid()) {
        dbg << "type: " << static_cast<int>(e.d->type) << ", ";
        if (e.d->blurRadius.isValid()) dbg << "blurRadius: " << e.d->blurRadius.toString() << ", ";
        if (e.d->blurGrow.has_value()) dbg << "blurGrow: " << e.d->blurGrow.value() << ", ";
        if (e.d->fillOverlay.isValid()) dbg << "fillOverlay: " << e.d->fillOverlay << ", ";
        dbg << "type: " << static_cast<int>(e.d->fillBlendMode) << ", ";
        if (e.d->glowColor.isValid()) dbg << "glowColor: " << e.d->glowColor << ", ";
        if (e.d->glowRadius.isValid()) dbg << "glowRadius: " << e.d->glowRadius.toString() << ", ";
        if (e.d->innerShadowColor.isValid()) dbg << "innerShadowColor: " << e.d->innerShadowColor << ", ";
        if (e.d->innerShadowBlurRadius.isValid()) dbg << "innerShadowBlurRadius: " << e.d->innerShadowBlurRadius.toString() << ", ";
        if (e.d->innerShadowOffset.isValid()) dbg << "innerShadowOffset: " << e.d->innerShadowOffset.toString() << ", ";
        if (e.d->innerShadowDirection.isValid()) dbg << "innerShadowDirection: " << e.d->innerShadowDirection.toString() << ", ";
        if (e.d->outerShadowColor.isValid()) dbg << "outerShadowColor: " << e.d->outerShadowColor << ", ";
        if (e.d->outerShadowBlurRadius.isValid()) dbg << "outerShadowBlurRadius: " << e.d->outerShadowBlurRadius.toString() << ", ";
        if (e.d->outerShadowOffset.isValid()) dbg << "outerShadowOffset: " << e.d->outerShadowOffset.toString() << ", ";
        if (e.d->outerShadowDirection.isValid()) dbg << "outerShadowDirection: " << e.d->outerShadowDirection.toString() << ", ";
        if (e.d->outerShadowHorizontalScalingFactor.has_value()) dbg << "outerShadowHorizontalScalingFactor: " << e.d->outerShadowHorizontalScalingFactor.value() << ", "; //in %
        if (e.d->outerShadowVerticalScalingFactor.has_value()) dbg << "outerShadowVerticalScalingFactor: " << e.d->outerShadowVerticalScalingFactor.value() << ", "; //in %
        if (e.d->outerShadowHorizontalSkewFactor.isValid()) dbg << "outerShadowHorizontalSkewFactor: " << e.d->outerShadowHorizontalSkewFactor.toString() << ", ";
        if (e.d->outerShadowVerticalSkewFactor.isValid()) dbg << "outerShadowVerticalSkewFactor: " << e.d->outerShadowVerticalSkewFactor.toString() << ", ";
        if (e.d->outerShadowRotateWithShape.has_value()) dbg << "outerShadowRotateWithShape: " << e.d->outerShadowRotateWithShape.value() << ", ";
        if (e.d->outerShadowAlignment.has_value()) dbg << "type: " << static_cast<int>(e.d->outerShadowAlignment.value()) << ", ";
        if (e.d->presetShadowColor.isValid()) dbg << "presetShadowColor: " << e.d->presetShadowColor << ", ";
        if (e.d->presetShadowOffset.isValid()) dbg << "presetShadowOffset: " << e.d->presetShadowOffset.toString() << ", ";
        if (e.d->presetShadowDirection.isValid()) dbg << "presetShadowDirection: " << e.d->presetShadowDirection.toString() << ", ";
        if (e.d->presetShadow.has_value()) dbg << "presetShadow: " << e.d->presetShadow.value() << ", ";
        if (e.d->reflectionBlurRadius.isValid()) dbg << "reflectionBlurRadius: " << e.d->reflectionBlurRadius.toString() << ", ";
        if (e.d->reflectionStartOpacity.has_value()) dbg << "reflectionStartOpacity: " << e.d->reflectionStartOpacity.value() << ", ";
        if (e.d->reflectionStartPosition.has_value()) dbg << "reflectionStartPosition: " << e.d->reflectionStartPosition.value() << ", ";
        if (e.d->reflectionEndOpacity.has_value()) dbg << "reflectionEndOpacity: " << e.d->reflectionEndOpacity.value() << ", ";
        if (e.d->reflectionEndPosition.has_value()) dbg << "reflectionEndPosition: " << e.d->reflectionEndPosition.value() << ", ";
        if (e.d->reflectionShadowOffset.isValid()) dbg << "reflectionShadowOffset: " << e.d->reflectionShadowOffset.toString() << ", ";
        if (e.d->reflectionGradientDirection.isValid()) dbg << "reflectionGradientDirection: " << e.d->reflectionGradientDirection.toString() << ", ";
        if (e.d->reflectionOffsetDirection.isValid()) dbg << "reflectionOffsetDirection: " << e.d->reflectionOffsetDirection.toString() << ", ";
        if (e.d->reflectionHorizontalScalingFactor.has_value()) dbg << "reflectionHorizontalScalingFactor: " << e.d->reflectionHorizontalScalingFactor.value() << ", "; //in %
        if (e.d->reflectionVerticalScalingFactor.has_value()) dbg << "reflectionVerticalScalingFactor: " << e.d->reflectionVerticalScalingFactor.value() << ", "; //in %
        if (e.d->reflectionHorizontalSkewFactor.isValid()) dbg << "reflectionHorizontalSkewFactor: " << e.d->reflectionHorizontalSkewFactor.toString() << ", ";
        if (e.d->reflectionVerticalSkewFactor.isValid()) dbg << "reflectionVerticalSkewFactor: " << e.d->reflectionVerticalSkewFactor.toString() << ", ";
        if (e.d->reflectionShadowAlignment.has_value()) dbg << "type: " << static_cast<int>(e.d->reflectionShadowAlignment.value()) << ", ";
        if (e.d->reflectionRotateWithShape.has_value()) dbg << "reflectionRotateWithShape: " << e.d->reflectionRotateWithShape.value() << ", ";
        if (e.d->softEdgesBlurRadius.isValid()) dbg << "softEdgesBlurRadius: " << e.d->softEdgesBlurRadius.toString() << ", ";
    }
    dbg << ")";

    return dbg;
}

}


