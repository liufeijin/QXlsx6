// xlsxcolor.cpp

#include <QDataStream>
#include <QXmlStreamReader>
#include <QXmlStreamWriter>
#include <QDebug>
#include <QtMath>

#include "xlsxcolor.h"
#include "xlsxutility_p.h"
#include "xlsxmain.h"

namespace QXlsx {

Color::Color() : type_(Type::Invalid)
{

}

Color::Color(QColor color) : type_{Type::RGB}
{
    setRgb(color);
}

Color::Color(Color::SchemeColor color) : type_{Type::Scheme}
{
    val = static_cast<int>(color);
}

Color::Color(Color::SystemColor color) : type_{Type::System}
{
    val = static_cast<int>(color);
}

Color::Color(const QString &colorName) : type_{Type::Preset}
{
    val = colorName;
}

Color::Color(const char *colorName): type_{Type::Preset}
{
    val = QString::fromLocal8Bit(colorName);
}

Color::Color(const Color &other) : type_{other.type_}, val{other.val}, tr{other.tr},
    lastColor{other.lastColor}, isCRGB{other.isCRGB}, isDirty{other.isDirty}, m_key{other.m_key}
{

}

Color &Color::operator=(const Color &other)
{
    if (*this != other) {
        type_= other.type_;
        val = other.val;
        tr = other.tr;
        lastColor = other.lastColor;
        isCRGB = other.isCRGB;
        isDirty = other.isDirty;
        m_key = other.m_key;
    }
    return *this;
}

Color::~Color()
{

}

bool Color::isValid() const
{
    if (type_ == Type::Invalid) return false;
    return val.isValid();
}

void Color::setRgb(const QColor &color)
{
    type_ = Type::RGB;
    val = color;
    isDirty = true;
}

void Color::setHsl(const QColor &color)
{
    type_ = Type::HSL;
    val = color;
    isDirty = true;
}

void Color::setIndexedColor(int index)
{
    type_ = Type::Indexed;
    val = index;
    isDirty = true;
}

void Color::setPresetColor(const QString &colorName)
{
    if (colorName.isEmpty()) return;

    type_ = Type::Preset;
    val = colorName;
    isDirty = true;
}

void Color::setSchemeColor(Color::SchemeColor color)
{
    type_ = Type::Scheme;
    val = static_cast<int>(color);
    isDirty = true;
}

void Color::setSystemColor(Color::SystemColor color)
{
    type_ = Type::System;
    val = static_cast<int>(color);
    isDirty = true;
}

void Color::setAutoColor()
{
    type_ = Type::Auto;
    val = true;
    isDirty = true;
}

void Color::addTransform(ColorTransform::Type transform, QVariant val)
{
    tr.vals.insert(transform, val);
    isDirty = true;
}

ColorTransform Color::transforms() const
{
    return tr;
}

QColor Color::transformed() const
{
    return tr.transformed(toQColor());
}

bool Color::hasTransform(ColorTransform::Type type) const
{
    return tr.hasTransform(type);
}
QVariant Color::transform(ColorTransform::Type type) const
{
    return tr.transform(type);
}

Color::Type Color::type() const
{
    return type_;
}

bool Color::isAutoColor() const
{
    return type_ == Type::Auto;
}

QColor Color::rgb() const
{
    return (type_ == Type::RGB)
            ? val.value<QColor>() : QColor();
}

QColor Color::hsl() const
{
    return type_ == Type::HSL ? val.value<QColor>() : QColor();
}

QColor Color::toQColor() const
{
    if (type_ == Type::HSL || type_ == Type::RGB) return val.value<QColor>();
    if (type_ == Type::Preset) return QColor(val.toString());
    return {};
}

QColor Color::presetColor() const
{
    return (type_ == Type::Preset) ? QColor(val.toString()) : QColor();
}

QString Color::presetColorName() const
{
    return (type_ == Type::Preset) ? val.toString() : QString();
}

Color::SchemeColor Color::schemeColor() const
{
    if (type_ != Type::Scheme) return SchemeColor::Style;
    SchemeColor c = static_cast<SchemeColor>(val.toInt());
    return c;
}

QString Color::schemeColorName() const
{
    if (type_ != Type::Scheme) return QString();

    SchemeColor c = static_cast<SchemeColor>(val.toInt());
    QString s;
    toString(c, s);

    return s;
}

Color::SystemColor Color::systemColor() const
{
    if (type_ != Type::System) return SystemColor::None;
    SystemColor c = static_cast<SystemColor>(val.toInt());
    return c;
}

QString Color::systemColorName() const
{
    if (type_ != Type::System) return QString();

    SystemColor c = static_cast<SystemColor>(val.toInt());
    QString s;
    toString(c, s);

    return s;
}

int Color::indexedColor() const
{
    return (type_ == Type::Indexed) ? val.toInt() : -1;
}

//QPair<int, double> Color::themeColor() const
//{
//    if (type_ != Type::Indexed || val.userType() != QMetaType::QVariantMap) return {-1, 0.0};

//    auto m = val.toMap();
//    return {m.value("theme").toInt(), m.value("tint").toDouble()};
//}

bool Color::write(QXmlStreamWriter &writer, const QString &node) const
{
    if (!node.isEmpty()) {
        //colors for the cells
        writer.writeEmptyElement(node);
        switch (type_) {
            case Type::Auto: writer.writeAttribute(QStringLiteral("auto"), val.toBool() ? "true" : "false"); break;
            case Type::HSL:
            case Type::RGB: {
                //write hsl as rgb, as there is no hsl in colors of cells
                if (val.userType() == qMetaTypeId<QColor>()) {
                    QColor c = val.value<QColor>();
                    writer.writeAttribute(QStringLiteral("rgb"), Color::toARGBString(tr.transformed(c)));
                }
                break;
            }
            case Type::Preset: {
                QColor c; c.setNamedColor(val.toString());
                writer.writeAttribute(QStringLiteral("rgb"), Color::toARGBString(tr.transformed(c)));
                break;
            }
            case Type::Scheme: {
                //not all scheme color can be written
                switch (schemeColor()) {
                    case SchemeColor::Background1:
                    case SchemeColor::Text1:
                    case SchemeColor::Background2:
                    case SchemeColor::Text2:
                    case SchemeColor::Style: break;
                    case SchemeColor::Accent1: writer.writeAttribute(QStringLiteral("theme"), QStringLiteral("4")); break;
                    case SchemeColor::Accent2: writer.writeAttribute(QStringLiteral("theme"), QStringLiteral("5")); break;
                    case SchemeColor::Accent3: writer.writeAttribute(QStringLiteral("theme"), QStringLiteral("6")); break;
                    case SchemeColor::Accent4: writer.writeAttribute(QStringLiteral("theme"), QStringLiteral("7")); break;
                    case SchemeColor::Accent5: writer.writeAttribute(QStringLiteral("theme"), QStringLiteral("8")); break;
                    case SchemeColor::Accent6: writer.writeAttribute(QStringLiteral("theme"), QStringLiteral("9")); break;
                    case SchemeColor::Hlink: writer.writeAttribute(QStringLiteral("theme"), QStringLiteral("10")); break;
                    case SchemeColor::FollowedHlink: writer.writeAttribute(QStringLiteral("theme"), QStringLiteral("11")); break;
                    case SchemeColor::Dark1: writer.writeAttribute(QStringLiteral("theme"), QStringLiteral("0")); break;
                    case SchemeColor::Light1: writer.writeAttribute(QStringLiteral("theme"), QStringLiteral("1")); break;
                    case SchemeColor::Dark2: writer.writeAttribute(QStringLiteral("theme"), QStringLiteral("2")); break;
                    case SchemeColor::Light2: writer.writeAttribute(QStringLiteral("theme"), QStringLiteral("3")); break;
                }
                if (tr.hasTransform(ColorTransform::Type::Tint)) {
                    auto tint = tr.transform(ColorTransform::Type::Tint).toDouble() / 100;
                    writer.writeAttribute(QStringLiteral("tint"), QString::number(tint));
                }
                else if (tr.hasTransform(ColorTransform::Type::Shade)) {
                    auto shade = -tr.transform(ColorTransform::Type::Shade).toDouble() / 100;
                    writer.writeAttribute(QStringLiteral("tint"), QString::number(shade));
                }
                break;
            }
            case Type::System: break;
            case Type::Indexed: {
                writer.writeAttribute(QStringLiteral("indexed"), val.toString());
                if (tr.hasTransform(ColorTransform::Type::Tint)) {
                    auto tint = tr.transform(ColorTransform::Type::Tint).toDouble() / 100;
                    writer.writeAttribute(QStringLiteral("tint"), QString::number(tint));
                }
                else if (tr.hasTransform(ColorTransform::Type::Shade)) {
                    auto shade = -tr.transform(ColorTransform::Type::Shade).toDouble() / 100;
                    writer.writeAttribute(QStringLiteral("tint"), QString::number(shade));
                }
                break;
            }
            case Type::Invalid: break;
        }
        return true;
    }

    switch (type_) {
        case Type::Auto:
        case Type::Indexed:
        case Type::Invalid: return false;
        case Type::RGB: {
            if (isCRGB) {
                writer.writeStartElement(QLatin1String("a:scrgbClr"));
                QColor col = val.value<QColor>();
                writer.writeAttribute(QLatin1String("r"), QString::number(col.redF()*100)+"%");
                writer.writeAttribute(QLatin1String("g"), QString::number(col.greenF()*100)+"%");
                writer.writeAttribute(QLatin1String("b"), QString::number(col.blueF()*100)+"%");
            }
            else {
                writer.writeStartElement(QLatin1String("a:srgbClr"));
                writer.writeAttribute(QLatin1String("val"), toRGBString(val.value<QColor>()));
            }
            tr.write(writer);
            writer.writeEndElement();
            break;
        }
        case Type::HSL: {
            writer.writeStartElement(QLatin1String("a:hslClr"));
            QColor col = val.value<QColor>();
            writer.writeAttribute(QLatin1String("hue"), QString::number(qint64(col.hueF()*21600000)));
            writeAttributePercent(writer, QLatin1String("sat"), col.saturationF()*100);
            writeAttributePercent(writer, QLatin1String("lum"), col.lightnessF()*100);

            tr.write(writer);
            writer.writeEndElement();
            break;
        }
        case Type::Preset: {
            writer.writeStartElement(QLatin1String("a:prstClr"));
            writer.writeAttribute(QLatin1String("val"), val.toString());

            tr.write(writer);
            writer.writeEndElement();
            break;
        }
        case Type::Scheme: {
            writer.writeStartElement(QLatin1String("a:schemeClr"));
            writer.writeAttribute(QLatin1String("val"), schemeColorName());

            tr.write(writer);
            writer.writeEndElement();
            break;
        }
        case Type::System: {
            writer.writeStartElement(QLatin1String("a:sysClr"));
            writer.writeAttribute(QLatin1String("val"), systemColorName());

            if (lastColor.isValid())
                writer.writeAttribute(QLatin1String("lastClr"), toRGBString(lastColor));

            tr.write(writer);
            writer.writeEndElement();
            break;
        }
    }

    return true;
}

bool Color::readSimple(QXmlStreamReader &reader)
{
    const auto& a = reader.attributes();
    if (a.hasAttribute(QLatin1String("rgb"))) {
        type_ = Type::RGB;
        const auto& colorString = a.value(QLatin1String("rgb")).toString();
        val.setValue(fromARGBString(colorString));
    } else if (a.hasAttribute(QLatin1String("indexed"))) {
        type_ = Type::Indexed;
        int index = a.value(QLatin1String("indexed")).toInt();
        val.setValue(index);
    } else if (a.hasAttribute(QLatin1String("theme"))) {
        type_ = Type::Scheme;
        const auto& theme = a.value(QLatin1String("theme")).toString();

        switch (theme.toInt()) {
            case 0: val = static_cast<int>(SchemeColor::Dark1); break;
            case 1: val = static_cast<int>(SchemeColor::Light1); break;
            case 2: val = static_cast<int>(SchemeColor::Dark2); break;
            case 3: val = static_cast<int>(SchemeColor::Light2); break;
            case 4: val = static_cast<int>(SchemeColor::Accent1); break;
            case 5: val = static_cast<int>(SchemeColor::Accent2); break;
            case 6: val = static_cast<int>(SchemeColor::Accent3); break;
            case 7: val = static_cast<int>(SchemeColor::Accent4); break;
            case 8: val = static_cast<int>(SchemeColor::Accent5); break;
            case 9: val = static_cast<int>(SchemeColor::Accent6); break;
            case 10: val = static_cast<int>(SchemeColor::Hlink); break;
            case 11: val = static_cast<int>(SchemeColor::FollowedHlink); break;
            default: return false;
        }
    } else if (a.hasAttribute(QLatin1String("auto"))) {
        val.setValue(fromST_Boolean(a.value(QLatin1String("auto"))));
    }
    if (a.hasAttribute(QLatin1String("tint"))) {
        auto tint = a.value(QLatin1String("tint")).toDouble();
        if (tint > 0) addTransform(ColorTransform::Type::Tint, tint*100);
        if (tint < 0) addTransform(ColorTransform::Type::Shade, -tint*100);
    }
    return true;
}

bool Color::readComplex(QXmlStreamReader &reader)
{
    const auto &name = reader.name();
    if (name == QLatin1String("scrgbClr")) {
        type_ = Type::RGB;
        isCRGB = true;
    }
    else if (name == QLatin1String("srgbClr")) type_ = Type::RGB;
    else if (name == QLatin1String("sysClr")) type_ = Type::System;
    else if (name == QLatin1String("hslClr")) type_ = Type::HSL;
    else if (name == QLatin1String("prstClr")) type_ = Type::Preset;
    else if (name == QLatin1String("schemeClr")) type_ = Type::Scheme;
    else type_ = Type::Invalid;

    const auto& attributes = reader.attributes();
    switch (type_) {
        case Type::RGB: {
            if (isCRGB) {
                auto r = fromST_Percent(attributes.value(QLatin1String("r")));
                auto g = fromST_Percent(attributes.value(QLatin1String("g")));
                auto b = fromST_Percent(attributes.value(QLatin1String("b")));
                QColor color = QColor::fromRgbF(r / 100.0,
                                                g / 100.0,
                                                b / 100.0);
                val = color;
            }
            else {
                const auto& colorString = attributes.value(QLatin1String("val")).toString();
                val.setValue(fromARGBString(colorString));
            }
            tr.read(reader);
            break;
        }
        case Type::HSL: {
            Angle h;
            parseAttribute(attributes, QLatin1String("hue"), h);
            auto s = fromST_Percent(attributes.value(QLatin1String("sat")));
            auto l = fromST_Percent(attributes.value(QLatin1String("lum")));
            QColor color = QColor::fromHslF(h.toDegrees() / 360.0,
                                            s / 100.0,
                                            l / 100.0);
            setRgb(color);
            tr.read(reader);
            break;
        }
        case Type::Preset: {
            const auto& colorString = attributes.value(QLatin1String("val")).toString();
            val.setValue(colorString);
            tr.read(reader);
            break;
        }
        case Type::Scheme: {
            const auto& colorString = attributes.value(QLatin1String("val")).toString();
            Color::SchemeColor col;
            fromString(colorString, col);
            setSchemeColor(col);
            tr.read(reader);
            break;
        }
        case Type::System: {
            const auto& colorString = attributes.value(QLatin1String("val")).toString();
            Color::SystemColor col;
            fromString(colorString, col);
            setSystemColor(col);
            tr.read(reader);

            if (attributes.hasAttribute(QLatin1String("lastClr")))
                lastColor = fromARGBString(attributes.value(QLatin1String("lastClr")).toString());
            break;
        }
        default: return false;
    }
    return true;
}

bool Color::read(QXmlStreamReader &reader)
{
    isDirty = true;
    const auto &name = reader.name();
    if (name == QLatin1String("scrgbClr") || name == QLatin1String("srgbClr") ||
        name == QLatin1String("sysClr") || name == QLatin1String("hslClr") ||
        name == QLatin1String("prstClr") || name == QLatin1String("schemeClr"))
        return readComplex(reader);
    return readSimple(reader);
}

bool Color::operator ==(const Color &other) const
{
    if (type_ != other.type_) return false;
    if (val != other.val) return false;
    if (tr.vals != other.tr.vals) return false;
    if (lastColor != other.lastColor) return false;
    return true;
}

bool Color::operator !=(const Color &other) const
{
    if (type_ != other.type_) return true;
    if (val != other.val) return true;
    if (tr.vals != other.tr.vals) return true;
    if (lastColor != other.lastColor) return true;
    return false;
}

Color::operator QVariant() const
{
    const auto& cref
#if QT_VERSION >= 0x060000 // Qt 6.0 or over
        = QMetaType::fromType<Color>();
#else
        = qMetaTypeId<Color>() ;
#endif
    QVariant v(cref, this);
    return v;
}


QColor Color::fromARGBString(const QString &c)
{
    QColor color;

    if (c.startsWith(u'#')) {
        color.setNamedColor(c);
    } else {
        color.setNamedColor(QLatin1Char('#') + c);
    }
    return color;
}

QString Color::toARGBString(const QColor &c)
{
    return QString::asprintf("%02X%02X%02X%02X", c.alpha(), c.red(), c.green(), c.blue());
}

QString Color::toRGBString(const QColor &c)
{
    return QString::asprintf("%02X%02X%02X", c.red(), c.green(), c.blue());
}


QDebug operator<<(QDebug dbg, const Color &c)
{
    if (!c.isValid())
        dbg.nospace() << "Color(invalid)";
    else switch (c.type_) {
        case Color::Type::RGB: {
            if (c.isCRGB)
                dbg.nospace() << "Color(crgb, " << c.rgb();
            else
                dbg.nospace() << "Color(rgb, " << c.rgb();
            break;
        }
        case Color::Type::HSL:
            dbg.nospace() << "Color(hsl, " << c.hsl().toHsl();  break;
        case Color::Type::Preset:
            dbg.nospace() << "Color(preset, " << c.presetColor();  break;
        case Color::Type::Scheme:
            dbg.nospace() << "Color(scheme, " << c.schemeColor();  break;
        case Color::Type::System:
            dbg.nospace() << "Color(system, " << c.systemColor();
            if (c.lastColor.isValid()) dbg.nospace() << ", lastColor " << c.lastColor;
            break;
        case Color::Type::Indexed:
            dbg.nospace() << "Color(indexed, " << c.indexedColor();  break;
        case Color::Type::Auto:
            dbg.nospace() << "Color(auto"; break;
        default: break;
    }
    if (!c.tr.vals.isEmpty()) {
        dbg.nospace() << ", transforms: "<< c.tr;
    }
    dbg.nospace() << ")";

    return dbg.space();
}

bool ColorTransform::hasTransform(Type type) const
{
    return vals.contains(type);
}
QVariant ColorTransform::transform(Type type) const
{
    return vals.value(type);
}

QColor &ColorTransform::applyTransform(Type type, const QVariant &val, QColor &color)
{
    switch (type) {
        case Type::Tint: {
            auto h = color.hslHueF();
            auto s = color.hslSaturationF();
            auto l = color.lightnessF();
            auto a = color.alphaF();
            auto l1 = val.toDouble() / 100.0;
            color = QColor::fromHslF(h, s, l*l1+(1-l1), a);
            break;
        }
        case Type::Shade: {
            auto h = color.hslHueF();
            auto s = color.hslSaturationF();
            auto l = color.lightnessF();
            auto a = color.alphaF();
            auto l1 = val.toDouble() / 100.0;
            color = QColor::fromHslF(h, s, l*l1, a);
            break;
        }
        case Type::Complement: {
            if (val.toBool()) {
                auto h = color.hslHueF();
                auto s = color.hslSaturationF();
                auto l = color.lightnessF();
                auto a = color.alphaF();
                h+=0.5;
                if (h>1) h-=1;
                color = QColor::fromHslF(h, s, l, a);
            }
            break;
        }
        case Type::Inverse: {
            if (val.toBool()) {
                auto h = color.hslHueF();
                auto s = color.hslSaturationF();
                auto l = color.lightnessF();
                auto a = color.alphaF();
                h+=0.5; if (h>1) h-=1;
                l+=0.5; if (l>1) l-=1;
                color = QColor::fromHslF(h, s, l, a);
            }
            break;
        }
        case Type::Grayscale: {
            if (val.toBool()) {
                auto a = color.alphaF();
                color = qGray(color.rgba());
                color.setAlphaF(a);
            }
            break;
        }
        case Type::Alpha: {
            color.setAlphaF(val.toDouble()/100);
            break;
        }
        case Type::AlphaOffset: {
            auto a = color.alphaF();
            a += val.toDouble()/100;
            if (a < 0) a = 0;
            if (a > 1) a = 1;
            color.setAlphaF(a);
            break;
        }
        case Type::AlphaModulation: {
            auto a = color.alphaF();
            a *= val.toDouble()/100;
            if (a < 0) a = 0;
            if (a > 1) a = 1;
            color.setAlphaF(a);
            break;
        }
        case Type::Hue: {
            auto s = color.hslSaturationF();
            auto l = color.lightnessF();
            auto a = color.alphaF();
            color = QColor::fromHslF(val.toDouble()/360, s, l, a);
            break;
        }
        case Type::HueOffset: {
            auto h = color.hslHueF();
            auto s = color.hslSaturationF();
            auto l = color.lightnessF();
            auto a = color.alphaF();
            h += val.toDouble()/360;
            if (h > 1) h -= 1;
            if (h < 0) h += 1;
            color = QColor::fromHslF(h, s, l, a);
            break;
        }
        case Type::HueModulation: {
            auto h = color.hslHueF();
            auto s = color.hslSaturationF();
            auto l = color.lightnessF();
            auto a = color.alphaF();
            h *= val.toDouble()/100;
            if (h > 1) h -= 1;
            color = QColor::fromHslF(h, s, l, a);
            break;
        }
        case Type::Saturation: {
            auto h = color.hslHueF();
            auto l = color.lightnessF();
            auto a = color.alphaF();
            color = QColor::fromHslF(h, val.toDouble()/100, l, a);
            break;
        }
        case Type::SaturationOffset: {
            auto h = color.hslHueF();
            auto s = color.hslSaturationF();
            auto l = color.lightnessF();
            auto a = color.alphaF();
            s += val.toDouble()/100;
            if (s > 1) s -= 1;
            if (s < 0) s += 1;
            color = QColor::fromHslF(h, s, l, a);
            break;
        }
        case Type::SaturationModulation: {
            auto h = color.hslHueF();
            auto s = color.hslSaturationF();
            auto l = color.lightnessF();
            auto a = color.alphaF();
            s *= val.toDouble()/100;
            if (s > 1) s -= 1;
            if (s < 0) s += 1;
            color = QColor::fromHslF(h, s, l, a);
            break;
        }
        case Type::Luminance: {
            auto h = color.hslHueF();
            auto s = color.hslSaturationF();
            auto a = color.alphaF();
            color = QColor::fromHslF(h, s, val.toDouble()/100, a);
            break;
        }
        case Type::LuminanceOffset: {
            auto h = color.hslHueF();
            auto s = color.hslSaturationF();
            auto l = color.lightnessF();
            auto a = color.alphaF();
            l += val.toDouble()/100;
            if (l > 1) l -= 1;
            if (l < 0) l += 1;
            color = QColor::fromHslF(h, s, l, a);
            break;
        }
        case Type::LuminanceModulation: {
            auto h = color.hslHueF();
            auto s = color.hslSaturationF();
            auto l = color.lightnessF();
            auto a = color.alphaF();
            l *= val.toDouble()/100;
            if (l > 1) l -= 1;
            if (l < 0) l += 1;
            color = QColor::fromHslF(h, s, l, a);
            break;
        }
        case Type::Red: {
            color.setRedF(val.toDouble()/100); break;
        }
        case Type::RedOffset: {
            auto r = color.redF();
            r += val.toDouble()/100;
            if (r < 0) r = 0;
            if (r > 1) r = 1;
            color.setRedF(r);
            break;
        }
        case Type::RedModulation: {
            auto r = color.redF();
            r *= val.toDouble()/100;
            if (r < 0) r = 0;
            if (r > 1) r = 1;
            color.setRedF(r);
            break;
        }
        case Type::Green: {
            color.setGreenF(val.toDouble()/100);
            break;
        }
        case Type::GreenOffset: {
            auto r = color.greenF();
            r += val.toDouble()/100;
            if (r < 0) r = 0;
            if (r > 1) r = 1;
            color.setGreenF(r);
            break;
        }
        case Type::GreenModulation: {
            auto r = color.greenF();
            r *= val.toDouble()/100;
            if (r < 0) r = 0;
            if (r > 1) r = 1;
            color.setGreenF(r);
            break;
        }
        case Type::Blue: {
            color.setBlueF(val.toDouble()/100);
            break;
        }
        case Type::BlueOffset: {
            auto r = color.blueF();
            r += val.toDouble()/100;
            if (r < 0) r = 0;
            if (r > 1) r = 1;
            color.setBlueF(r);
            break;
        }
        case Type::BlueModulation: {
            auto r = color.blueF();
            r *= val.toDouble()/100;
            if (r < 0) r = 0;
            if (r > 1) r = 1;
            color.setBlueF(r);
            break;
        }
        case Type::Gamma: {
            if (val.toBool()) {
                auto gamma = [](double val) -> double {
                    return (val <= 0.0031308 ? val*12.92 : 1.055*std::pow(val, 1.0 / 2.4) - 0.055);
                };
                color.setRedF(gamma(color.redF()));
                color.setGreenF(gamma(color.greenF()));
                color.setBlueF(gamma(color.blueF()));
            }
            break;
        }
        case Type::InverseGamma: {
            if (val.toBool()) {
                auto invgamma = [](double val) -> double {
                    return (val <= 0.04045 ? val/12.92 : std::pow((val+0.055)/1.055, 2.4) );
                };
                color.setRedF(invgamma(color.redF()));
                color.setGreenF(invgamma(color.greenF()));
                color.setBlueF(invgamma(color.blueF()));
            }
            break;
        }
    }
    return color;
}

QColor ColorTransform::transformed(const QColor &inputColor) const
{
    QColor result = inputColor;
    for (auto key: vals.keys()) {
        applyTransform(key, vals.value(key), result);
    }

    return result;
}

void ColorTransform::read(QXmlStreamReader &reader)
{
    const auto &name = reader.name();

    while (!reader.atEnd()) {
        auto token = reader.readNext();
        if (token == QXmlStreamReader::StartElement) {
            auto val = reader.attributes().value(QLatin1String("val"));
            if (reader.name() == QLatin1String("alpha"))
                vals.insert(Type::Alpha, fromST_Percent(val));
            if (reader.name() == QLatin1String("alphaOff"))
                vals.insert(Type::AlphaOffset, fromST_Percent(val));
            if (reader.name() == QLatin1String("alphaMod"))
                vals.insert(Type::AlphaModulation, fromST_Percent(val));
            if (reader.name() == QLatin1String("blue"))
                vals.insert(Type::Blue, fromST_Percent(val));
            if (reader.name() == QLatin1String("blueOff"))
                vals.insert(Type::BlueOffset, fromST_Percent(val));
            if (reader.name() == QLatin1String("blueMod"))
                vals.insert(Type::BlueModulation, fromST_Percent(val));
            if (reader.name() == QLatin1String("comp"))
                vals.insert(Type::Complement, true);
            if (reader.name() == QLatin1String("gamma"))
                vals.insert(Type::Gamma, true);
            if (reader.name() == QLatin1String("gray"))
                vals.insert(Type::Grayscale, true);
            if (reader.name() == QLatin1String("green"))
                vals.insert(Type::Green, fromST_Percent(val));
            if (reader.name() == QLatin1String("greenOff"))
                vals.insert(Type::GreenOffset, fromST_Percent(val));
            if (reader.name() == QLatin1String("greenMod"))
                vals.insert(Type::GreenModulation, fromST_Percent(val));
            if (reader.name() == QLatin1String("hue"))
                vals.insert(Type::Hue, double(val.toULongLong()) / 60000.0);
            if (reader.name() == QLatin1String("hueMod"))
                vals.insert(Type::HueModulation, fromST_Percent(val));
            if (reader.name() == QLatin1String("hueOff"))
                vals.insert(Type::HueOffset, fromST_Percent(val));
            if (reader.name() == QLatin1String("inv"))
                vals.insert(Type::Inverse, true);
            if (reader.name() == QLatin1String("invGamma"))
                vals.insert(Type::InverseGamma, true);
            if (reader.name() == QLatin1String("lum"))
                vals.insert(Type::Luminance, fromST_Percent(val));
            if (reader.name() == QLatin1String("lumOff"))
                vals.insert(Type::LuminanceOffset, fromST_Percent(val));
            if (reader.name() == QLatin1String("lumMod"))
                vals.insert(Type::LuminanceModulation, fromST_Percent(val));
            if (reader.name() == QLatin1String("red"))
                vals.insert(Type::Red, fromST_Percent(val));
            if (reader.name() == QLatin1String("redOff"))
                vals.insert(Type::RedOffset, fromST_Percent(val));
            if (reader.name() == QLatin1String("redMod"))
                vals.insert(Type::RedModulation, fromST_Percent(val));
            if (reader.name() == QLatin1String("sat"))
                vals.insert(Type::Saturation, fromST_Percent(val));
            if (reader.name() == QLatin1String("satOff"))
                vals.insert(Type::SaturationOffset, fromST_Percent(val));
            if (reader.name() == QLatin1String("satMod"))
                vals.insert(Type::SaturationModulation, fromST_Percent(val));
            if (reader.name() == QLatin1String("shade"))
                vals.insert(Type::Shade, fromST_Percent(val));
            if (reader.name() == QLatin1String("tint"))
                vals.insert(Type::Tint, fromST_Percent(val));
            else reader.skipCurrentElement();
        }
        else if (token == QXmlStreamReader::EndElement && reader.name() == name)
            break;
    }
}

void ColorTransform::write(QXmlStreamWriter &writer) const
{
    for (auto i = vals.constBegin(); i!= vals.constEnd(); ++i) {
        switch (i.key()) {
            case Type::Tint:
                writer.writeEmptyElement("a:tint");
                writer.writeAttribute("val", QString::number(i.value().toDouble(), 'f', 2)+"%");
                break;
            case Type::Shade:
                writer.writeEmptyElement("a:shade");
                writer.writeAttribute("val", QString::number(i.value().toDouble(), 'f', 2)+"%");
                break;
            case Type::Complement:
                if (i.value().toBool()) writer.writeEmptyElement("a:comp");
                break;
            case Type::Inverse:
                if (i.value().toBool()) writer.writeEmptyElement("a:inv");
                break;
            case Type::Grayscale:
                if (i.value().toBool()) writer.writeEmptyElement("a:gray");
                break;
            case Type::Alpha:
                writer.writeEmptyElement("a:alpha");
                writer.writeAttribute("val", QString::number(i.value().toDouble(), 'f', 2)+"%");
                break;
            case Type::AlphaOffset:
                writer.writeEmptyElement("a:alphaOff");
                writer.writeAttribute("val", QString::number(i.value().toDouble(), 'f', 2)+"%");
                break;
            case Type::AlphaModulation:
                writer.writeEmptyElement("a:alphaMod");
                writer.writeAttribute("val", QString::number(i.value().toDouble(), 'f', 2)+"%");
                break;
            case Type::Hue:
                writer.writeEmptyElement("a:hue");
                writer.writeAttribute("val", QString::number(qint64(i.value().toDouble()*60000)));
                break;
            case Type::HueOffset:
                writer.writeEmptyElement("a:hueOff");
                writer.writeAttribute("val", QString::number(qint64(i.value().toDouble()*60000)));
                break;
            case Type::HueModulation:
                writer.writeEmptyElement("a:hueMod");
                writer.writeAttribute("val", QString::number(i.value().toDouble(), 'f', 2)+"%");
                break;
            case Type::Saturation:
                writer.writeEmptyElement("a:sat");
                writer.writeAttribute("val", QString::number(i.value().toDouble(), 'f')+"%");
                break;
            case Type::SaturationOffset:
                writer.writeEmptyElement("a:satOff");
                writer.writeAttribute("val", QString::number(i.value().toDouble(), 'f')+"%");
                break;
            case Type::SaturationModulation:
                writer.writeEmptyElement("a:satMod");
                writer.writeAttribute("val", QString::number(i.value().toDouble(), 'f')+"%");
                break;
            case Type::Luminance:
                writer.writeEmptyElement("a:lum");
                writer.writeAttribute("val", QString::number(i.value().toDouble(), 'f')+"%");
                break;
            case Type::LuminanceOffset:
                writer.writeEmptyElement("a:lumOff");
                writer.writeAttribute("val", QString::number(i.value().toDouble(), 'f')+"%");
                break;
            case Type::LuminanceModulation:
                writer.writeEmptyElement("a:lumMod");
                writer.writeAttribute("val", QString::number(i.value().toDouble(), 'f')+"%");
                break;
            case Type::Red:
                writer.writeEmptyElement("a:red");
                writer.writeAttribute("val", QString::number(i.value().toDouble(), 'f')+"%");
                break;
            case Type::RedOffset:
                writer.writeEmptyElement("a:redOff");
                writer.writeAttribute("val", QString::number(i.value().toDouble(), 'f')+"%");
                break;
            case Type::RedModulation:
                writer.writeEmptyElement("a:redMod");
                writer.writeAttribute("val", QString::number(i.value().toDouble(), 'f')+"%");
                break;
            case Type::Green:
                writer.writeEmptyElement("a:green");
                writer.writeAttribute("val", QString::number(i.value().toDouble(), 'f')+"%");
                break;
            case Type::GreenOffset:
                writer.writeEmptyElement("a:greenOff");
                writer.writeAttribute("val", QString::number(i.value().toDouble(), 'f')+"%");
                break;
            case Type::GreenModulation:
                writer.writeEmptyElement("a:greenMod");
                writer.writeAttribute("val", QString::number(i.value().toDouble(), 'f')+"%");
                break;
            case Type::Blue:
                writer.writeEmptyElement("a:blue");
                writer.writeAttribute("val", QString::number(i.value().toDouble(), 'f')+"%");
                break;
            case Type::BlueOffset:
                writer.writeEmptyElement("a:blueOff");
                writer.writeAttribute("val", QString::number(i.value().toDouble(), 'f')+"%");
                break;
            case Type::BlueModulation:
                writer.writeEmptyElement("a:blueMod");
                writer.writeAttribute("val", QString::number(i.value().toDouble(), 'f')+"%");
                break;
            case Type::Gamma:
                if (i.value().toBool()) writer.writeEmptyElement("a:gamma");
                break;
            case Type::InverseGamma:
                if (i.value().toBool()) writer.writeEmptyElement("a:invGamma");
                break;
        }
    }
}

ColorTransform::operator QVariant() const
{
    const auto& cref
#if QT_VERSION >= 0x060000 // Qt 6.0 or over
        = QMetaType::fromType<ColorTransform>();
#else
        = qMetaTypeId<ColorTransform>() ;
#endif
    QVariant v(cref, this);
    return v;
}

QByteArray Color::idKey() const
{
    if (isDirty) {
        QByteArray bytes;
        QXmlStreamWriter w(&bytes);
        write(w);

        m_key = bytes;
        isDirty = false;
    }

    return m_key;
}

uint qHash(const Color &c, uint seed) noexcept
{
    return qHash(c.idKey(), seed);
}

#if !defined(QT_NO_DATASTREAM)
QDataStream &operator<<(QDataStream &s, const Color &color)
{
    s << static_cast<int>(color.type_);
    s << color.val;
    s << color.lastColor;
    s << color.tr;
    return s;
}

QDataStream &operator>>(QDataStream &s, Color &color)
{
    int t;
    s >> t;
    color.type_ = static_cast<Color::Type>(t);
    s >> color.val;
    s >> color.lastColor;
    s >> color.tr;

    return s;
}

QDataStream &operator<<(QDataStream &s, const ColorTransform &tr)
{
    const auto &keys = tr.vals.keys();
    using st = decltype(tr.vals)::size_type;
    st t = keys.size();
    s << t;
    for (auto key: qAsConst(keys)) s << static_cast<int>(key) << tr.vals.value(key);
    return s;
}

QDataStream &operator>>(QDataStream &s, ColorTransform &tr)
{
    using st = decltype(tr.vals)::size_type;
    st t;
    s >> t;
    for (st i = 0; i < t; ++i) {
        int key;
        s >> key;
        QVariant val;
        s >> val;
        tr.vals.insert(static_cast<ColorTransform::Type>(key), val);
    }

    return s;
}

#endif

}

