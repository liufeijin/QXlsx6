#include "xlsxfillformat.h"
#include "xlsxutility_p.h"
#include "xlsxmediafile_p.h"
#include "xlsxworkbook.h"
#include <QDebug>
#include <QtMath>
#include <QBuffer>

namespace QXlsx {

class FillFormatPrivate: public QSharedData
{
public:
    FillFormatPrivate();
    FillFormatPrivate(const FillFormatPrivate &other);
    ~FillFormatPrivate();

    void parse(const QBrush &brush);
    QBrush toBrush() const;

    FillFormat::FillType type;
    Color color;

    //Gradient fill properties
    QMap<double, Color> gradientList;
    //linear gradient
    Angle linearShadeAngle; //0..360
    std::optional<bool> linearShadeScaled;
    //path gradient
    std::optional<FillFormat::PathShadeType> pathShadeType;
    std::optional<RelativeRect> pathShadeRect;
    std::optional<RelativeRect> tileRect;
    std::optional<FillFormat::TileFlipMode> tileFlipMode;
    std::optional<bool> rotWithShape;

    //Pattern fill properties
    Color foregroundColor;
    Color backgroundColor;
    std::optional<FillFormat::PatternType> patternType;

    //Blip fill - limited support
    std::optional<int> dpi;
    QSharedPointer<MediaFile> blipImage;
    QString blipID; // internal
    double alpha;
    std::optional<FillFormat::BlipCompression> compression;

    bool operator==(const FillFormatPrivate &other) const;
};

FillFormatPrivate::FillFormatPrivate()
{

}

FillFormatPrivate::FillFormatPrivate(const FillFormatPrivate &other)
    : QSharedData(other),
      type(other.type),
      color(other.color),
      gradientList(other.gradientList),
      linearShadeAngle(other.linearShadeAngle),
      linearShadeScaled(other.linearShadeScaled),
      pathShadeType(other.pathShadeType),
      pathShadeRect(other.pathShadeRect),
      tileRect(other.tileRect),
      tileFlipMode(other.tileFlipMode),
      rotWithShape(other.rotWithShape),
      foregroundColor(other.foregroundColor),
      backgroundColor(other.backgroundColor),
      patternType(other.patternType)
{

}

FillFormatPrivate::~FillFormatPrivate()
{

}

bool FillFormatPrivate::operator==(const FillFormatPrivate &other) const
{
    if (type != other.type) return false;
    if (color != other.color) return false;
    if (gradientList != other.gradientList) return false;
    if (linearShadeAngle != other.linearShadeAngle) return false;
    if (linearShadeScaled != other.linearShadeScaled) return false;
    if (pathShadeType != other.pathShadeType) return false;
    if (pathShadeRect != other.pathShadeRect) return false;
    if (tileRect != other.tileRect) return false;
    if (tileFlipMode != other.tileFlipMode) return false;
    if (rotWithShape != other.rotWithShape) return false;
    if (foregroundColor != other.foregroundColor) return false;
    if (backgroundColor != other.backgroundColor) return false;
    if (patternType != other.patternType) return false;
    return true;
}

void FillFormatPrivate::parse(const QBrush &brush)
{
    switch (brush.style()) {
        case Qt::NoBrush: type = FillFormat::FillType::NoFill; break;
        case Qt::SolidPattern: {
            type = FillFormat::FillType::SolidFill;
            color = brush.color(); break;
        }
        case Qt::Dense1Pattern: {
            type = FillFormat::FillType::PatternFill;
            patternType = FillFormat::PatternType::Percent80;
            foregroundColor = brush.color();
            break;
        }
        case Qt::Dense2Pattern: {
            type = FillFormat::FillType::PatternFill;
            patternType = FillFormat::PatternType::Percent75;
            foregroundColor = brush.color();
            break;
        }
        case Qt::Dense3Pattern: {
            type = FillFormat::FillType::PatternFill;
            patternType = FillFormat::PatternType::Percent60;
            foregroundColor = brush.color();
            break;
        }
        case Qt::Dense4Pattern: {
            type = FillFormat::FillType::PatternFill;
            patternType = FillFormat::PatternType::Percent50;
            foregroundColor = brush.color();
            break;
        }
        case Qt::Dense5Pattern: {
            type = FillFormat::FillType::PatternFill;
            patternType = FillFormat::PatternType::Percent30;
            foregroundColor = brush.color();
            break;
        }
        case Qt::Dense6Pattern: {
            type = FillFormat::FillType::PatternFill;
            patternType = FillFormat::PatternType::Percent20;
            foregroundColor = brush.color();
            break;
        }
        case Qt::Dense7Pattern: {
            type = FillFormat::FillType::PatternFill;
            patternType = FillFormat::PatternType::Percent5;
            foregroundColor = brush.color();
            break;
        }
        case Qt::HorPattern: {
            type = FillFormat::FillType::PatternFill;
            patternType = FillFormat::PatternType::Horizontal;
            foregroundColor = brush.color();
            break;
        }
        case Qt::VerPattern: {
            type = FillFormat::FillType::PatternFill;
            patternType = FillFormat::PatternType::Vertical;
            foregroundColor = brush.color();
            break;
        }
        case Qt::CrossPattern: {
            type = FillFormat::FillType::PatternFill;
            patternType = FillFormat::PatternType::Cross;
            foregroundColor = brush.color();
            break;
        }
        case Qt::BDiagPattern: {
            type = FillFormat::FillType::PatternFill;
            patternType = FillFormat::PatternType::UpwardDiagonal;
            foregroundColor = brush.color();
            break;
        }
        case Qt::FDiagPattern: {
            type = FillFormat::FillType::PatternFill;
            patternType = FillFormat::PatternType::DownwardDiagonal;
            foregroundColor = brush.color();
            break;
        }
        case Qt::DiagCrossPattern: {
            type = FillFormat::FillType::PatternFill;
            patternType = FillFormat::PatternType::OpenDiamond;
            foregroundColor = brush.color();
            break;
        }
        case Qt::LinearGradientPattern: {
            const auto gradient = brush.gradient();
            if (gradient->type() == QGradient::LinearGradient) {
                auto mode = gradient->coordinateMode();
                const auto stops = gradient->stops();
                auto spread = gradient->spread();
                auto start = static_cast<const QLinearGradient*>(gradient)->start();
                auto end = static_cast<const QLinearGradient*>(gradient)->finalStop();

                //vector from start to end will give us the gradient angle
                auto angle = atan2((end-start).y(), (end-start).x())/M_PI*360; //gradient angle
                linearShadeAngle = Angle(angle);
                if (mode == QGradient::LogicalMode)
                    linearShadeScaled = false;
                if (mode == QGradient::ObjectMode)
                    linearShadeScaled = true;
                for (auto st: stops)
                    gradientList.insert(st.first, st.second);
                switch (spread) {
                    case QGradient::PadSpread: tileFlipMode = FillFormat::TileFlipMode::None; break;
                    case QGradient::RepeatSpread: tileFlipMode = FillFormat::TileFlipMode::XY; break;
                    case QGradient::ReflectSpread: tileFlipMode = FillFormat::TileFlipMode::XY; break;
                }
            }
            break;
        }
        case Qt::RadialGradientPattern:
            break;
        case Qt::ConicalGradientPattern:
            break;
        case Qt::TexturePattern: //TODO: blip fill
            break;

    }
}

QBrush FillFormatPrivate::toBrush() const
{
    QBrush brush;

    return brush;
}

FillFormat::FillFormat()
{

}

FillFormat::FillFormat(FillFormat::FillType type)
{
    d = new FillFormatPrivate;
    d->type = type;
}

FillFormat::FillFormat(const QBrush &brush)
{
    d = new FillFormatPrivate;
    d->parse(brush);
}

FillFormat::FillFormat(const FillFormat &other) : d(other.d)
{

}

FillFormat &FillFormat::operator=(const FillFormat &rhs)
{
    if (*this != rhs)
        d = rhs.d;
    return *this;
}

FillFormat::~FillFormat()
{

}

void FillFormat::setBrush(const QBrush &brush)
{
    d = new FillFormatPrivate;
    d->parse(brush);
}

QBrush FillFormat::toBrush() const
{
    if (d) return d->toBrush();
    return {};
}

FillFormat::FillType FillFormat::type() const
{
    if (d) return d->type;
    return FillType::NoFill;
}

void FillFormat::setType(FillFormat::FillType type)
{
    if (!d) d = new FillFormatPrivate;
    d->type = type;
}

Color FillFormat::color() const
{
    if (d) return d->color;
    return Color();
}

void FillFormat::setColor(const Color &color)
{
    if (!d) {
        d = new FillFormatPrivate;
        d->type = FillType::SolidFill;
    }
    d->color = color;
}

void FillFormat::setColor(const QColor &color)
{
    if (!d) {
        d = new FillFormatPrivate;
        d->type = FillType::SolidFill;
    }
    d->color = color;
}

QMap<double, Color> FillFormat::gradientList() const
{
    if (d) return d->gradientList;
    return {};
}

void FillFormat::setGradientList(const QMap<double, Color> &list)
{
    if (!d) d = new FillFormatPrivate;
    d->gradientList = list;
}

void FillFormat::addGradientStop(double stop, const Color &color)
{
    if (!d) d = new FillFormatPrivate;
    d->gradientList.insert(stop, color);
}

Angle FillFormat::linearShadeAngle() const
{
    if (d)  return d->linearShadeAngle;
    return {};
}

void FillFormat::setLinearShadeAngle(Angle val)
{
    if (!d) d = new FillFormatPrivate;
    d->linearShadeAngle = val;
}

std::optional<bool> FillFormat::linearShadeScaled() const
{
    if (d)  return d->linearShadeScaled;
    return {};
}

void FillFormat::setLinearShadeScaled(bool scaled)
{
    if (!d) d = new FillFormatPrivate;
    d->linearShadeScaled = scaled;
}

std::optional<FillFormat::PathShadeType> FillFormat::pathShadeType() const
{
    if (d)  return d->pathShadeType;
    return {};
}

void FillFormat::setPathShadeType(FillFormat::PathShadeType pathShadeType)
{
    if (!d) d = new FillFormatPrivate;
    d->pathShadeType = pathShadeType;
}

std::optional<RelativeRect> FillFormat::pathShadeRect() const
{
    if (d)  return d->pathShadeRect;
    return {};
}

void FillFormat::setPathShadeRect(RelativeRect rect)
{
    if (!d) d = new FillFormatPrivate;
    d->pathShadeRect = rect;
}

std::optional<RelativeRect> FillFormat::tileRect() const
{
    if (d)  return d->tileRect;
    return {};
}

void FillFormat::setTileRect(RelativeRect rect)
{
    if (!d) d = new FillFormatPrivate;
    d->tileRect = rect;
}

std::optional<FillFormat::TileFlipMode> FillFormat::tileFlipMode() const
{
    if (d)  return d->tileFlipMode;
    return {};
}

void FillFormat::setTileFlipMode(FillFormat::TileFlipMode tileFlipMode)
{
    if (!d) d = new FillFormatPrivate;
    d->tileFlipMode = tileFlipMode;
}

std::optional<bool> FillFormat::rotateWithShape() const
{
    if (d)  return d->rotWithShape;
    return {};
}

void FillFormat::setRotateWithShape(bool val)
{
    if (!d) d = new FillFormatPrivate;
    d->rotWithShape = val;
}

Color FillFormat::foregroundColor() const
{
    if (d) return d->foregroundColor;
    return Color();
}

void FillFormat::setForegroundColor(const Color &color)
{
    if (!d) d = new FillFormatPrivate;
    d->foregroundColor = color;
}

Color FillFormat::backgroundColor() const
{
    if (d) return d->backgroundColor;
    return Color();
}

void FillFormat::setBackgroundColor(const Color &color)
{
    if (!d) d = new FillFormatPrivate;
    d->backgroundColor = color;
}

std::optional<FillFormat::PatternType> FillFormat::patternType()
{
    if (d) return d->patternType;
    return {};
}

void FillFormat::setPatternType(FillFormat::PatternType patternType)
{
    if (!d) d = new FillFormatPrivate;
    d->patternType = patternType;
}

void FillFormat::setPicture(const QImage &picture)
{
    if (!d) d = new FillFormatPrivate;
    QByteArray ba;
    QBuffer buffer(&ba);
    buffer.open(QIODevice::WriteOnly);
    picture.save(&buffer, "PNG");

    d->blipImage = QSharedPointer<MediaFile>(new MediaFile(ba, QStringLiteral("png"), QStringLiteral("image/png")));
}

QImage FillFormat::picture() const
{
    if (d && d->blipImage) {
        QImage pic;
        pic.loadFromData(d->blipImage->contents());
        return pic;
    }
    return {};
}

bool FillFormat::isValid() const
{
    if (d) return true;
    return false;
}

void FillFormat::setPictureID(int id)
{
    if (!d) d = new FillFormatPrivate;
    d->blipID = QString("rId%1").arg(id);
}

int FillFormat::registerBlip(Workbook *workbook)
{
    if (!d || !d->blipImage) return -1;
    if (d->blipImage->contents().isEmpty()) return -1;

    workbook->addMediaFile(d->blipImage);
    return d->blipImage->index();
}

void FillFormat::write(QXmlStreamWriter &writer) const
{
    switch (d->type) {
        case FillType::NoFill : writeNoFill(writer); break;
        case FillType::SolidFill : writeSolidFill(writer); break;
        case FillType::GradientFill : writeGradientFill(writer); break;
        case FillType::BlipFill : writeBlipFill(writer); break;
        case FillType::PatternFill : writePatternFill(writer); break;
        case FillType::GroupFill : writeGroupFill(writer); break;
    }
}

void FillFormat::read(QXmlStreamReader &reader)
{
    const QString &name = reader.name().toString();
    FillType t;
    fromString(name, t);

    if (!d) {
        d = new FillFormatPrivate;
        d->type = t;
    }

    switch (d->type) {
        case FillType::NoFill: break; //NO-OP
        case FillType::SolidFill : readSolidFill(reader); break;

        case FillType::GradientFill : readGradientFill(reader); break;
        //TODO: blip fill type
//        case FillType::BlipFill : readBlipFill(reader); break;
        case FillType::PatternFill : readPatternFill(reader); break;
        case FillType::GroupFill : readGroupFill(reader); break;
        default: break;
    }
}

bool FillFormat::operator==(const FillFormat &other) const
{
    if (d == other.d) return true;
    if (!d || !other.d) return false;
    return *this->d.constData() == *other.d.constData();
}

bool FillFormat::operator!=(const FillFormat &other) const
{
    return !(operator==(other));
}

FillFormat::operator QVariant() const
{
    const auto& cref
#if QT_VERSION >= 0x060000 // Qt 6.0 or over
        = QMetaType::fromType<FillFormat>();
#else
        = qMetaTypeId<FillFormat>() ;
#endif
    return QVariant(cref, this);
}

void FillFormat::readNoFill(QXmlStreamReader &reader)
{
    Q_ASSERT(reader.name() == QLatin1String("noFill"));

    //NO_OP
}

void FillFormat::readSolidFill(QXmlStreamReader &reader)
{
    Q_ASSERT(reader.name() == QLatin1String("solidFill"));

    while (!reader.atEnd()) {
        auto token = reader.readNext();
        if (token == QXmlStreamReader::StartElement)
            d->color.read(reader);
        else if (token == QXmlStreamReader::EndElement && reader.name() == QLatin1String("solidFill"))
            break;
    }
}

void FillFormat::readGradientFill(QXmlStreamReader &reader)
{
    Q_ASSERT(reader.name() == QLatin1String("gradFill"));

    const auto &attr = reader.attributes();
    if (attr.hasAttribute(QLatin1String("flip"))) {
        TileFlipMode t;
        fromString(attr.value(QLatin1String("flip")).toString(), t);
        d->tileFlipMode = t;
    }
    parseAttributeBool(attr, QLatin1String("rotWithShape"), d->rotWithShape);

//    <xsd:complexType name="CT_GradientFillProperties">
//        <xsd:sequence>
//          <xsd:element name="gsLst" type="CT_GradientStopList" minOccurs="0" maxOccurs="1"/>
//          <xsd:group ref="EG_ShadeProperties" minOccurs="0" maxOccurs="1"/>
//          <xsd:element name="tileRect" type="CT_RelativeRect" minOccurs="0" maxOccurs="1"/>
//        </xsd:sequence>
//        <xsd:attribute name="flip" type="ST_TileFlipMode" use="optional" default="none"/>
//        <xsd:attribute name="rotWithShape" type="xsd:boolean" use="optional"/>
//      </xsd:complexType>

    while (!reader.atEnd()) {
        auto token = reader.readNext();
        if (token == QXmlStreamReader::StartElement) {
            if (reader.name() == QLatin1String("gsLst")) {
                readGradientList(reader);
            }
            else if (reader.name() == QLatin1String("lin")) {
                //linear shade properties
                const auto &attr = reader.attributes();
                parseAttribute(attr, QLatin1String("ang"), d->linearShadeAngle);
                parseAttributeBool(attr, QLatin1String("scaled"), d->linearShadeScaled);
            }
            else if (reader.name() == QLatin1String("path")) {
                //path shade properties
                const auto &attr = reader.attributes();
                if (attr.hasAttribute(QLatin1String("path"))) {
                    auto s = attr.value(QLatin1String("path"));
                    if (s == QLatin1String("shape")) d->pathShadeType = PathShadeType::Shape;
                    else if (s == QLatin1String("circle")) d->pathShadeType = PathShadeType::Circle;
                    else if (s == QLatin1String("rect")) d->pathShadeType = PathShadeType::Rectangle;
                }

                reader.readNextStartElement();
                if (reader.name() == QLatin1String("fillToRect")) {
                    RelativeRect r;
                    r.read(reader);
                    d->pathShadeRect = r;
                }
            }
            else if (reader.name() == QLatin1String("tileRect")) {
                RelativeRect r;
                r.read(reader);
                d->tileRect = r;
            }
            else reader.skipCurrentElement();
        }
        else if (token == QXmlStreamReader::EndElement && reader.name() == QLatin1String("gradFill"))
            break;
    }
}

void FillFormat::writeNoFill(QXmlStreamWriter &writer) const
{
    writer.writeEmptyElement("a:noFill");
}

void FillFormat::writeSolidFill(QXmlStreamWriter &writer) const
{
    writer.writeStartElement("a:solidFill");
    d->color.write(writer);
    writer.writeEndElement();
}

void FillFormat::writeGradientFill(QXmlStreamWriter &writer) const
{
    writer.writeStartElement(QLatin1String("a:gradFill"));
    if (d->tileFlipMode.has_value()) {
        QString s;
        toString(d->tileFlipMode.value(), s);
        writer.writeAttribute("flip", s);
    }
    if (d->rotWithShape.has_value()) writer.writeAttribute("rotWithShape", d->rotWithShape.value() ? "true" : "false");
    writeGradientList(writer);

    if (d->linearShadeAngle.isValid() || d->linearShadeScaled.has_value()) {
        writer.writeEmptyElement("a:lin");
        if (d->linearShadeAngle.isValid())
            writer.writeAttribute("ang", d->linearShadeAngle.toString());
        if (d->linearShadeScaled.has_value())
            writer.writeAttribute("scaled", d->linearShadeScaled.value() ? "true" : "false");
    }

    if (d->pathShadeType.has_value()) {
        writer.writeStartElement("a:path");
        switch (d->pathShadeType.value()) {
            case PathShadeType::Shape: writer.writeAttribute("path", "shape"); break;
            case PathShadeType::Circle: writer.writeAttribute("path", "circle"); break;
            case PathShadeType::Rectangle: writer.writeAttribute("path", "rect"); break;
        }
        if (d->pathShadeRect.has_value())
            d->pathShadeRect->write(writer, "a:fillToRect");
        writer.writeEndElement();
    }
    if (d->tileRect.has_value() && d->tileRect->isValid())
        d->tileRect->write(writer, "a:tileRect");

    writer.writeEndElement();
}

void FillFormat::readPatternFill(QXmlStreamReader &reader)
{
    const auto &name = reader.name();
    if (reader.attributes().hasAttribute(QLatin1String("prst"))) {
        PatternType t;
        fromString(reader.attributes().value(QLatin1String("prst")).toString(), t);
        d->patternType = t;
    }
    while (!reader.atEnd()) {
        auto token = reader.readNext();
        if (token == QXmlStreamReader::StartElement) {
            if (reader.name() == QLatin1String("fgClr")) {
                d->foregroundColor.read(reader);
            }
            else if (reader.name() == QLatin1String("bgClr")) {
                d->backgroundColor.read(reader);
            }
            else reader.skipCurrentElement();
        }
        else if (token == QXmlStreamReader::EndElement && reader.name() == name)
            break;
    }
}

void FillFormat::writePatternFill(QXmlStreamWriter &writer) const
{
    if (!d->backgroundColor.isValid() && !d->foregroundColor.isValid() && !d->patternType.has_value()) {
        writer.writeEmptyElement(QLatin1String("a:pattFill"));
        return;
    }

    writer.writeStartElement(QLatin1String("a:pattFill"));
    if (d->patternType.has_value()) {
        QString s;
        toString(d->patternType.value(), s);
        writer.writeAttribute(QLatin1String("prst"), s);
    }
    if (d->backgroundColor.isValid()) {
        writer.writeStartElement(QLatin1String("a:bgClr"));
        d->backgroundColor.write(writer);
        writer.writeEndElement();
    }
    if (d->foregroundColor.isValid()) {
        writer.writeStartElement(QLatin1String("a:fgClr"));
        d->foregroundColor.write(writer);
        writer.writeEndElement();
    }
    writer.writeEndElement();
}

void FillFormat::readGroupFill(QXmlStreamReader &reader)
{
    Q_UNUSED(reader);
    //no-op
}

void FillFormat::writeGroupFill(QXmlStreamWriter &writer) const
{
    writer.writeEmptyElement(QLatin1String("a:grpFill"));
}

void FillFormat::writeBlipFill(QXmlStreamWriter &writer) const
{
    if (!d->blipImage || d->blipID.isEmpty()) return;
    writer.writeStartElement(QLatin1String("a:blipFill"));
    writeAttribute(writer, QLatin1String("rotWithShape"), d->rotWithShape);
    writeAttribute(writer, QLatin1String("dpi"), d->dpi);
    writer.writeStartElement(QLatin1String("a:blip"));
    writer.writeAttribute(QLatin1String("xmlns:r"), QLatin1String("http://schemas.openxmlformats.org/officeDocument/2006/relationships"));
    writer.writeAttribute(QLatin1String("r:embed"), d->blipID);
    writer.writeEndElement(); //a:blip
    writer.writeEmptyElement(QLatin1String("a:srcRect"));
    writer.writeStartElement(QLatin1String("a:stretch"));
    writer.writeEmptyElement(QLatin1String("a:fillRect"));
    writer.writeEndElement(); //a:stretch
    writer.writeEndElement(); //a:blipFill
}

void FillFormat::readGradientList(QXmlStreamReader &reader)
{
    while (!reader.atEnd()) {
        auto token = reader.readNext();
        if (token == QXmlStreamReader::StartElement) {
            if (reader.name() == QLatin1String("gs")) {
                double pos = fromST_Percent(reader.attributes().value("pos"));
                reader.readNextStartElement();
                Color col;
                col.read(reader);
                d->gradientList.insert(pos, col);
            }
            else reader.skipCurrentElement();
        }
        else if (token == QXmlStreamReader::EndElement && reader.name() == QLatin1String("gsLst"))
            break;
    }
}

void FillFormat::writeGradientList(QXmlStreamWriter &writer) const
{
    if (d->gradientList.isEmpty()) return;
    writer.writeStartElement(QLatin1String("a:gsLst"));
    for (auto i = d->gradientList.constBegin(); i!= d->gradientList.constEnd(); ++i) {
        writer.writeStartElement(QLatin1String("a:gs"));
        writer.writeAttribute("pos", toST_Percent(i.key()));
        i.value().write(writer);
        writer.writeEndElement();
    }
    writer.writeEndElement();
}

#ifndef QT_NO_DEBUG_STREAM
QDebug operator<<(QDebug dbg, const FillFormat &f)
{
    QDebugStateSaver saver(dbg);
    dbg.setAutoInsertSpaces(false);
    dbg << "QXlsx::FillFormat(" ;
    dbg << "type: "<< static_cast<int>(f.d->type) << ", ";
    if (f.d->color.isValid()) dbg << "color: " << f.d->color << ", ";
    dbg << "gradientList: "<< f.d->gradientList << ", ";
    if (f.d->linearShadeAngle.isValid()) dbg << "linearShadeAngle: " << f.d->linearShadeAngle.toString() << ", ";
    if (f.d->linearShadeScaled.has_value()) dbg << "linearShadeScaled: " << f.d->linearShadeScaled.value() << ", ";
    if (f.d->pathShadeType.has_value()) dbg << "pathShadeType: " << static_cast<int>(f.d->pathShadeType.value()) << ", ";
    if (f.d->pathShadeRect.has_value()) dbg << "pathShadeRect: " << f.d->pathShadeRect.value() << ", ";
    if (f.d->tileRect.has_value()) dbg << "tileRect: " << f.d->tileRect.value() << ", ";
    if (f.d->tileFlipMode.has_value())
        dbg << "tileFlipMode: " << static_cast<int>(f.d->tileFlipMode.value()) << ", ";
    if (f.d->rotWithShape.has_value()) dbg << "rotWithShape: " << f.d->rotWithShape.value() << ", ";
    if (f.d->foregroundColor.isValid()) dbg << "foregroundColor: " << f.d->foregroundColor << ", ";
    if (f.d->backgroundColor.isValid()) dbg << "backgroundColor: " << f.d->backgroundColor << ", ";
    if (f.d->patternType.has_value())
        dbg << "patternType: " << static_cast<int>(f.d->patternType.value());

    return dbg;
}
#endif

}




