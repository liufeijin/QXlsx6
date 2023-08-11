#include "xlsxshapeformat.h"

QT_BEGIN_NAMESPACE_XLSX

ShapePrivate::ShapePrivate()
{

}

ShapePrivate::ShapePrivate(const ShapePrivate &other)
    : QSharedData(other) ,
      blackWhiteMode(other.blackWhiteMode),
      xfrm(other.xfrm),
      presetGeometry(other.presetGeometry),
      fill(other.fill),
      line(other.line)
{

}

ShapePrivate::~ShapePrivate()
{

}

bool ShapePrivate::operator ==(const ShapePrivate &other) const
{
    if (blackWhiteMode != other.blackWhiteMode) return false;
    if (xfrm != other.xfrm) return false;
    if (presetGeometry != other.presetGeometry) return false;
    if (fill != other.fill) return false;
    if (line != other.line) return false;

    return true;
}

std::optional<ShapeFormat::BlackWhiteMode> ShapeFormat::blackWhiteMode() const
{
    if (d) return d->blackWhiteMode;
    return {};
}

void ShapeFormat::setBlackWhiteMode(ShapeFormat::BlackWhiteMode val)
{
    if (!d) d = new ShapePrivate;
    d->blackWhiteMode = val;
}

std::optional<Transform2D> ShapeFormat::transform2D() const
{
    if (d) return d->xfrm;
    return {};
}

void ShapeFormat::setTransform2D(Transform2D val)
{
    if (!d) d = new ShapePrivate;
    d->xfrm = val;
}

std::optional<PresetGeometry2D> ShapeFormat::presetGeometry() const
{
    if (d) return d->presetGeometry;
    return {};
}

void ShapeFormat::setPresetGeometry(PresetGeometry2D val)
{
    if (!d) d = new ShapePrivate;
    d->presetGeometry = val;
}

FillFormat &ShapeFormat::fill()
{
    if (!d) d = new ShapePrivate;
    return d->fill;
}

FillFormat ShapeFormat::fill() const
{
    if (d) return d->fill;
    return {};
}

void ShapeFormat::setFill(const FillFormat &val)
{
    if (!d) d = new ShapePrivate;
    d->fill = val;
}

LineFormat ShapeFormat::line() const
{
    if (d) return d->line;
    return {};
}

LineFormat &ShapeFormat::line()
{
    if (!d) d = new ShapePrivate;
    return d->line;
}

void ShapeFormat::setLine(const LineFormat &line)
{
    if (!d) d = new ShapePrivate;
    d->line = line;
}

Effect ShapeFormat::effectList() const
{
    if (d) return d->effectList;
    return {};
}

void ShapeFormat::setEffectList(const Effect &effectList)
{
    if (!d) d = new ShapePrivate;
    d->effectList = effectList;
}

std::optional<Scene3D> ShapeFormat::scene3D() const
{
    if (d) return d->scene3D;
    return {};
}

void ShapeFormat::setScene3D(const Scene3D scene3D)
{
    if (!d) d = new ShapePrivate;
    d->scene3D = scene3D;
}

std::optional<Shape3D> ShapeFormat::shape3D() const
{
    if (d) return d->shape3D;
    return {};
}

void ShapeFormat::setShape3D(const Shape3D &shape3D)
{
    if (!d) d = new ShapePrivate;
    d->shape3D = shape3D;
}

bool ShapeFormat::isValid() const
{
    if (d) return true;
    return false;
}

void ShapeFormat::write(QXmlStreamWriter &writer, const QString &name) const
{
    if (!d) return;

    writer.writeStartElement(name);
    if (d->blackWhiteMode.has_value()) {
        QString s;
        toString(d->blackWhiteMode.value(), s);
        writer.writeAttribute(QLatin1String("bwMode"), s);
    }
    if (d->xfrm.has_value()) d->xfrm->write(writer, QLatin1String("a:xfrm"));
    if (d->presetGeometry.has_value()) {
        d->presetGeometry->write(writer, QLatin1String("a:prstGeom"));
    }
    if (d->fill.isValid()) d->fill.write(writer);
    if (d->line.isValid()) d->line.write(writer, QLatin1String("a:ln"));
    if (d->effectList.isValid()) d->effectList.write(writer);
    if (d->scene3D.has_value()) d->scene3D->write(writer, QLatin1String("a:scene3d"));
    if (d->shape3D.has_value()) d->shape3D->write(writer, QLatin1String("a:sp3d"));

    writer.writeEndElement(); //"a:spPr"
}

void ShapeFormat::read(QXmlStreamReader &reader)
{
    if (!d) d = new ShapePrivate;

    const auto &name = reader.name();

    if (reader.attributes().hasAttribute(QLatin1String("bwMode"))) {
        BlackWhiteMode t;
        fromString(reader.attributes().value(QLatin1String("bwMode")).toString(), t);
        d->blackWhiteMode = t;
    }

    while (!reader.atEnd()) {
        auto token = reader.readNext();
        if (token == QXmlStreamReader::StartElement) {
            if (reader.name() == QLatin1String("xfrm")) {
                Transform2D t;
                t.read(reader);
                d->xfrm = t;
            }
            else if (reader.name() == QLatin1String("prstGeom")) {
                PresetGeometry2D p;
                p.read(reader);
                d->presetGeometry = p;
            }
            else if (reader.name() == QLatin1String("noFill")) {
                d->fill.read(reader);
            }
            else if (reader.name() == QLatin1String("solidFill")) {
                d->fill.read(reader);
            }
            else if (reader.name() == QLatin1String("gradFill")) {
                d->fill.read(reader);
            }
            else if (reader.name() == QLatin1String("blipFill")) {
                d->fill.read(reader);
            }
            else if (reader.name() == QLatin1String("pattFill")) {
                d->fill.read(reader);
            }
            else if (reader.name() == QLatin1String("grpFill")) {
                d->fill.read(reader);
            }
            else if (reader.name() == QLatin1String("ln")) {
                d->line.read(reader);
            }
            else if (reader.name() == QLatin1String("effectLst")) {
                d->effectList.read(reader);
            }
            else if (reader.name() == QLatin1String("effectDag")) {
                // TODO: Effect DAG
            }
            else if (reader.name() == QLatin1String("scene3d")) {
                Scene3D s;
                s.read(reader);
                d->scene3D = s;
            }
            else if (reader.name() == QLatin1String("sp3d")) {
                Shape3D s;
                s.read(reader);
                d->shape3D = s;
            }
        }
        else if (token == QXmlStreamReader::EndElement && reader.name() == name)
            break;
    }
}

bool ShapeFormat::operator ==(const ShapeFormat &other) const
{
    if (d == other.d) return true;
    if (!d || !other.d) return false;
    return *this->d.constData() == *other.d.constData();
}

bool ShapeFormat::operator !=(const ShapeFormat &other) const
{
    return !(operator==(other));
}

QDebug operator<<(QDebug dbg, const ShapeFormat &f)
{
    // TODO: write all members
    QDebugStateSaver saver(dbg);

    dbg.nospace() << "QXlsx::ShapeFormat()";

    return dbg;
}

QT_END_NAMESPACE_XLSX


