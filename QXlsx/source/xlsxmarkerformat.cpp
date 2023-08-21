#include <QtGlobal>
#include <QDataStream>
#include <QDebug>

#include "xlsxmarkerformat.h"

namespace QXlsx {

class MarkerFormatPrivate : public QSharedData
{
public:
    std::optional<int> size; // [2..72]
    std::optional<MarkerFormat::MarkerType> markerType;
    ShapeFormat shape;

    MarkerFormatPrivate();
    MarkerFormatPrivate(const MarkerFormatPrivate &other);
    ~MarkerFormatPrivate();

    bool operator==(const MarkerFormatPrivate &other) const;
};

MarkerFormatPrivate::MarkerFormatPrivate()
{
}

MarkerFormatPrivate::MarkerFormatPrivate(const MarkerFormatPrivate &other)
    : QSharedData(other)

{

}

MarkerFormatPrivate::~MarkerFormatPrivate()
{

}

bool MarkerFormatPrivate::operator==(const MarkerFormatPrivate &other) const
{
    return (size == other.size && markerType == other.markerType
            && shape == other.shape);
}

MarkerFormat::MarkerFormat()
{
    //The d pointer is initialized with a null pointer
}

MarkerFormat::MarkerFormat(MarkerFormat::MarkerType type)
{
    d = new MarkerFormatPrivate;
    d->markerType = type;
}

/*!
   Creates a new format with the same attributes as the \a other format.
 */
MarkerFormat::MarkerFormat(const MarkerFormat &other)
    :d(other.d)
{

}

/*!
   Assigns the \a other format to this format, and returns a
   reference to this format.
 */
MarkerFormat &MarkerFormat::operator =(const MarkerFormat &other)
{
    if (*this != other) d = other.d;
    return *this;
}

/*!
 * Destroys this format.
 */
MarkerFormat::~MarkerFormat()
{
}

std::optional<MarkerFormat::MarkerType> MarkerFormat::type() const
{
    if (d) return d->markerType;
    return {};
}

void MarkerFormat::setType(MarkerFormat::MarkerType type)
{
    if (!d) d = new MarkerFormatPrivate;
    d->markerType = type;
}

void MarkerFormat::write(QXmlStreamWriter &writer) const
{
    if (!d) return;

    writer.writeStartElement("c:marker");

    if (d->markerType.has_value()) {
        writer.writeEmptyElement("c:symbol");
        QString s; toString(d->markerType.value(), s);
        writer.writeAttribute("val", s);
    }
    writeEmptyElement(writer, QLatin1String("c:size"), d->size);
    if (d->shape.isValid()) d->shape.write(writer, "c:spPr");
    writer.writeEndElement();
}

void MarkerFormat::read(QXmlStreamReader &reader)
{
    if (!d) d = new MarkerFormatPrivate;
    const auto &name = reader.name();
    while (!reader.atEnd()) {
        auto token = reader.readNext();
        if (token == QXmlStreamReader::StartElement) {
            if (reader.name() == QLatin1String("symbol")) {
                MarkerType t;
                fromString(reader.attributes().value(QLatin1String("val")).toString(), t);
                d->markerType = t;
            }
            else if (reader.name() == QLatin1String("size"))
                d->size = reader.attributes().value(QLatin1String("val")).toInt();
            else if (reader.name() == QLatin1String("spPr"))
                d->shape.read(reader);
            else reader.skipCurrentElement();
        }
        else if (token == QXmlStreamReader::EndElement && reader.name() == name)
            break;
    }
}

ShapeFormat MarkerFormat::shape() const
{
    if (d) return d->shape;
    return {};
}

ShapeFormat &MarkerFormat::shape()
{
    return d->shape;
}

void MarkerFormat::setShape(ShapeFormat shape)
{
    if (!d) d = new MarkerFormatPrivate;
    d->shape = shape;
}

std::optional<int> MarkerFormat::size() const
{
    if (d) return d->size;
    return {};
}

void MarkerFormat::setSize(int size)
{
    if (!d) d = new MarkerFormatPrivate;
    d->size = size;
}

/*!
    Returns true if the format is valid; otherwise returns false.
 */
bool MarkerFormat::isValid() const
{
    if (d)
        return true;
    return false;
}

bool MarkerFormat::operator==(const MarkerFormat &other) const
{
    if (d == other.d) return true;
    if (!d || !other.d) return false;
    return (*this->d.constData() == *other.d.constData());
}
bool MarkerFormat::operator!=(const MarkerFormat &other) const
{
    return !operator==(other);
}

#ifndef QT_NO_DEBUG_STREAM
QDebug operator<<(QDebug dbg, const MarkerFormat &f)
{
    dbg.nospace() << "QXlsx::MarkerFormat(" << f.d->markerType.value() << ")";
    return dbg.space();
}
#endif

}
