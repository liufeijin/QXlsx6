#include "xlsxlegend.h"

namespace QXlsx {

class LegendPrivate : public QSharedData
{
public:
    LegendPrivate();
    LegendPrivate(const LegendPrivate &other);
    ~LegendPrivate();

    std::optional<Legend::Position> pos;
    QList<LegendEntry> entries;
    Layout layout;
    std::optional<bool> overlay;
    ShapeFormat shape;
    TextFormat text;

    bool operator==(const LegendPrivate &other) const;
};

Legend::Legend()
{

}

Legend::Legend(Legend::Position position) : d{new LegendPrivate}
{
    d->pos = position;
}

Legend::Legend(const Legend &other) : d{other.d}
{

}

Legend::~Legend()
{

}

Legend &Legend::operator=(const Legend &other)
{
    if (*this != other)
        d = other.d;
    return *this;
}

bool Legend::isValid() const
{
    if (d) return true;
    return false;
}

Legend Legend::defaultLegend()
{
    Legend l;
    l.d = new LegendPrivate;
    return l;
}

std::optional<Legend::Position> Legend::position() const
{
    if (d) return d->pos;
    return {};
}

void Legend::setPosition(Legend::Position position)
{
    if (!d) d = new LegendPrivate;
    d->pos = position;
}

std::optional<bool> Legend::overlay() const
{
    if (d) return d->overlay;
    return {};
}

void Legend::setOverlay(bool overlay)
{
    if (!d) d = new LegendPrivate;
    d->overlay = overlay;
}

Layout Legend::layout() const
{
    if (d) return d->layout;
    return {};
}

Layout &Legend::layout()
{
    if (!d) d = new LegendPrivate;
    return d->layout;
}

void Legend::setLayout(const Layout &layout)
{
    if (!d) d = new LegendPrivate;
    d->layout = layout;
}

ShapeFormat Legend::shape() const
{
    if (d) return d->shape;
    return {};
}

ShapeFormat &Legend::shape()
{
    if (!d) d = new LegendPrivate;
    return d->shape;
}

void Legend::setShape(const ShapeFormat &shape)
{
    if (!d) d = new LegendPrivate;
    d->shape = shape;
}

TextFormat Legend::text() const
{
    if (d) return d->text;
    return {};
}

TextFormat &Legend::text()
{
    if (!d) d = new LegendPrivate;
    return d->text;
}

void Legend::setText(const TextFormat &text)
{
    if (!d) d = new LegendPrivate;
    d->text = text;
}

void Legend::setEntryVisible(int index, bool visible)
{
    if (!d) d = new LegendPrivate;

    auto e = std::find_if(d->entries.begin(), d->entries.end(), [index](LegendEntry &e){
             return e.idx == index;
    });
    if (e != d->entries.end()) e->visible = visible;
    else {
        d->entries << LegendEntry(index);
        d->entries.last().visible = visible;
    }
}

std::optional<bool> Legend::entryVisible(int index) const
{
    if (!d) return std::nullopt;

    const auto e = std::find_if(d->entries.constBegin(), d->entries.constEnd(), [index](const LegendEntry &e){
             return e.idx == index;
    });
    if (e != d->entries.end()) return e->visible;

    return {};
}

LegendEntry &Legend::entry(int index)
{
    if (!d) d = new LegendPrivate;

    auto e = std::find_if(d->entries.begin(), d->entries.end(), [index](const LegendEntry &e) {
        return e.idx == index;
    });
    if (e != d->entries.end()) return *e;

    d->entries << LegendEntry(index);
    return d->entries.last();
}

void Legend::read(QXmlStreamReader &reader)
{
    if (!d) d = new LegendPrivate;

    const auto &name = reader.name();

    while (!reader.atEnd()) {
        auto token = reader.readNext();
        if (token == QXmlStreamReader::StartElement) {
            const auto &a = reader.attributes();
            if (reader.name() == QLatin1String("legendPos")) {
                Legend::Position p;
                fromString(a.value(QLatin1String("val")).toString(), p);
                d->pos = p;
            }
            else if (reader.name() == QLatin1String("legendEntry")) {
                LegendEntry e; e.read(reader);
                d->entries << e;
            }
            else if (reader.name() == QLatin1String("layout"))
                d->layout.read(reader);
            else if (reader.name() == QLatin1String("overlay"))
                d->overlay = fromST_Boolean(a.value(QLatin1String("val")));
            else if (reader.name() == QLatin1String("spPr"))
                d->shape.read(reader);
            else if (reader.name() == QLatin1String("txPr"))
                d->text.read(reader);
            else if (reader.name() == QLatin1String("extLst")) {
                //TODO:
                reader.skipCurrentElement();
            }
        }
        else if (token == QXmlStreamReader::EndElement && reader.name() == name)
            break;
    }
}

void Legend::write(QXmlStreamWriter &writer) const
{
    if (!d) return;

    writer.writeStartElement(QLatin1String("c:legend"));
    if (d->pos.has_value()) {
        QString s; toString(d->pos.value(), s);
        writeEmptyElement(writer, QLatin1String("c:legendPos"), s);
    }
    for (const auto &e: qAsConst(d->entries)) e.write(writer, QLatin1String("c:legendEntry"));
    d->layout.write(writer, QLatin1String("c:layout"));
    writeEmptyElement(writer, QLatin1String("c:overlay"), d->overlay);
    if (d->shape.isValid()) d->shape.write(writer, QLatin1String("c:spPr"));
    if (d->text.isValid()) d->text.write(writer, QLatin1String("c:txPr"));
    writer.writeEndElement();
}

bool Legend::operator ==(const Legend &other) const
{
    if (d == other.d) return true;
    if (!d || !other.d) return false;
    return (*this->d.constData() == *other.d.constData());
}
bool Legend::operator !=(const Legend &other) const
{
    return !operator==(other);
}

LegendPrivate::LegendPrivate()
{

}

LegendPrivate::LegendPrivate(const LegendPrivate &other) : QSharedData{other}
  , pos{other.pos}, entries{other.entries}, layout{other.layout},
    overlay{other.overlay}, shape{other.shape}, text{other.text}
{

}

LegendPrivate::~LegendPrivate()
{

}

bool LegendPrivate::operator==(const LegendPrivate &other) const
{
    if (pos != other.pos) return false;
    if (entries != other.entries) return false;
    if (layout != other.layout) return false;
    if (overlay != other.overlay) return false;
    if (shape != other.shape) return false;
    if (text != other.text) return false;
    return true;
}

LegendEntry::LegendEntry()
{

}

LegendEntry::LegendEntry(int index)
    : idx{index}
{

}

LegendEntry::LegendEntry(int index, bool visible, const TextFormat &text)
    : idx{index}, visible{visible}, entry{text}
{

}

bool LegendEntry::isValid() const
{
    if (idx > -1) {
        return visible.has_value() || entry.isValid();
    }
    return false;
}
bool LegendEntry::operator==(const LegendEntry &other) const
{
    //indexes should be different, otherwise LegendEntry is invalid
    return visible == other.visible && entry == other.entry;
}

void LegendEntry::read(QXmlStreamReader &reader)
{
    const auto &name = reader.name();

    while (!reader.atEnd()) {
        auto token = reader.readNext();
        if (token == QXmlStreamReader::StartElement) {
            const auto &a = reader.attributes();
            if (reader.name() == QLatin1String("idx")) {
                parseAttributeInt(a, QLatin1String("val"), idx);
            }
            else if (reader.name() == QLatin1String("delete")) {
                visible = !fromST_Boolean(a.value(QLatin1String("val")));
            }
            else if (reader.name() == QLatin1String("txPr"))
                entry.read(reader);
            else reader.skipCurrentElement();
        }
        else if (token == QXmlStreamReader::EndElement && reader.name() == name)
            break;
    }
}

void LegendEntry::write(QXmlStreamWriter &writer, const QString &name) const
{
    writer.writeStartElement(name);
    writeEmptyElement(writer, QLatin1String("c:idx"), idx);
    if (visible.has_value()) writeEmptyElement(writer, QLatin1String("c:delete"), !visible.value());
    if (entry.isValid()) entry.write(writer, QLatin1String("c:txPr"));
    writer.writeEndElement();
}

}
