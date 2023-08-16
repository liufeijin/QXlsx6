#include "xlsxlegend.h"

QT_BEGIN_NAMESPACE_XLSX

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
    Text text;
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
    d = other.d;
    return *this;
}

bool Legend::isValid() const
{
    if (d) return true;
    return false;
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

Text Legend::text() const
{
    if (d) return d->text;
    return {};
}

Text &Legend::text()
{
    if (!d) d = new LegendPrivate;
    return d->text;
}

void Legend::setText(const Text &text)
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
        d->entries << LegendEntry(index, visible);
    }
}

std::optional<bool> Legend::entryVisible(int index) const
{
    if (!d) return {};

    const auto e = std::find_if(d->entries.constBegin(), d->entries.constEnd(), [index](const LegendEntry &e){
             return e.idx == index;
    });
    if (e != d->entries.end()) return e->visible;

    return {};
}

void Legend::addEntry(int index, const Text &text, bool visible)
{
    if (!d) d = new LegendPrivate;

    auto e = std::find_if(d->entries.begin(), d->entries.end(), [index](LegendEntry &e){
             return e.idx == index;
    });
    if (e != d->entries.end()) {
        e->visible = visible;
        e->entry = text;
    }
    else d->entries << LegendEntry(index, visible, text);
}

LegendEntry *Legend::entry(int index)
{
    if (!d) d = new LegendPrivate;

    auto e = std::find_if(d->entries.begin(), d->entries.end(), [index](const LegendEntry &e) {
        return e.idx == index;
    });
    if (e != d->entries.end()) return &*e;
    return nullptr;
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
                d->text.read(reader, false);
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
    for (const auto &e: d->entries) e.write(writer, QLatin1String("c:legendEntry"));
    d->layout.write(writer, QLatin1String("c:layout"));
    writeEmptyElement(writer, QLatin1String("c:overlay"), d->overlay);
    if (d->shape.isValid()) d->shape.write(writer, QLatin1String("c:spPr"));
    if (d->text.isValid()) d->text.write(writer, QLatin1String("c:txPr"), false);
    writer.writeEndElement();
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

LegendEntry::LegendEntry()
{

}

LegendEntry::LegendEntry(int index, bool visible)
    : idx{index}, visible{visible}
{

}

LegendEntry::LegendEntry(int index, bool visible, const Text &text)
    : idx{index}, visible{visible}, entry{text}
{

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
    if (entry.isValid()) entry.write(writer, QLatin1String("c:txPr"), false);
    writer.writeEndElement();
}

QT_END_NAMESPACE_XLSX
