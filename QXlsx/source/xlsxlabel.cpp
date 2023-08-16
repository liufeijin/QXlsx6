#include "xlsxlabel.h"

QT_BEGIN_NAMESPACE_XLSX

class LabelPrivate : public QSharedData
{
public:
    LabelPrivate();
    LabelPrivate(const LabelPrivate &other);
    ~LabelPrivate();

    bool operator ==(const LabelPrivate &other) const;

    int index;
    std::optional<bool> visible;
    SharedLabelProperties properties;
    Layout layout;
    Text text;
};

LabelPrivate::LabelPrivate()
{

}

LabelPrivate::LabelPrivate(const LabelPrivate &other) : QSharedData(other),
    index{other.index}, visible{other.visible}, properties{other.properties},
    layout{other.layout}, text{other.text}
{
}

LabelPrivate::~LabelPrivate()
{

}

bool LabelPrivate::operator ==(const LabelPrivate &other) const
{
    //int index; indexes are unique, so do not compare them
    if (visible != other.visible) return false;
    if (properties != other.properties) return false;
    if (layout != other.layout) return false;
    if (text != other.text) return false;
    return true;
}

Label::Label()
{

}

Label::Label(int index, ShowParameters show, Label::Position position)
{
    d = new LabelPrivate;
    d->index = index;
    d->properties.showFlags = show;
    d->properties.pos = position;
}

Label::Label(const Label &other) : d(other.d)
{

}

Label &Label::operator=(const Label &other)
{
    d = other.d;
    return *this;
}

Label::~Label()
{

}

int Label::index() const
{
    if (d) return d->index;
    return -1;
}

void Label::setIndex(int index)
{
    if (!d) d = new LabelPrivate;
    d->index = index;
}

std::optional<bool> Label::visible() const
{
    if (d) return d->visible;
    return {};
}

void Label::setVisible(bool visible)
{
    if (!d) d = new LabelPrivate;
    d->visible = visible;
}

SharedLabelProperties Label::properties() const
{
    if (d) return d->properties;
    return {};
}

void Label::setProperties(SharedLabelProperties properties)
{
    if (!d) d = new LabelPrivate;
    d->properties = properties;
}

Layout Label::layout() const
{
    if (d) return d->layout;
    return {};
}

void Label::setLayout(const Layout &layout)
{
    if (!d) d = new LabelPrivate;
    d->layout = layout;
}

Text Label::text() const
{
    if (d) return d->text;
    return {};
}

void Label::setText(const Text &text)
{
    if (!d) d = new LabelPrivate;
    d->text = text;
}

void Label::setShowCategory(bool show)
{
    if (!d) d = new LabelPrivate;
    d->properties.showFlags.setFlag(ShowCategory, show);
}

void Label::setPosition(Label::Position position)
{
    if (!d) d = new LabelPrivate;
    d->properties.pos = position;
}

Label::Position Label::position() const
{
    if (d) return d->properties.pos.value_or(Position::BestFit);
    return Position::BestFit;
}

void Label::read(QXmlStreamReader &reader)
{
    if (!d) d = new LabelPrivate;
    const auto &name = reader.name();

    while (!reader.atEnd()) {
        auto token = reader.readNext();
        if (token == QXmlStreamReader::StartElement) {
            if (reader.name() == QLatin1String("idx")) {
                d->index = reader.attributes().value(QLatin1String("val")).toInt();
            }
            else if (reader.name() == QLatin1String("delete")) {
                d->visible = !fromST_Boolean(reader.attributes().value(QLatin1String("val")));
            }
            else if (reader.name() == QLatin1String("layout")) {
                d->layout.read(reader);
            }
            else if (reader.name() == QLatin1String("tx")) {
                d->text.read(reader);
            }
            else d->properties.read(reader);
        }
        else if (token == QXmlStreamReader::EndElement && reader.name() == name)
            break;
    }
}

void Label::write(QXmlStreamWriter &writer) const
{
    if (!d) return;
    writer.writeStartElement("c:dLbl");
    writer.writeEmptyElement("c:idx");
    writer.writeAttribute("val", QString::number(d->index));

    if (d->visible.has_value()) {
        writeEmptyElement(writer, QLatin1String("c:delete"), !d->visible.value());
    }
    else {
        d->layout.write(writer, QLatin1String("c:layout"));
        if (d->text.isValid()) d->text.write(writer, QLatin1String("c:tx"));
        d->properties.write(writer);
    }

    writer.writeEndElement();
}

bool Label::isValid() const
{
    if (d) return true;
    return false;
}

bool Label::operator ==(const Label &other) const
{
    if (d == other.d) return true;
    if (!d || !other.d) return false;
    return *this->d.constData() == *other.d.constData();
}

bool Label::operator !=(const Label &other) const
{
    return !operator==(other);
}

Label::operator QVariant() const
{
    const auto& cref
#if QT_VERSION >= 0x060000 // Qt 6.0 or over
        = QMetaType::fromType<Label>();
#else
        = qMetaTypeId<Label>() ;
#endif
    return QVariant(cref, this);
}

Labels::Labels()
{

}

Labels::Labels(const Labels &other) : d(other.d)
{

}

Labels::Labels(Label::ShowParameters showFlags, Label::Position pos)
{
    d = new LabelsPrivate;
    SharedLabelProperties p;
    p.pos = pos;
    p.showFlags = showFlags;
    d->defaultProperties = p;
}

Labels::~Labels()
{

}

Labels &Labels::operator=(const Labels &other)
{
    d = other.d;
    return *this;
}

bool Labels::operator==(const Labels &other) const
{
    if (d == other.d) return true;
    if (!d || !other.d) return false;
    return *this->d.constData() == *other.d.constData();
}

bool Labels::operator!=(const Labels &other) const
{
    return !(operator==(other));
}

bool Labels::isValid() const
{
    if (d) return true;
    return false;
}

std::optional<bool> Labels::visible() const
{
    if (d) return d->visible;
    return {};
}

void Labels::addLabel(const Label &label)
{
    if (!d) d = new LabelsPrivate;
    d->labels << label;
}

void Labels::addLabel(int index, Label::ShowParameters showFlags, Label::Position position)
{
    if (!d) d = new LabelsPrivate;
    d->labels << Label(index, showFlags, position);
}

std::optional<std::reference_wrapper<Label> > Labels::label(int index)
{
    if (!d) d = new LabelsPrivate;
    if (index < 0 || index >= d->labels.size()) return std::nullopt;
    return d->labels[index];
}

Label Labels::label(int index) const
{
    if (d && index >= 0 && index < d->labels.size()) return d->labels[index];
    return {};
}

std::optional<std::reference_wrapper<Label> > Labels::labelForPoint(int index)
{
    if (!d) d = new LabelsPrivate;
    int idx=-1;
    for (int i=0; i<d->labels.size(); ++i) {
        if (d->labels.at(i).index() == index) {
            idx = i;
            break;
        }
    }
    if (idx >= 0) return d->labels[idx];
    return std::nullopt;
}

Label Labels::labelForPoint(int index) const
{
    if (!d) return {};
    for (auto &l: d->labels) {
        if (l.index() == index) return l;
    }
    return {};
}

void Labels::setVisible(bool visible)
{
    if (!d) d = new LabelsPrivate;
    d->visible = visible;
}

std::optional<bool> Labels::showLeaderLines() const
{
    if (d) return d->showLeaderLines;
    return {};
}

void Labels::setShowLeaderLines(bool showLeaderLines)
{
    if (!d) d = new LabelsPrivate;
    d->showLeaderLines = showLeaderLines;
}

ShapeFormat Labels::leaderLines() const
{
    if (d) return d->leaderLines;
    return {};
}

ShapeFormat &Labels::leaderLines()
{
    if (!d) d = new LabelsPrivate;
    return d->leaderLines;
}

void Labels::setLeaderLines(const ShapeFormat &format)
{
    if (!d) d = new LabelsPrivate;
    d->leaderLines = format;
}

SharedLabelProperties Labels::defaultProperties() const
{
    if (d) return d->defaultProperties;
    return {};
}

void Labels::setDefaultProperties(SharedLabelProperties defaultProperties)
{
    if (!d) d = new LabelsPrivate;
    d->defaultProperties = defaultProperties;
}

void Labels::read(QXmlStreamReader &reader)
{
    if (!d) d = new LabelsPrivate;
    const auto &name = reader.name();

    while (!reader.atEnd()) {
        auto token = reader.readNext();
        if (token == QXmlStreamReader::StartElement) {
            if (reader.name() == QLatin1String("dLbl")) {
                Label l;
                l.read(reader);
                d->labels << l;
            }
            else if (reader.name() == QLatin1String("delete")) {
                d->visible = !fromST_Boolean(reader.attributes().value(QLatin1String("val")));
            }
            else if (reader.name() == QLatin1String("showLeaderLines")) {
                d->showLeaderLines = fromST_Boolean(reader.attributes().value(QLatin1String("val")));
            }
            else if (reader.name() == QLatin1String("leaderLines")) {
                reader.readNextStartElement();
                d->leaderLines.read(reader);
            }
            else d->defaultProperties.read(reader);
        }
        else if (token == QXmlStreamReader::EndElement && reader.name() == name)
            break;
    }
}

void Labels::write(QXmlStreamWriter &writer) const
{
    writer.writeStartElement(QLatin1String("c:dLbls"));
    for (const auto &l: d->labels) l.write(writer);

    if (d->visible.has_value()) {
        writeEmptyElement(writer, QLatin1String("c:delete"), !d->visible.value());
    }
    else {
        d->defaultProperties.write(writer);
        writeEmptyElement(writer, QLatin1String("c:showLeaderLines"), d->showLeaderLines);
        if (d->leaderLines.isValid()) {
            writer.writeEmptyElement(QLatin1String("c:leaderLines"));
            d->leaderLines.write(writer, QLatin1String("c:spPr"));
            writer.writeEndElement();
        }
    }

    writer.writeEndElement();
}

Labels::operator QVariant() const
{
    const auto& cref
#if QT_VERSION >= 0x060000 // Qt 6.0 or over
        = QMetaType::fromType<Labels>();
#else
        = qMetaTypeId<Labels>() ;
#endif
    return QVariant(cref, this);
}

LabelsPrivate::LabelsPrivate()
{

}

LabelsPrivate::LabelsPrivate(const LabelsPrivate &other) : QSharedData(other),
    labels{other.labels}, visible{other.visible}, showLeaderLines{other.showLeaderLines},
    defaultProperties{other.defaultProperties}
{

}

LabelsPrivate::~LabelsPrivate()
{

}

QDebug operator<<(QDebug dbg, const SharedLabelProperties &f)
{
    QDebugStateSaver state(dbg);

    dbg.nospace() << "numberFormat: " << (f.numberFormat.isEmpty() ? QString("not def") : f.numberFormat) << ", ";
    dbg.nospace() << "sourceLinked: " << (f.formatSourceLinked ? QString("true") : QString("false")) << ", ";
    dbg.nospace() << f.shape << ", ";
    dbg.nospace() << f.text << ", ";
    dbg.nospace() << "showVal: " << toST_Boolean(f.showFlags.testFlag(Label::ShowValue)) << ", ";
    dbg.nospace() << "showCatName: " << toST_Boolean(f.showFlags.testFlag(Label::ShowCategory)) << ", ";
    dbg.nospace() << "showSerName: " << toST_Boolean(f.showFlags.testFlag(Label::ShowSeries)) << ", ";
    dbg.nospace() << "showPercent: " << toST_Boolean(f.showFlags.testFlag(Label::ShowPercent)) << ", ";
    dbg.nospace() << "showBubbleSize: " << toST_Boolean(f.showFlags.testFlag(Label::ShowBubbleSize)) << ", ";
    dbg.nospace() << "showLegendKey: " << toST_Boolean(f.showFlags.testFlag(Label::ShowLegendKey)) << ", ";
    dbg.nospace() << "separator: " << (f.separator.isEmpty() ? QString("not def") : f.separator);

    return dbg;
}

QDebug operator<<(QDebug dbg, const Labels &f)
{
    QDebugStateSaver state(dbg);
    dbg.setAutoInsertSpaces(false);

    dbg << "QXlsx::Labels(";
    if (!f.isValid()) dbg << "empty";
    else {
        dbg << "visible: " << f.visible().value_or(false);
        dbg << "showLeaderLines: " << f.showLeaderLines().value_or(false);
        dbg << f.defaultProperties();
    }
    dbg << ")";

    return dbg;
}

void SharedLabelProperties::read(QXmlStreamReader &reader)
{
    const auto &name = reader.name();
    const auto &a = reader.attributes();
    if (name == QLatin1String("numFmt")) {
        numberFormat = a.value(QLatin1String("formatCode")).toString();
        formatSourceLinked = fromST_Boolean(a.value(QLatin1String("sourceLinked")));
    }
    else if (name == QLatin1String("spPr")) {
        shape.read(reader);
    }
    else if (name == QLatin1String("txPr")) {
        text.read(reader, false);
    }
    else if (name == QLatin1String("showLegendKey")) {
        showFlags.setFlag(Label::ShowLegendKey, fromST_Boolean(a.value(QLatin1String("val"))));
    }
    else if (name == QLatin1String("showVal")) {
        showFlags.setFlag(Label::ShowValue, fromST_Boolean(a.value(QLatin1String("val"))));
    }
    else if (name == QLatin1String("showCatName")) {
        showFlags.setFlag(Label::ShowCategory, fromST_Boolean(a.value(QLatin1String("val"))));
    }
    else if (name == QLatin1String("showSerName")) {
        showFlags.setFlag(Label::ShowSeries, fromST_Boolean(a.value(QLatin1String("val"))));
    }
    else if (name == QLatin1String("showPercent")) {
        showFlags.setFlag(Label::ShowPercent, fromST_Boolean(a.value(QLatin1String("val"))));
    }
    else if (name == QLatin1String("showBubbleSize")) {
        showFlags.setFlag(Label::ShowBubbleSize, fromST_Boolean(a.value(QLatin1String("val"))));
    }
    else if (name == QLatin1String("separator")) {
        separator = reader.readElementText();
    }
    else if (name == QLatin1String("dLblPos")) {
        Label::Position p; Label::fromString(a.value(QLatin1String("val")).toString(), p);
        pos = p;
    }
    else reader.skipCurrentElement();
}

void SharedLabelProperties::write(QXmlStreamWriter &writer) const
{
    if (!numberFormat.isEmpty()) {
        writer.writeEmptyElement(QLatin1String("c:numFmt"));
        writer.writeAttribute(QLatin1String("formatCode"), numberFormat);
        writer.writeAttribute(QLatin1String("sourceLinked"), toST_Boolean(formatSourceLinked));
    }
    if (shape.isValid()) shape.write(writer, QLatin1String("c:spPr"));
    if (text.isValid()) text.write(writer, QLatin1String("c:txPr"), false);
    if (pos.has_value()) {
        QString s; Label::toString(pos.value(), s);
        writeEmptyElement(writer, QLatin1String("c:dLblPos"), s);
    }
    writeEmptyElement(writer, QLatin1String("c:showVal"), showFlags.testFlag(Label::ShowValue));
    writeEmptyElement(writer, QLatin1String("c:showCatName"), showFlags.testFlag(Label::ShowCategory));
    writeEmptyElement(writer, QLatin1String("c:showSerName"), showFlags.testFlag(Label::ShowSeries));
    writeEmptyElement(writer, QLatin1String("c:showPercent"), showFlags.testFlag(Label::ShowPercent));
    writeEmptyElement(writer, QLatin1String("c:showBubbleSize"), showFlags.testFlag(Label::ShowBubbleSize));
    writeEmptyElement(writer, QLatin1String("c:showLegendKey"), showFlags.testFlag(Label::ShowLegendKey));

    if (!separator.isEmpty()) writer.writeTextElement(QLatin1String("c:separator"), separator);
}

bool LabelsPrivate::operator ==(const LabelsPrivate &other) const
{
    if (labels != other.labels) return false;
    if (visible != other.visible) return false;
    if (showLeaderLines != other.showLeaderLines) return false;
    // TODO: <xsd:element name="leaderLines" type="CT_ChartLines" minOccurs="0" maxOccurs="1"/>
    if (defaultProperties != other.defaultProperties) return false;
    return true;
}

bool SharedLabelProperties::operator==(const SharedLabelProperties &other) const
{
    if (showFlags != other.showFlags) return false;
    if (pos != other.pos) return false;
    if (numberFormat != other.numberFormat) return false;
    if (formatSourceLinked != other.formatSourceLinked) return false;
    if (shape != other.shape) return false;
    if (text != other.text) return false;
    if (separator != other.separator) return false;
    return true;
}

bool SharedLabelProperties::operator!=(const SharedLabelProperties &other) const
{
    return !operator==(other);
}

QDebug operator<<(QDebug dbg, const Label &f)
{
    QDebugStateSaver saver(dbg);
    dbg.setAutoInsertSpaces(false);

    dbg << "QXlsx::Label(";
    if (!f.isValid())
        dbg << "empty";
    else {
        dbg << "index: " << f.d->index << ", ";
        dbg << "visible: " << (f.d->visible.has_value() ? "true" : "false") << ", ";
        dbg << "properties: " << f.d->properties << ", ";
        dbg << "layout: " << f.d->layout << ", ";
        dbg << "text: " << f.d->text;
    }
    dbg << ")";

    return dbg;
}

QT_END_NAMESPACE_XLSX
