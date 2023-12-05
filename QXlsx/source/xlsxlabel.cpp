#include "xlsxlabel.h"

namespace QXlsx {

class SharedLabelProperties
{
public:
    Label::ShowParameters showFlags{};
    std::optional<Label::Position> pos{std::nullopt};
    QString numberFormat;
    bool formatSourceLinked{false};
    ShapeFormat shape;
    TextFormat text;
    QString separator;

    void read(QXmlStreamReader &reader);
    void write(QXmlStreamWriter &writer) const;

    bool operator==(const SharedLabelProperties &other) const;
    bool operator!=(const SharedLabelProperties &other) const;
};

QDebug operator<<(QDebug dbg, const SharedLabelProperties &f);

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
    ExtensionList extList;
};

LabelPrivate::LabelPrivate()
{

}

LabelPrivate::LabelPrivate(const LabelPrivate &other) : QSharedData(other),
    index{other.index}, visible{other.visible}, properties{other.properties},
    layout{other.layout}, text{other.text}, extList{other.extList}
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
    if (extList != other.extList) return false;
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
    if (*this != other)
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

Layout Label::layout() const
{
    if (d) return d->layout;
    return {};
}

Layout &Label::layout()
{
    if (!d) d = new LabelPrivate;
    return d->layout;
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

Label::ShowParameters Label::showParameters() const
{
    if (d) return d->properties.showFlags;
    return {};
}

void Label::setShowParameters(ShowParameters showParameters)
{
    if (!d) d = new LabelPrivate;
    d->properties.showFlags = showParameters;
}

void Label::setShowCategory(bool show)
{
    if (!d) d = new LabelPrivate;
    d->properties.showFlags.setFlag(ShowCategory, show);
}

void Label::setShowLegendKey(bool show)
{
    if (!d) d = new LabelPrivate;
    d->properties.showFlags.setFlag(ShowLegendKey, show);
}
void Label::setShowValue(bool show)
{
    if (!d) d = new LabelPrivate;
    d->properties.showFlags.setFlag(ShowValue, show);
}
void Label::setShowSeries(bool show)
{
    if (!d) d = new LabelPrivate;
    d->properties.showFlags.setFlag(ShowSeries, show);
}
void Label::setShowPercent(bool show)
{
    if (!d) d = new LabelPrivate;
    d->properties.showFlags.setFlag(ShowPercent, show);
}
void Label::setShowBubbleSize(bool show)
{
    if (!d) d = new LabelPrivate;
    d->properties.showFlags.setFlag(ShowBubbleSize, show);
}
void Label::setShowParameter(ShowParameter parameter, bool value)
{
    if (!d) d = new LabelPrivate;
    d->properties.showFlags.setFlag(parameter, value);
}

bool Label::testShowParameter(ShowParameter parameter) const
{
    if (d) return d->properties.showFlags.testFlag(parameter);
    return false;
}

bool Label::testShowCategory() const
{
    if (d) return d->properties.showFlags.testFlag(ShowParameter::ShowCategory);
    return false;
}
bool Label::testShowLegendKey() const
{
    if (d) return d->properties.showFlags.testFlag(ShowParameter::ShowLegendKey);
    return false;
}
bool Label::testShowValue() const
{
    if (d) return d->properties.showFlags.testFlag(ShowParameter::ShowValue);
    return false;
}
bool Label::testShowSeries() const
{
    if (d) return d->properties.showFlags.testFlag(ShowParameter::ShowSeries);
    return false;
}
bool Label::testShowPercent() const
{
    if (d) return d->properties.showFlags.testFlag(ShowParameter::ShowPercent);
    return false;
}
bool Label::testShowBubbleSize() const
{
    if (d) return d->properties.showFlags.testFlag(ShowParameter::ShowBubbleSize);
    return false;
}

void Label::setPosition(Label::Position position)
{
    if (!d) d = new LabelPrivate;
    d->properties.pos = position;
}

std::optional<Label::Position> Label::position() const
{
    if (d) return d->properties.pos;
    return {};
}

QString Label::numberFormat() const
{
    if (d) return d->properties.numberFormat;
    return {};
}
void Label::setNumberFormat(const QString &numberFormat, bool formatSourceLinked)
{
    if (!d) d = new LabelPrivate;
    d->properties.numberFormat = numberFormat;
    d->properties.formatSourceLinked = formatSourceLinked;
}
ShapeFormat Label::shape() const
{
    if (d) return d->properties.shape;
    return {};
}
ShapeFormat &Label::shape()
{
    if (!d) d = new LabelPrivate;
    return d->properties.shape;
}
void Label::setShape(const ShapeFormat &shape)
{
    if (!d) d = new LabelPrivate;
    d->properties.shape = shape;
}
TextFormat Label::textFormat() const
{
    if (d) return d->properties.text;
    return {};
}
TextFormat &Label::textFormat()
{
    if (!d) d = new LabelPrivate;
    return d->properties.text;
}
void Label::setTextFormat(const TextFormat &textFormat)
{
    if (!d) d = new LabelPrivate;
    d->properties.text = textFormat;
}
QString Label::separator() const
{
    if (d) return d->properties.separator;
    return {};
}
void Label::setSeparator(const QString &separator)
{
    if (!d) d = new LabelPrivate;
    d->properties.separator = separator;
}

void Label::clearProperties()
{
    if (d) d->properties = SharedLabelProperties{};
}

void Label::read(QXmlStreamReader &reader)
{
    if (!d) d = new LabelPrivate;
    const auto &name = reader.name();

    while (!reader.atEnd()) {
        auto token = reader.readNext();
        if (token == QXmlStreamReader::StartElement) {
            const auto &a = reader.attributes();
            if (reader.name() == QLatin1String("idx")) {
                parseAttributeInt(a, QLatin1String("val"), d->index);
            }
            else if (reader.name() == QLatin1String("delete")) {
                parseAttributeBool(a, QLatin1String("val"), d->visible);
            }
            else if (reader.name() == QLatin1String("layout")) {
                d->layout.read(reader);
            }
            else if (reader.name() == QLatin1String("tx")) {
                d->text.read(reader);
            }
            else if (reader.name() == QLatin1String("extLst")) {
                d->extList.read(reader);
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
    writeEmptyElement(writer, QLatin1String("c:idx"), d->index);

    if (!d->visible.value_or(false)) {
        writeEmptyElement(writer, QLatin1String("c:delete"), true);
    }
    else {
        d->layout.write(writer, QLatin1String("c:layout"));
        if (d->text.isValid()) d->text.write(writer, QLatin1String("c:tx"));
        d->properties.write(writer);
    }
    d->extList.write(writer, QLatin1String("c:extLst"));

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

class LabelsPrivate : public QSharedData
{
public:
    LabelsPrivate();
    LabelsPrivate(const LabelsPrivate &other);
    ~LabelsPrivate();

    bool operator ==(const LabelsPrivate &other) const;

    QList<Label> labels;
    std::optional<bool> visible;
    std::optional<bool> showLeaderLines;
    ShapeFormat leaderLines;
    SharedLabelProperties defaultProperties;
    ExtensionList extLst;
};

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
    if (*this != other)
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

bool Labels::addLabel(const Label &label)
{
    if (!d) d = new LabelsPrivate;
    if (auto it = std::find_if(d->labels.cbegin(), d->labels.cend(),
                               [label](const auto &l){return label.index() == l.index();});
        it == d->labels.cend()) {
        d->labels << label;
        return true;
    }
    return false;
}

bool Labels::addLabel(int index, Label::ShowParameters showFlags, Label::Position position)
{
    if (!d) d = new LabelsPrivate;
    if (auto it = std::find_if(d->labels.cbegin(), d->labels.cend(),
                               [index](const auto &l){return index == l.index();});
        it == d->labels.cend()) {
        d->labels << Label(index, showFlags, position);
        return true;
    }
    return false;
}

bool Labels::removeLabel(int index)
{
    if (d && index >= 0 && index < d->labels.size()) {
        d->labels.removeAt(index);
        return true;
    }
    return false;
}

bool Labels::removeLabelForPoint(int index)
{
    if (!d) return false;
    int idx = -1;
    for (int i=0; i<d->labels.size(); ++i) {
        if (d->labels.at(i).index() == index) {
            idx = i;
            break;
        }
    }
    if (idx >= 0) {
        d->labels.removeAt(index);
        return true;
    }
    return false;
}

void Labels::removeLabels()
{
    if (d) d->labels.clear();
}

Label& Labels::label(int index)
{
    if (!d) d = new LabelsPrivate;
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
    if (index < 0) return std::nullopt;

    int idx = -1;
    for (int i=0; i<d->labels.size(); ++i) {
        if (d->labels.at(i).index() == index) {
            idx = i;
            break;
        }
    }
    if (idx >= 0) return d->labels[idx];
    d->labels << Label(index, showParameters(), position().value_or(Label::Position::BestFit));
    return d->labels.last();
}

Label Labels::labelForPoint(int index) const
{
    if (!d) return {};
    for (auto &l: qAsConst(d->labels)) {
        if (l.index() == index) return l;
    }
    return {};
}

int Labels::labelsCount() const
{
    if (d) return d->labels.size();
    return 0;
}

bool Labels::hasLabelForPoint(int index) const
{
    if (!d) return false;
    return std::find_if(d->labels.cbegin(), d->labels.cend(), [index](const Label &l)
                        {return l.index() == index;}) != d->labels.cend();
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

Label::ShowParameters Labels::showParameters() const
{
    if (d) return d->defaultProperties.showFlags;
    return {};
}

void Labels::setShowParameters(Label::ShowParameters showParameters)
{
    if (!d) d = new LabelsPrivate;
    d->defaultProperties.showFlags = showParameters;
}

void Labels::setShowCategory(bool show)
{
    if (!d) d = new LabelsPrivate;
    d->defaultProperties.showFlags.setFlag(Label::ShowCategory, show);
}

void Labels::setShowLegendKey(bool show)
{
    if (!d) d = new LabelsPrivate;
    d->defaultProperties.showFlags.setFlag(Label::ShowLegendKey, show);
}
void Labels::setShowValue(bool show)
{
    if (!d) d = new LabelsPrivate;
    d->defaultProperties.showFlags.setFlag(Label::ShowValue, show);
}
void Labels::setShowSeries(bool show)
{
    if (!d) d = new LabelsPrivate;
    d->defaultProperties.showFlags.setFlag(Label::ShowSeries, show);
}
void Labels::setShowPercent(bool show)
{
    if (!d) d = new LabelsPrivate;
    d->defaultProperties.showFlags.setFlag(Label::ShowPercent, show);
}
void Labels::setShowBubbleSize(bool show)
{
    if (!d) d = new LabelsPrivate;
    d->defaultProperties.showFlags.setFlag(Label::ShowBubbleSize, show);
}
void Labels::setShowParameter(Label::ShowParameter parameter, bool value)
{
    if (!d) d = new LabelsPrivate;
    d->defaultProperties.showFlags.setFlag(parameter, value);
}

bool Labels::testShowParameter(Label::ShowParameter parameter) const
{
    if (d) return d->defaultProperties.showFlags.testFlag(parameter);
    return false;
}

bool Labels::testShowCategory() const
{
    if (d) return d->defaultProperties.showFlags.testFlag(Label::ShowParameter::ShowCategory);
    return false;
}
bool Labels::testShowLegendKey() const
{
    if (d) return d->defaultProperties.showFlags.testFlag(Label::ShowParameter::ShowLegendKey);
    return false;
}
bool Labels::testShowValue() const
{
    if (d) return d->defaultProperties.showFlags.testFlag(Label::ShowParameter::ShowValue);
    return false;
}
bool Labels::testShowSeries() const
{
    if (d) return d->defaultProperties.showFlags.testFlag(Label::ShowParameter::ShowSeries);
    return false;
}
bool Labels::testShowPercent() const
{
    if (d) return d->defaultProperties.showFlags.testFlag(Label::ShowParameter::ShowPercent);
    return false;
}
bool Labels::testShowBubbleSize() const
{
    if (d) return d->defaultProperties.showFlags.testFlag(Label::ShowParameter::ShowBubbleSize);
    return false;
}

void Labels::setPosition(Label::Position position)
{
    if (!d) d = new LabelsPrivate;
    d->defaultProperties.pos = position;
}

std::optional<Label::Position> Labels::position() const
{
    if (d) return d->defaultProperties.pos;
    return {};
}

QString Labels::numberFormat() const
{
    if (d) return d->defaultProperties.numberFormat;
    return {};
}
void Labels::setNumberFormat(const QString &numberFormat, bool formatSourceLinked)
{
    if (!d) d = new LabelsPrivate;
    d->defaultProperties.numberFormat = numberFormat;
    d->defaultProperties.formatSourceLinked = formatSourceLinked;
}
ShapeFormat Labels::shape() const
{
    if (d) return d->defaultProperties.shape;
    return {};
}
ShapeFormat &Labels::shape()
{
    if (!d) d = new LabelsPrivate;
    return d->defaultProperties.shape;
}
void Labels::setShape(const ShapeFormat &shape)
{
    if (!d) d = new LabelsPrivate;
    d->defaultProperties.shape = shape;
}
TextFormat Labels::textFormat() const
{
    if (d) return d->defaultProperties.text;
    return {};
}
TextFormat &Labels::textFormat()
{
    if (!d) d = new LabelsPrivate;
    return d->defaultProperties.text;
}
void Labels::setTextFormat(const TextFormat &textFormat)
{
    if (!d) d = new LabelsPrivate;
    d->defaultProperties.text = textFormat;
}
QString Labels::separator() const
{
    if (d) return d->defaultProperties.separator;
    return {};
}
void Labels::setSeparator(const QString &separator)
{
    if (!d) d = new LabelsPrivate;
    d->defaultProperties.separator = separator;
}

void Labels::clearDefaultProperties()
{
    if (d) d->defaultProperties = {};
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
            else if (reader.name() == QLatin1String("extLst")) {
                d->extLst.read(reader);
            }
            else d->defaultProperties.read(reader);
        }
        else if (token == QXmlStreamReader::EndElement && reader.name() == name)
            break;
    }
}

void Labels::write(QXmlStreamWriter &writer) const
{
    if (!d) return;

    writer.writeStartElement(QLatin1String("c:dLbls"));
    for (const auto &l: qAsConst(d->labels)) l.write(writer);

    if (!d->visible.value_or(false) && d->labels.isEmpty()) {
        writeEmptyElement(writer, QLatin1String("c:delete"), true);
    }
    else {
        d->defaultProperties.write(writer);
        writeEmptyElement(writer, QLatin1String("c:showLeaderLines"), d->showLeaderLines);
        if (d->leaderLines.isValid()) {
            writer.writeStartElement(QLatin1String("c:leaderLines"));
            d->leaderLines.write(writer, QLatin1String("c:spPr"));
            writer.writeEndElement(); //c:leaderLines
        }
    }
    if (d->extLst.isValid()) d->extLst.write(writer, QLatin1String("c:extLst"));

    writer.writeEndElement(); //c:dLbls
}

QList<std::reference_wrapper<FillFormat> > Labels::fills()
{
    QList<std::reference_wrapper<FillFormat> > result;

    if (d) {
        if (d->leaderLines.isValid()) result << d->leaderLines.fills();
        if (d->defaultProperties.shape.isValid()) result << d->defaultProperties.shape.fills();
        for (auto &l: d->labels) {
            if (!l.isValid()) continue;
            if (l.d->properties.shape.isValid()) result << l.d->properties.shape.fills();
        }
    }

    return result;
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
    defaultProperties{other.defaultProperties}, extLst{other.extLst}
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
        text.read(reader);
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
    shape.write(writer, QLatin1String("c:spPr"));
    text.write(writer, QLatin1String("c:txPr"));
    if (pos.has_value())
        writeEmptyElement(writer, QLatin1String("c:dLblPos"), Label::toString(pos.value()));
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
    if (leaderLines != other.leaderLines) return false;
    if (defaultProperties != other.defaultProperties) return false;
    if (extLst != other.extLst) return false;
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

}
