#include <QtGlobal>
#include <QDataStream>
#include <QDebug>

#include "xlsxaxis.h"
#include "xlsxutility_p.h"

namespace QXlsx {

class AxisPrivate : public QSharedData
{
public:
    int id = -1;
    Axis::Scaling scaling;
    std::optional<bool> visible;
    Axis::Position position = Axis::Position::None;
    Axis::Type type = Axis::Type::None;

    int crossAxis = -1; //make sure this is synchronized with crossAxisPt
    Axis* crossAxisPt = nullptr; //make sure this is synchronized with crossAxis
    std::optional<Axis::CrossesType> crossesType;
    std::optional<double> crossesPosition;

    bool majorGridlinesOn = false;
    bool minorGridlinesOn = false;
    ShapeFormat majorGridlines;
    ShapeFormat minorGridlines;

    Title title;
    TextFormat textProperties;
    NumberFormat numberFormat;

    std::optional<Axis::TickMark> majorTickMark;
    std::optional<Axis::TickMark> minorTickMark;

    std::optional<Axis::TickLabelPosition> tickLabelPosition;
    ShapeFormat shape;

    //Cat, Date
    std::optional<bool> axAuto;
    std::optional<int> labelOffset;

    //Cat
    std::optional<Axis::LabelAlignment> labelAlignment;
    std::optional<bool> noMultiLevelLabels;

    //Date, Val
    std::optional<double> majorUnit;
    std::optional<double> minorUnit;

    //Date
    std::optional<Axis::TimeUnit> baseTimeUnit;
    std::optional<Axis::TimeUnit> majorTimeUnit;
    std::optional<Axis::TimeUnit> minorTimeUnit;

    //Cat, Ser
    std::optional<int> tickLabelSkip;
    std::optional<int> tickMarkSkip;

    //Val
    std::optional<Axis::CrossesBetweenType> crossesBetween;
    DisplayUnits displayUnits;

    AxisPrivate();
    AxisPrivate(const AxisPrivate &other);
    ~AxisPrivate();

    bool operator == (const AxisPrivate &other) const;
};

void Axis::Scaling::write(QXmlStreamWriter &writer) const
{
    if (!isValid()) {
        writer.writeEmptyElement(QStringLiteral("c:scaling"));
        return;
    }

    writer.writeStartElement(QStringLiteral("c:scaling")); // CT_Scaling (mandatory value)
    if (reversed.has_value()) {
        writer.writeEmptyElement(QStringLiteral("c:orientation"));
        writer.writeAttribute(QStringLiteral("val"), reversed.value() ? QStringLiteral("maxMin") : QStringLiteral("minMax")); // ST_Orientation
    }
    if (logBase.has_value()) {
        writer.writeEmptyElement(QStringLiteral("c:logBase"));
        writer.writeAttribute(QStringLiteral("val"), QString::number(logBase.value()));
    }
    if (min.has_value()) {
        writer.writeEmptyElement(QStringLiteral("c:min"));
        writer.writeAttribute(QStringLiteral("val"), QString::number(min.value()));
    }
    if (max.has_value()) {
        writer.writeEmptyElement(QStringLiteral("c:max"));
        writer.writeAttribute(QStringLiteral("val"), QString::number(max.value()));
    }
    writer.writeEndElement(); // c:scaling
}

void Axis::Scaling::read(QXmlStreamReader &reader)
{
    Q_ASSERT(reader.name() == QStringLiteral("scaling"));

    while (!reader.atEnd()) {
        auto token = reader.readNext();
        if (token == QXmlStreamReader::StartElement) {
            auto val = reader.attributes().value(QLatin1String("val"));
            if (reader.name() == QLatin1String("orientation")) {
                reversed = val.toString() == "maxMin";
            }
            else if (reader.name() == QLatin1String("logBase")) {
                logBase = val.toDouble();
            }
            else if (reader.name() == QLatin1String("min")) {
                min = val.toDouble();
            }
            else if (reader.name() == QLatin1String("max")) {
                max = val.toDouble();
            }
            else if (reader.name() == QLatin1String("extLst")) {
                reader.skipCurrentElement();
            }
        }
        else if (reader.tokenType() == QXmlStreamReader::EndElement &&
                 reader.name() == QLatin1String("scaling")) {
            break;
        }
    }
}

bool Axis::Scaling::isValid() const
{
    return (logBase.has_value() || reversed.has_value() || min.has_value() || max.has_value());
}

bool Axis::Scaling::operator ==(const Axis::Scaling &other) const
{
    return (logBase == other.logBase &&
            reversed == other.reversed &&
            min == other.min &&
            max == other.max);
}

bool Axis::Scaling::operator !=(const Axis::Scaling &other) const
{
    return !operator==(other);
}

Axis::Axis()
{

}

Axis::Axis(Axis::Type type, Axis::Position position)
{
    d = new AxisPrivate;
    d->type = type;
    d->position = position;
}

Axis::Axis(const Axis &other) : d(other.d)
{

}

Axis &Axis::operator=(const Axis &other)
{
    if (*this != other)
        d = other.d;
    return *this;
}

Axis::~Axis()
{

}

bool Axis::isValid() const
{
    if (d) {
        if (d->type != Type::None && d->id > 0) return true;
    }
    return false;
}

Axis::Type Axis::type() const
{
    if (d) return d->type;
    return Axis::Type::None;
}

void Axis::setType(Type type)
{
    if (!d) d = new AxisPrivate;
    d->type = type;
}

Axis::Position Axis::position() const
{
    if (d) return d->position;
    return Axis::Position::None;
}

void Axis::setPosition(Axis::Position position)
{
    if (!d) d = new AxisPrivate;
    d->position = position;
}

std::optional<bool> Axis::visible() const
{
    if (d) return d->visible;
    return {};
}

void Axis::setVisible(bool visible)
{
    if (!d) d = new AxisPrivate;
    d->visible = visible;
}

int Axis::id() const
{
    if (d) return d->id;
    return -1;
}

void Axis::setId(int id)
{
    if (!d) d = new AxisPrivate;
    d->id = id;
}

Axis * Axis::crossAxis() const
{
    if (d) return d->crossAxisPt;
    return nullptr;
}

void Axis::setCrossAxis(Axis *axis)
{
    if (!d) d = new AxisPrivate;

    if (d->crossAxisPt != axis) {
        d->crossAxisPt = axis;

        if (axis) {
            d->crossAxis = axis->id();
            axis->setCrossAxis(this);
        }
    }
}

std::optional<double> Axis::crossesAt() const
{
    if (d) return d->crossesPosition;
    return {};
}

void Axis::setCrossesAt(double val)
{
    if (!d) d = new AxisPrivate;
    d->crossesType = CrossesType::Position;
    d->crossesPosition = val;
}

std::optional<Axis::CrossesType> Axis::crossesType() const
{
    if (d) return d->crossesType;
    return {};
}

void Axis::setCrossesType(Axis::CrossesType val)
{
    if (!d) d = new AxisPrivate;
    d->crossesType = val;
}

ShapeFormat &Axis::majorGridLines()
{
    if (!d) d = new AxisPrivate;
    return d->majorGridlines;
}

ShapeFormat Axis::majorGridLines() const
{
    if (d) return d->majorGridlines;
    return {};
}

void Axis::setMajorGridLines(const ShapeFormat &val)
{
    if (!d) d = new AxisPrivate;
    d->majorGridlines = val;
}

ShapeFormat &Axis::minorGridLines()
{
    if (!d) d = new AxisPrivate;
    return d->minorGridlines;
}

ShapeFormat Axis::minorGridLines() const
{
    if (d) return d->minorGridlines;
    return {};
}

void Axis::setMinorGridLines(const ShapeFormat &val)
{
    if (!d) d = new AxisPrivate;
    d->minorGridlines = val;
}

void Axis::setMajorGridLines(const QColor &color, double width, LineFormat::StrokeType strokeType)
{
    if (!d) d = new AxisPrivate;
    LineFormat lf(FillFormat::FillType::SolidFill, width, color);
    lf.setStrokeType(strokeType);
    d->majorGridlines.setLine(lf);
}

void Axis::setMinorGridLines(const QColor &color, double width, LineFormat::StrokeType strokeType)
{
    if (!d) d = new AxisPrivate;
    LineFormat lf(FillFormat::FillType::SolidFill, width, color);
    lf.setStrokeType(strokeType);
    d->minorGridlines.setLine(lf);
}

void Axis::setMajorGridLines(bool on)
{
    if (!d) d = new AxisPrivate;
    d->majorGridlinesOn = on;
}

void Axis::setMinorGridLines(bool on)
{
    if (!d) d = new AxisPrivate;
    d->minorGridlinesOn = on;
}

void Axis::setMajorTickMark(Axis::TickMark tickMark)
{
    if (!d) d = new AxisPrivate;
    d->majorTickMark = tickMark;
}

void Axis::setMinorTickMark(Axis::TickMark tickMark)
{
    if (!d) d = new AxisPrivate;
    d->minorTickMark = tickMark;
}

std::optional<Axis::TickMark> Axis::majorTickMark() const
{
    if (d)
        return d->majorTickMark;
    return {};
}

std::optional<Axis::TickMark> Axis::minorTickMark() const
{
    if (d)
        return d->minorTickMark;
    return {};
}

QString Axis::titleAsString() const
{
    if (d) return d->title.toPlainString();
    return QString();
}

Title Axis::title() const
{
    if (d) return d->title;
    return Title();
}

Title &Axis::title()
{
    if (!d) d = new AxisPrivate;
    return d->title;
}

void Axis::setTitle(const QString &title)
{
    if (!d) d = new AxisPrivate;
    d->title.setPlainString(title);
}

void Axis::setTitle(const Title &title)
{
    if (!d) d = new AxisPrivate;
    d->title = title;
}

QString Axis::numberFormat() const
{
    if (d) return d->numberFormat.format;
    return {};
}

void Axis::setNumberFormat(const QString &formatCode)
{
    if (!d) d = new AxisPrivate;
    d->numberFormat.format = formatCode;
}

Axis::Scaling &Axis::scaling()
{
    if (!d) d = new AxisPrivate;
    return d->scaling;
}

Axis::Scaling Axis::scaling() const
{
    if (d) return d->scaling;
    return {};
}

void Axis::setScaling(Axis::Scaling scaling)
{
    if (!d) d = new AxisPrivate;
    d->scaling = scaling;
}

QPair<double, double> Axis::range() const
{
    if (d) return {d->scaling.min.value_or(0.0), d->scaling.max.value_or(0.0)};
    return {0.0, 0.0};
}

void Axis::setRange(double min, double max)
{
    if (!d) d = new AxisPrivate;
    d->scaling.min = min;
    d->scaling.max = max;
}

std::optional<bool> Axis::autoAxis() const
{
    if (d) return d->axAuto;
    return {};
}

void Axis::setAutoAxis(bool autoAxis)
{
    if (!d) d = new AxisPrivate;
    d->axAuto = autoAxis;
}

std::optional<Axis::LabelAlignment> Axis::labelAlignment() const
{
    if (d) return d->labelAlignment;
    return {};
}

void Axis::setLabelAlignment(Axis::LabelAlignment labelAlignment)
{
    if (!d) d = new AxisPrivate;
    d->labelAlignment = labelAlignment;
}

std::optional<int> Axis::labelOffset() const
{
    if (d) return d->labelOffset;
    return {};
}

void Axis::setLabelOffset(int labelOffset)
{
    if (!d) d = new AxisPrivate;
    d->labelOffset = labelOffset;
}

std::optional<double> Axis::majorTickDistance() const
{
    if (d) return d->majorUnit;
    return {};
}

void Axis::setMajorTickDistance(double distance)
{
    if (!d) d = new AxisPrivate;
    d->majorUnit = distance;
}

std::optional<double> Axis::minorTickDistance() const
{
    if (d) return d->minorUnit;
    return {};
}

void Axis::setMinorTickDistance(double distance)
{
    if (!d) d = new AxisPrivate;
    d->minorUnit = distance;
}

std::optional<Axis::TimeUnit> Axis::baseTimeUnit() const
{
    if (d) return d->baseTimeUnit;
    return {};
}

void Axis::setBaseTimeUnit(Axis::TimeUnit baseTimeUnit)
{
    if (!d) d = new AxisPrivate;
    d->baseTimeUnit = baseTimeUnit;
}

std::optional<Axis::TimeUnit> Axis::majorTimeUnit() const
{
    if (d) return d->majorTimeUnit;
    return {};
}

void Axis::setMajorTimeUnit(Axis::TimeUnit majorTimeUnit)
{
    if (!d) d = new AxisPrivate;
    d->majorTimeUnit = majorTimeUnit;
}

std::optional<Axis::TimeUnit> Axis::minorTimeUnit() const
{
    if (d) return d->minorTimeUnit;
    return {};
}

void Axis::setMinorTimeUnit(Axis::TimeUnit minorTimeUnit)
{
    if (!d) d = new AxisPrivate;
    d->minorTimeUnit = minorTimeUnit;
}

std::optional<int> Axis::tickLabelSkip()
{
    if (d) return d->tickLabelSkip;
    return {};
}

void Axis::setTickLabelSkip(int skip)
{
    if (!d) d = new AxisPrivate;
    d->tickLabelSkip = skip;
}

std::optional<int> Axis::tickMarkSkip()
{
    if (d) return d->tickMarkSkip;
    return {};
}

void Axis::setTickMarkSkip(int skip)
{
    if (!d) d = new AxisPrivate;
    d->tickMarkSkip = skip;
}

std::optional<bool> Axis::noMultiLevelLabels() const
{
    if (d) return d->noMultiLevelLabels;
    return {};
}

void Axis::setNoMultiLevelLabels(bool val)
{
    if (!d) d = new AxisPrivate;
    d->noMultiLevelLabels = val;
}

std::optional<Axis::CrossesBetweenType> Axis::crossesBetween() const
{
    if (d) return d->crossesBetween;
    return {};
}

void Axis::setCrossesBetween(Axis::CrossesBetweenType crossesBetween)
{
    if (!d) d = new AxisPrivate;
    d->crossesBetween = crossesBetween;
}

void Axis::write(QXmlStreamWriter &writer) const
{
    QString name;
    switch (d->type) {
        case Type::Category: name = "c:catAx"; break;
        case Type::Series: name = "c:serAx"; break;
        case Type::Value: name = "c:valAx"; break;
        case Type::Date: name = "c:dateAx"; break;
        default: break;
    }
    if (name.isEmpty()) return;

    writer.writeStartElement(name);
    writeEmptyElement(writer, QLatin1String("c:axId"), d->id);
    d->scaling.write(writer);

    if (d->visible.has_value())
        writeEmptyElement(writer, QLatin1String("c:delete"), !d->visible.value());

    writer.writeEmptyElement("c:axPos");
    QString s; toString(d->position, s); writer.writeAttribute(QLatin1String("val"), s);

    if (d->majorGridlines.isValid() || d->majorGridlinesOn) {
        if (!d->majorGridlines.isValid()) {
            writer.writeEmptyElement("c:majorGridlines");
        }
        else {
            writer.writeStartElement("c:majorGridlines");
            d->majorGridlines.write(writer, "c:spPr");
            writer.writeEndElement();
        }
    }

    if (d->minorGridlines.isValid() || d->minorGridlinesOn) {
        if (!d->minorGridlines.isValid()) {
            writer.writeEmptyElement("c:minorGridlines");
        }
        else {
            writer.writeStartElement("c:minorGridlines");
            d->minorGridlines.write(writer, "c:spPr");
            writer.writeEndElement();
        }
    }

    if (d->title.isValid()) d->title.write(writer, QLatin1String("c:title"));

    writer.writeEmptyElement("c:crossAx");
    if (d->crossAxisPt != nullptr)
        writer.writeAttribute("val", QString::number(d->crossAxisPt->id()));
    else if (d->crossAxis != -1)
        writer.writeAttribute("val", QString::number(d->crossAxis));

    if (d->crossesType.has_value()) {
        switch (d->crossesType.value()) {
            case CrossesType::Maximum: writer.writeEmptyElement(QLatin1String("c:crosses"));
                writer.writeAttribute("val", "max");
                break;
            case CrossesType::Minimum: writer.writeEmptyElement(QLatin1String("c:crosses"));
                writer.writeAttribute("val", "min");
                break;
            case CrossesType::AutoZero: writer.writeEmptyElement(QLatin1String("c:crosses"));
                writer.writeAttribute("val", "autoZero");
                break;
            case CrossesType::Position: writer.writeEmptyElement(QLatin1String("c:crossesAt"));
                writer.writeAttribute("val", QString::number(d->crossesPosition.value_or(0.0)));
                break;
        }
    }

    d->numberFormat.write(writer, QLatin1String("c:numFmt"));

    if (d->majorTickMark.has_value()) {
        QString s;
        toString(d->majorTickMark.value(), s);
        writer.writeEmptyElement("c:majorTickMark");
        writer.writeAttribute("val", s);
    }
    if (d->minorTickMark.has_value()) {
        QString s;
        toString(d->minorTickMark.value(), s);
        writer.writeEmptyElement("c:minorTickMark");
        writer.writeAttribute("val", s);
    }
    if (d->tickLabelPosition.has_value()) {
        QString s; toString(d->tickLabelPosition.value(), s);
        writeEmptyElement(writer, QLatin1String("c:tickLblPos"), s);
    }
    d->shape.write(writer, QLatin1String("c:spPr"));

    if (d->textProperties.isValid()) d->textProperties.write(writer, QLatin1String("c:txPr"));

    if (d->type == Type::Category || d->type == Type::Date) {
        writeEmptyElement(writer, QLatin1String("c:auto"), d->axAuto);
    }
    if (d->type == Type::Category && d->labelAlignment.has_value()) {
        QString s; toString(d->labelAlignment.value(), s);
        writeEmptyElement(writer, QLatin1String("c:lblAlgn"), s);
    }

    if (d->type == Type::Category || d->type == Type::Date) {
        if (d->labelOffset.has_value())
            writeEmptyElement(writer, QLatin1String("c:lblOffset"), toST_PercentInt(d->labelOffset.value()));
    }

    if (d->type == Type::Value || d->type == Type::Date) {
        writeEmptyElement(writer, QLatin1String("c:majorUnit"), d->majorUnit);
        writeEmptyElement(writer, QLatin1String("c:minorUnit"), d->minorUnit);
    }
    if (d->type == Type::Date) {
        QString s;
        if (d->baseTimeUnit.has_value()) {
            toString(d->baseTimeUnit.value(), s);
            writeEmptyElement(writer, QLatin1String("c:baseTimeUnit"), s);
        }
        if (d->majorTimeUnit.has_value()) {
            toString(d->majorTimeUnit.value(), s);
            writeEmptyElement(writer, QLatin1String("c:majorTimeUnit"), s);
        }
        if (d->minorTimeUnit.has_value()) {
            toString(d->minorTimeUnit.value(), s);
            writeEmptyElement(writer, QLatin1String("c:minorTimeUnit"), s);
        }
    }
    if (d->type == Type::Value || d->type == Type::Series) {
        writeEmptyElement(writer, QLatin1String("c:tickLblSkip"), d->tickLabelSkip);
        writeEmptyElement(writer, QLatin1String("c:tickMarkSkip"), d->tickMarkSkip);
    }
    if (d->type == Type::Category)
        writeEmptyElement(writer, QLatin1String("c:noMultiLvlLbl"), d->noMultiLevelLabels);
    if (d->type == Type::Value) {
        if (d->crossesBetween.has_value()) {
            QString s; toString(d->crossesBetween.value(), s);
            writeEmptyElement(writer, QLatin1String("c:crossBetween"), s);
        }
        if (d->displayUnits.isValid()) d->displayUnits.write(writer, QLatin1String("c:dispUnits"));
    }

    writer.writeEndElement();
}

void Axis::read(QXmlStreamReader &reader)
{
    if (!d) d = new AxisPrivate;

    const auto &name = reader.name();
    d->type = Type::None;
    if (name == QLatin1String("catAx")) d->type = Type::Category;
    if (name == QLatin1String("valAx")) d->type = Type::Value;
    if (name == QLatin1String("serAx")) d->type = Type::Series;
    if (name == QLatin1String("dateAx")) d->type = Type::Date;

    while (!reader.atEnd()) {
        auto token = reader.readNext();
        if (token == QXmlStreamReader::StartElement) {
            const auto &a = reader.attributes();
            if (reader.name() == QLatin1String("axId")) {
                parseAttributeInt(a, QLatin1String("val"), d->id);
            }
            else if (reader.name() == QLatin1String("scaling")) {
                d->scaling.read(reader);
            }
            else if (reader.name() == QLatin1String("delete")) {
                bool vis;
                parseAttributeBool(a, QLatin1String("val"), vis);
                d->visible = !vis;
            }
            else if (reader.name() == QLatin1String("axPos")) {
                const auto &pos = a.value("val").toString();
                fromString(pos, d->position);
            }
            else if (reader.name() == QLatin1String("majorGridlines")) {
                d->majorGridlinesOn = true;
                if (reader.readNextStartElement())
                    d->majorGridlines.read(reader);
            }
            else if (reader.name() == QLatin1String("minorGridlines")) {
                d->minorGridlinesOn = true;
                if (reader.readNextStartElement())
                    d->minorGridlines.read(reader);
            }
            else if (reader.name() == QLatin1String("spPr")) {
                d->shape.read(reader);
            }
            else if (reader.name() == QLatin1String("title")) {
                d->title.read(reader);
            }
            else if (reader.name() == QLatin1String("txPr")) {
                d->textProperties.read(reader);
            }
            else if (reader.name() == QLatin1String("crossAx")) {
                d->crossAxis = a.value("val").toInt();
            }
            else if (reader.name() == QLatin1String("crosses")) {
                const auto &s = a.value(QLatin1String("val"));
                if (s == QLatin1String("autoZero")) d->crossesType = CrossesType::AutoZero;
                else if (s == QLatin1String("min")) d->crossesType = CrossesType::Minimum;
                if (s == QLatin1String("max")) d->crossesType = CrossesType::Maximum;
            }
            else if (reader.name() == QLatin1String("crossesAt")) {
                d->crossesType = CrossesType::Position;
                d->crossesPosition = a.value(QLatin1String("val")).toDouble();
            }
            else if (reader.name() == QLatin1String("majorTickMark")) {
                TickMark t;
                fromString(a.value(QLatin1String("val")).toString(), t);
                d->majorTickMark = t;
            }
            else if (reader.name() == QLatin1String("minorTickMark")) {
                TickMark t;
                fromString(a.value(QLatin1String("val")).toString(), t);
                d->minorTickMark = t;
            }
            else if (reader.name() == QLatin1String("numFmt"))
                d->numberFormat.read(reader);
            else if (reader.name() == QLatin1String("auto")) {
                parseAttributeBool(a, QLatin1String("val"), d->axAuto);
            }
            else if (reader.name() == QLatin1String("lblAlgn")) {
                LabelAlignment l; fromString(a.value(QLatin1String("val")).toString(), l);
                d->labelAlignment = l;
            }
            else if (reader.name() == QLatin1String("lblOffset")) {
                d->labelOffset = fromST_PercentInt(a.value(QLatin1String("val")));
            }
            else if (reader.name() == QLatin1String("majorUnit")) {
                parseAttributeDouble(a, QLatin1String("val"), d->majorUnit);
            }
            else if (reader.name() == QLatin1String("minorUnit")) {
                parseAttributeDouble(a, QLatin1String("val"), d->minorUnit);
            }
            else if (reader.name() == QLatin1String("tickLblPos")) {
                TickLabelPosition t; fromString(a.value(QLatin1String("val")).toString(), t);
                d->tickLabelPosition = t;
            }
            else if (reader.name() == QLatin1String("baseTimeUnit")) {
                TimeUnit t; fromString(a.value(QLatin1String("val")).toString(), t);
                d->baseTimeUnit = t;
            }
            else if (reader.name() == QLatin1String("majorTimeUnit")) {
                TimeUnit t; fromString(a.value(QLatin1String("val")).toString(), t);
                d->majorTimeUnit = t;
            }
            else if (reader.name() == QLatin1String("minorTimeUnit")) {
                TimeUnit t; fromString(a.value(QLatin1String("val")).toString(), t);
                d->minorTimeUnit = t;
            }
            else if (reader.name() == QLatin1String("tickLblSkip")) {
                parseAttributeInt(a, QLatin1String("val"), d->tickLabelSkip);
            }
            else if (reader.name() == QLatin1String("tickMarkSkip")) {
                parseAttributeInt(a, QLatin1String("val"), d->tickMarkSkip);
            }
            else if (reader.name() == QLatin1String("noMultiLvlLbl")) {
                parseAttributeBool(a, QLatin1String("val"), d->noMultiLevelLabels);
            }
            else if (reader.name() == QLatin1String("crossBetween")) {
                CrossesBetweenType t; fromString(a.value(QLatin1String("val")).toString(), t);
                d->crossesBetween = t;
            }
            else if (reader.name() == QLatin1String("dispUnits")) {
                d->displayUnits.read(reader);
            }
            else if (reader.name() == QLatin1String("extLst")) {
                reader.skipCurrentElement();
            }
        }
        else if (token == QXmlStreamReader::EndElement && reader.name() == name)
            break;
    }
}

bool Axis::operator ==(const Axis &other) const
{
    if (d == other.d) return true;
    if (!d || !other.d) return false;
    return *this->d.constData() == *other.d.constData();
}

bool Axis::operator !=(const Axis &other) const
{
    return !(operator==(other));
}

AxisPrivate::AxisPrivate() : position(Axis::Position::None), type(Axis::Type::None)
{

}

AxisPrivate::AxisPrivate(const AxisPrivate &other) : QSharedData(other),
    scaling{other.scaling}, visible{other.visible}, position{other.position},
    type{other.type}, crossAxis{other.crossAxis}, crossAxisPt{other.crossAxisPt},
    crossesType{other.crossesType}, crossesPosition{other.crossesPosition},
    majorGridlines{other.majorGridlines}, minorGridlines{other.minorGridlines},
    title{other.title}, textProperties{other.textProperties},
    numberFormat{other.numberFormat}, majorTickMark{other.majorTickMark},
    minorTickMark{other.minorTickMark}, tickLabelPosition{other.tickLabelPosition},
    shape{other.shape}, axAuto{other.axAuto}, labelOffset{other.labelOffset},
    labelAlignment{other.labelAlignment}, noMultiLevelLabels{other.noMultiLevelLabels},
    majorUnit{other.majorUnit}, minorUnit{other.minorUnit}, baseTimeUnit{other.baseTimeUnit},
    majorTimeUnit{other.majorTimeUnit}, minorTimeUnit{other.minorTimeUnit},
    tickLabelSkip{other.tickLabelSkip}, tickMarkSkip{other.tickMarkSkip},
    crossesBetween{other.crossesBetween}, displayUnits{other.displayUnits}
{

}

AxisPrivate::~AxisPrivate()
{

}

bool AxisPrivate::operator ==(const AxisPrivate &other) const
{
    //int id; //ids are always different
    if (scaling != other.scaling) return false;
    if (visible != other.visible) return false;
    if (position != other.position) return false;
    if (type != other.type) return false;
    if (crossAxis != other.crossAxis) return false;
    if (crossAxisPt != other.crossAxisPt) return false;
    if (crossesType != other.crossesType) return false;
    if (crossesPosition != other.crossesPosition) return false;
    if (majorGridlines != other.majorGridlines) return false;
    if (minorGridlines != other.minorGridlines) return false;
    if (majorTickMark != other.majorTickMark) return false;
    if (minorTickMark != other.minorTickMark) return false;
    if (numberFormat != other.numberFormat) return false;
    if (title != other.title) return false;
    if (majorTickMark != other.majorTickMark) return false;
    if (minorTickMark != other.minorTickMark) return false;
    if (axAuto != other.axAuto) return false;
    if (labelOffset != other.labelOffset) return false;
    if (labelAlignment != other.labelAlignment) return false;
    if (noMultiLevelLabels != other.noMultiLevelLabels) return false;
    if (majorUnit != other.majorUnit) return false;
    if (minorUnit != other.minorUnit) return false;
    if (baseTimeUnit != other.baseTimeUnit) return false;
    if (majorTimeUnit != other.majorTimeUnit) return false;
    if (minorTimeUnit != other.minorTimeUnit) return false;
    if (tickLabelSkip != other.tickLabelSkip) return false;
    if (tickMarkSkip != other.tickMarkSkip) return false;
    if (crossesBetween != other.crossesBetween) return false;
    if (displayUnits != other.displayUnits) return false;

    return true;
}

class DisplayUnitsPrivate : public QSharedData
{
public:
    DisplayUnitsPrivate();
    DisplayUnitsPrivate(const DisplayUnitsPrivate &other);
    ~DisplayUnitsPrivate();
    bool operator==(const DisplayUnitsPrivate &other) const;

    std::optional<double> customUnit;
    std::optional<DisplayUnits::BuiltInUnit> builtInUnit;
    Title title;
    bool labelVisible = true;
    ExtensionList extLst;
};

DisplayUnitsPrivate::DisplayUnitsPrivate()
{

}
DisplayUnitsPrivate::DisplayUnitsPrivate(const DisplayUnitsPrivate &other)
    : QSharedData{other}, customUnit{other.customUnit}, builtInUnit{other.builtInUnit},
      title{other.title}, labelVisible{other.labelVisible}, extLst{other.extLst}
{

}
DisplayUnitsPrivate::~DisplayUnitsPrivate()
{

}
bool DisplayUnitsPrivate::operator==(const DisplayUnitsPrivate &other) const
{
    if (customUnit != other.customUnit) return false;
    if (builtInUnit != other.builtInUnit) return false;
    if (title != other.title) return false;
    if (labelVisible != other.labelVisible) return false;
    if (extLst != other.extLst) return false;
    return true;
}

DisplayUnits::DisplayUnits()
{

}

DisplayUnits::DisplayUnits(double customUnit) : d{new DisplayUnitsPrivate}
{
    d->customUnit = customUnit;
}

DisplayUnits::DisplayUnits(DisplayUnits::BuiltInUnit unit) : d{new DisplayUnitsPrivate}
{
    d->builtInUnit = unit;
}

DisplayUnits::DisplayUnits(const DisplayUnits &other) : d{other.d}
{

}

DisplayUnits::~DisplayUnits()
{

}

DisplayUnits &DisplayUnits::operator=(const DisplayUnits &other)
{
    if (*this != other)
        d = other.d;
    return *this;
}

std::optional<double> DisplayUnits::customUnit() const
{
    if (d) return d->customUnit;
    return {};
}

void DisplayUnits::setCustomUnit(double customUnit)
{
    if (!d) d = new DisplayUnitsPrivate;
    d->customUnit = customUnit;
}

std::optional<DisplayUnits::BuiltInUnit> DisplayUnits::builtInUnit() const
{
    if (d) return d->builtInUnit;
    return {};
}

void DisplayUnits::setBuiltInUnit(DisplayUnits::BuiltInUnit builtInUnit)
{
    if (!d) d = new DisplayUnitsPrivate;
    d->builtInUnit = builtInUnit;
}

bool DisplayUnits::labelVisible() const
{
    if (d) return d->labelVisible;
    return {};
}

void DisplayUnits::setLabelVisible(bool visible)
{
    if (!d) d = new DisplayUnitsPrivate;
    d->labelVisible = visible;
}

Title DisplayUnits::title() const
{
    if (d) return d->title;
    return {};
}

Title &DisplayUnits::title()
{
    if (!d) d = new DisplayUnitsPrivate;
    return d->title;
}

void DisplayUnits::setTitle(const Title &title)
{
    if (!d) d = new DisplayUnitsPrivate;
    d->title = title;
}

bool DisplayUnits::isValid() const
{
    if (d) return true;
    return false;
}

void DisplayUnits::read(QXmlStreamReader &reader)
{
    if (!d) d = new DisplayUnitsPrivate;
    const auto &name = reader.name();
    while (!reader.atEnd()) {
        auto token = reader.readNext();
        if (token == QXmlStreamReader::StartElement) {
            const auto &a = reader.attributes();
            if (reader.name() == QLatin1String("custUnit"))
                parseAttributeDouble(a, QLatin1String("val"), d->customUnit);
            else if (reader.name() == QLatin1String("builtInUnit")) {
                BuiltInUnit u; fromString(a.value(QLatin1String("val")).toString(), u);
                d->builtInUnit = u;
            }
            else if (reader.name() == QLatin1String("dispUnitsLbl")) {
                d->labelVisible = true;
                d->title.read(reader);
            }
            else if (reader.name() == QLatin1String("extLst")) {
                d->extLst.read(reader);
            }
        }
        else if (token == QXmlStreamReader::EndElement && reader.name() == name)
            break;
    }
}

void DisplayUnits::write(QXmlStreamWriter &writer, const QString &name) const
{
    if (!isValid()) return;
    writer.writeStartElement(name);
    if (d->customUnit.has_value()) writeEmptyElement(writer, QLatin1String("c:custUnit"), d->customUnit);
    else if (d->builtInUnit.has_value()) {
        QString s; toString(d->builtInUnit.value(), s);
        writeEmptyElement(writer, QLatin1String("c:builtInUnit"), s);
    }
    if (d->labelVisible && d->title.isValid())
        d->title.write(writer, QLatin1String("c:dispUnitsLbl"));
    writer.writeEndElement();
}

bool DisplayUnits::operator ==(const DisplayUnits &other) const
{
    if (d == other.d) return true;
    if (!d || !other.d) return false;
    return *this->d.constData() == *other.d.constData();
}

bool DisplayUnits::operator !=(const DisplayUnits &other) const
{
    return !operator==(other);
}

DisplayUnits Axis::displayUnits() const
{
    if (d) return d->displayUnits;
    return {};
}

DisplayUnits &Axis::displayUnits()
{
    if (!d) d = new AxisPrivate;
    return d->displayUnits;
}

void Axis::setDisplayUnits(const DisplayUnits &displayUnits)
{
    if (!d) d = new AxisPrivate;
    d->displayUnits = displayUnits;
}

TextFormat &Axis::textProperties()
{
    if (!d) d = new AxisPrivate;
    return d->textProperties;
}

TextFormat Axis::textProperties() const
{
    if (d) return d->textProperties;
    return {};
}

void Axis::setTextProperties(const TextFormat &textProperties)
{
    if (!d) d = new AxisPrivate;
    d->textProperties = textProperties;
}

}
