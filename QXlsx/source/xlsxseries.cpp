#include "xlsxseries.h"

#include <optional>

QT_BEGIN_NAMESPACE_XLSX



SeriesPrivate::SeriesPrivate()
{

}

SeriesPrivate::SeriesPrivate(const SeriesPrivate &other) : QSharedData(other)
  //, TODO: all the members
{

}

SeriesPrivate::~SeriesPrivate()
{

}

Series::Series()
{

}

Series::Series(int index, int order)
{
    d = new SeriesPrivate;
    d->type = Type::None;
    d->index = index;
    d->order = order;
}

Series::Series(Series::Type type)
{
    d = new SeriesPrivate;
    d->type = type;
}

Series::Series(const Series &other): d(other.d)
{

}

Series::~Series()
{

}

int Series::index() const
{
    if (d) return d->index;
    return -1;
}

void Series::setIndex(int index)
{
    if (!d) d = new SeriesPrivate;
    d->index = index;
}

Series::Type Series::type() const
{
    if (d) return d->type;
    return Type::None;
}

void Series::setType(Series::Type type)
{
    if (!d) d = new SeriesPrivate;
    d->type = type;
}

int Series::order() const
{
    if (d) return d->order;
    return -1;
}

void Series::setOrder(int order)
{
    if (!d) d = new SeriesPrivate;
    d->order = order;
}

DataSource Series::categoryDataSource() const
{
    if (d) return d->xVal;
    return {};
}

DataSource &Series::categoryDataSource()
{
    if (!d) d = new SeriesPrivate;
    return d->xVal;
}

DataSource Series::valueDataSource() const
{
    if (d) return d->yVal;
    return {};
}

DataSource &Series::valueDataSource()
{
    if (!d) d = new SeriesPrivate;
    return d->yVal;
}

DataSource Series::bubbleSizeDataSource() const
{
    if (d) return d->bubbleVal;
    return {};
}

DataSource &Series::bubbleSizeDataSource()
{
    if (!d) d = new SeriesPrivate;
    return d->bubbleVal;
}

void Series::setCategorySource(const QString &reference, DataSource::Type type)
{
    if (!d) d = new SeriesPrivate;
    d->xVal = DataSource(type, reference);
}

void Series::setValueSource(const QString &reference, DataSource::Type type)
{
    if (!d) d = new SeriesPrivate;
    d->yVal = DataSource(type, reference);
}

void Series::setBubbleSizeSource(const QString &reference, DataSource::Type type)
{
    if (!d) d = new SeriesPrivate;
    d->bubbleVal = DataSource(type, reference);
}

void Series::setCategoryData(const QVector<double> &data)
{
    if (!d) d = new SeriesPrivate;
    d->xVal = DataSource(data);
}

void Series::setValueData(const QVector<double> &data)
{
    if (!d) d = new SeriesPrivate;
    d->yVal = DataSource(data);
}

void Series::setCategoryData(const QStringList &data)
{
    if (!d) d = new SeriesPrivate;
    d->xVal = DataSource(data);
}

void Series::setData(const QVector<double> &category, const QVector<double> &value)
{
    setCategoryData(category);
    setValueData(value);
}

void Series::setData(const QStringList &category, const QVector<double> &value)
{
    setCategoryData(category);
    setValueData(value);
}

void Series::setDataSource(const QString &categoryReference, const QString &valueReference)
{
    setCategorySource(categoryReference);
    setValueSource(valueReference);
}

void Series::setBubbleSizeData(const QVector<double> &data)
{
    if (!d) d = new SeriesPrivate;
    d->bubbleVal = DataSource(data);
}

void Series::setNameReference(const QString &reference)
{
    if (!d) d = new SeriesPrivate;
    d->name.setStringReference(reference);
}

void Series::setName(const Text &name)
{
    if (!d) d = new SeriesPrivate;
    d->name = name;
}

void Series::setName(const QString &name)
{
    if (!d) d = new SeriesPrivate;
    d->name.setPlainString(name);
}

Text Series::name() const
{
    if (d) return d->name;
    return {};
}

Text &Series::name()
{
    if (!d) d = new SeriesPrivate;
    return d->name;
}

DataSource Series::categoryData() const
{
    if (d) return d->xVal;
    return {};
}

DataSource Series::valueData() const
{
    if (d) return d->yVal;
    return {};
}

DataSource Series::bubbleSizeData() const
{
    if (d) return d->bubbleVal;
    return {};
}

DataSource &Series::categoryData()
{
    if (!d) d = new SeriesPrivate;
    return d->xVal;
}

DataSource &Series::valueData()
{
    if (!d) d = new SeriesPrivate;
    return d->yVal;
}

DataSource &Series::bubbleSizeData()
{
    if (!d) d = new SeriesPrivate;
    return d->bubbleVal;
}

void Series::setLine(const LineFormat &line)
{
    if (!d) d = new SeriesPrivate;
    d->shape.setLine(line);
}

LineFormat Series::line() const
{
    if (d && d->shape.isValid()) return d->shape.line();
    return {};
}

LineFormat &Series::line()
{
    if (!d) d = new SeriesPrivate;
    return d->shape.line();
}

void Series::setShape(const ShapeFormat &shape)
{
    if (!d) d = new SeriesPrivate;
    d->shape = shape;
}

ShapeFormat Series::shape() const
{
    if (!d) return d->shape;
    return {};
}

ShapeFormat &Series::shape()
{
    if (!d) d = new SeriesPrivate;
    return d->shape;
}

void Series::setMarker(const MarkerFormat &marker)
{
    if (!d) d = new SeriesPrivate;
    d->marker = marker;
}

MarkerFormat Series::marker() const
{
    if (d) return d->marker;
    return {};
}

MarkerFormat &Series::marker()
{
    if (!d) d = new SeriesPrivate;
    return d->marker;
}

void Series::setLabels(QVector<int> labels, Label::ShowParameters showFlags, Label::Position pos)
{
    if (!d) d = new SeriesPrivate;
    for (int label: labels)
        d->labels.addLabel(label, showFlags, pos);
    d->labels.setVisible(false);
}

void Series::setDefaultLabels(Label::ShowParameters showFlags, Label::Position pos)
{
    if (!d) d = new SeriesPrivate;
    Labels l(showFlags, pos);
    l.setVisible(true);
    d->labels = l;
}

Labels Series::defaultLabels() const
{
    if (d) return d->labels;
    return {};
}

Labels &Series::defaultLabels()
{
    if (!d) d = new SeriesPrivate;
    return d->labels;
}

Label &Series::label(int index)
{
    if (!d) d = new SeriesPrivate;
    return d->labels.label(index);
}

Label Series::label(int index) const
{
    if (d) return d->labels.label(index);
    return {};
}

Label &Series::labelForPoint(int index)
{
    if (!d) d = new SeriesPrivate;
    return d->labels.labelForPoint(index);
}

Label Series::labelForPoint(int index) const
{
    if (d) return d->labels.labelForPoint(index);
    return {};
}

std::optional<bool> Series::invertColorsIfNegative() const
{
    if (d) return d->invertIfNegative;
    return {};
}

void Series::setInvertColorsIfNegative(bool invert)
{
    if (!d) d = new SeriesPrivate;
    d->invertIfNegative = invert;
}

std::optional<Series::BarShape> Series::barShape() const
{
    if (d) return d->barShape;
    return {};
}

void Series::setBarShape(Series::BarShape barShape)
{
    if (!d) d = new SeriesPrivate;
    d->barShape = barShape;
}

std::optional<bool> Series::bubble3D() const
{
    if (d) return d->bubble3D;
    return {};
}

void Series::setBubble3D(bool val)
{
    if (!d) d = new SeriesPrivate;
    d->bubble3D = val;
}

std::optional<PictureOptions> Series::pictureOptions() const
{
    if (d) return d->pictureOptions;
    return {};
}

void Series::setPictureOptions(const PictureOptions &pictureOptions)
{
    if (!d) d = new SeriesPrivate;
    d->pictureOptions = pictureOptions;
}

void Series::read(QXmlStreamReader &reader)
{
    if (!d) d = new SeriesPrivate;

    const auto &name = reader.name();
    while (!reader.atEnd()) {
        auto token = reader.readNext();
        if (token == QXmlStreamReader::StartElement) {
            const auto &a = reader.attributes();
            if (reader.name() == QLatin1String("idx")) {
//                parseAttributeInt(a, QLatin1String(""), );
                parseAttributeInt(a, QLatin1String("val"), d->index);
            }
            else if (reader.name() == QLatin1String("order")) {
                parseAttributeInt(a, QLatin1String("val"), d->order);
            }
            else if (reader.name() == QLatin1String("tx")) {
                d->name.read(reader);
            }
            else if (reader.name() == QLatin1String("spPr")) {
                d->shape.read(reader);
            }
            else if (reader.name() == QLatin1String("marker")) {
                d->marker.read(reader);
            }
            else if (reader.name() == QLatin1String("dPt")) {
                DataPoint p;
                p.read(reader);
                d->dataPoints.append(p);
            }
            else if (reader.name() == QLatin1String("dLbls")) {
                d->labels.read(reader);
            }
            else if (reader.name() == QLatin1String("cat")) {
                reader.readNextStartElement();
                d->xVal.read(reader);
            }
            else if (reader.name() == QLatin1String("xVal")) {
                reader.readNextStartElement();
                d->xVal.read(reader);
            }
            else if (reader.name() == QLatin1String("val")) {
                reader.readNextStartElement();
                d->yVal.read(reader);
            }
            else if (reader.name() == QLatin1String("yVal")) {
                reader.readNextStartElement();
                d->yVal.read(reader);
            }
            else if (reader.name() == QLatin1String("invertIfNegative")) {
                parseAttributeBool(a, QLatin1String("val"), d->invertIfNegative);
            }
            else if (reader.name() == QLatin1String("shape")) {
                BarShape b;
                fromString(a.value(QLatin1String("val")).toString(), b);
                d->barShape = b;
            }
            else if (reader.name() == QLatin1String("explosion")) {
                parseAttributeInt(a, QLatin1String("val"), d->pieExplosion);
            }
            else if (reader.name() == QLatin1String("bubble3D")) {
                parseAttributeBool(a, QLatin1String("val"), d->bubble3D);
            }
            else if (reader.name() == QLatin1String("pictureOptions")) {
                PictureOptions p; p.read(reader);
                d->pictureOptions = p;
            }
//            else if (reader.name() == QLatin1String("")) {}
//            else if (reader.name() == QLatin1String("")) {}
//            else if (reader.name() == QLatin1String("")) {}
//            else if (reader.name() == QLatin1String("")) {}
//            else if (reader.name() == QLatin1String("")) {}
        }
        else if (token == QXmlStreamReader::EndElement && reader.name() == name)
            break;
    }
}

void Series::write(QXmlStreamWriter &writer) const
{
    Q_ASSERT(d->type != Type::None);

    writer.writeStartElement(QLatin1String("c:ser"));
    writeEmptyElement(writer, QLatin1String("c:idx"), d->index);
    writeEmptyElement(writer, QLatin1String("c:order"), d->order);
    if (d->name.isValid()) d->name.write(writer, QLatin1String("c:tx"));
    if (d->shape.isValid()) d->shape.write(writer, QLatin1String("a:spPr"));
    if (!d->dataPoints.isEmpty() && d->type != Type::Surface) {
        for (auto &dPt: d->dataPoints) dPt.write(writer);
    }
    if (d->labels.isValid() && d->type != Type::Surface) d->labels.write(writer);
    QString xVal = (d->type == Type::Scatter || d->type == Type::Bubble) ? QLatin1String("c:xVal") : QLatin1String("c:cat");
    d->xVal.write(writer, xVal);
    QString yVal = (d->type == Type::Scatter || d->type == Type::Bubble) ? QLatin1String("c:yVal") : QLatin1String("c:val");
    d->yVal.write(writer, yVal);
    if (d->type == Type::Bubble) d->bubbleVal.write(writer, QLatin1String("c:bubbleSize"));

    if (d->type == Type::Line || d->type == Type::Scatter || d->type == Type::Radar)
        if (d->marker.isValid()) d->marker.write(writer);
    if (d->type == Type::Line || d->type == Type::Scatter)
        writeEmptyElement(writer, QLatin1String("c:smooth"), d->smooth);
    if (d->type == Type::Bubble || d->type == Type::Bar)
        writeEmptyElement(writer, QLatin1String("c:invertIfNegative"), d->invertIfNegative);
    if (d->barShape.has_value() && d->type == Type::Bar) {
        writer.writeEmptyElement(QLatin1String("c:shape"));
        QString s; toString(d->barShape.value(), s);
        writer.writeAttribute(QLatin1String("val"), s);
    }
    if (d->type == Type::Pie)
        writeEmptyElement(writer, QLatin1String("c:explosion"), d->pieExplosion);
    if (d->type == Type::Bubble)
        writeEmptyElement(writer, QLatin1String("c:bubble3D"), d->bubble3D);
    if (d->type == Type::Area || d->type == Type::Bar)
        if (d->pictureOptions.has_value())
            d->pictureOptions->write(writer, QLatin1String("c:pictureOptions"));
    if (d->type == Type::Line || d->type == Type::Scatter || d->type == Type::Bar ||
        d->type == Type::Area || d->type == Type::Bubble)
        if (d->errorBars.has_value()) d->errorBars->write(writer, QLatin1String("c:errBars"));

    writer.writeEndElement();
}

void PictureOptions::read(QXmlStreamReader &reader)
{
    const auto &name = reader.name();
    while (!reader.atEnd()) {
        auto token = reader.readNext();
        if (token == QXmlStreamReader::StartElement) {
            const auto &a = reader.attributes();
            if (reader.name() == QLatin1String("applyToFront"))
                //parseAttributeInt(a, QLatin1String(""), );
                parseAttributeBool(a, QLatin1String("val"), applyToFront);
            else if (reader.name() == QLatin1String("applyToSides"))
                parseAttributeBool(a, QLatin1String("val"), applyToSides);
            else if (reader.name() == QLatin1String("applyToEnd"))
                parseAttributeBool(a, QLatin1String("val"), applyToEnd);
            else if (reader.name() == QLatin1String("pictureFormat")) {
                auto s = a.value(QLatin1String("val"));
                if (s == QLatin1String("stretch")) format = PictureFormat::Stretch;
                if (s == QLatin1String("stack")) format = PictureFormat::Stack;
                if (s == QLatin1String("stackScale")) format = PictureFormat::StackScale;
            }
            else if (reader.name() == QLatin1String("pictureStackUnit"))
                parseAttributeDouble(a, QLatin1String("val"), pictureStackUnit);
        }
        else if (token == QXmlStreamReader::EndElement && reader.name() == name)
            break;
    }
}

void PictureOptions::write(QXmlStreamWriter &writer, const QString &name) const
{
    writer.writeStartElement(name);
    if (applyToFront.has_value()) {
        writer.writeEmptyElement(QLatin1String("c:applyToFront"));
        writer.writeAttribute(QLatin1String("val"), toST_Boolean(applyToFront.value()));
    }
    if (applyToSides.has_value()) {
        writer.writeEmptyElement(QLatin1String("c:applyToSides"));
        writer.writeAttribute(QLatin1String("val"), toST_Boolean(applyToSides.value()));
    }
    if (applyToEnd.has_value()) {
        writer.writeEmptyElement(QLatin1String("c:applyToEnd"));
        writer.writeAttribute(QLatin1String("val"), toST_Boolean(applyToEnd.value()));
    }
    if (format.has_value()) {
        writer.writeEmptyElement(QLatin1String("c:pictureFormat"));
        switch (format.value()) {
        case PictureFormat::Stretch: writer.writeAttribute(QLatin1String("val"), QLatin1String("stretch")); break;
        case PictureFormat::Stack: writer.writeAttribute(QLatin1String("val"), QLatin1String("stack")); break;
        case PictureFormat::StackScale: writer.writeAttribute(QLatin1String("val"), QLatin1String("stackScale")); break;
        }
    }
    if (pictureStackUnit.has_value()) {
        writer.writeEmptyElement(QLatin1String("c:pictureStackUnit"));
        writer.writeAttribute(QLatin1String("val"), QString::number(pictureStackUnit.value()));
    }
    writer.writeEndElement();
}

void DataPoint::read(QXmlStreamReader &reader)
{
    const auto &name = reader.name();
    while (!reader.atEnd()) {
        auto token = reader.readNext();
        if (token == QXmlStreamReader::StartElement) {
            const auto &a = reader.attributes();
            if (reader.name() == QLatin1String("idx"))
                //parseAttributeInt(a, QLatin1String(""), );
                parseAttributeInt(a, QLatin1String("val"), index);
            else if (reader.name() == QLatin1String("invertIfNegative"))
                parseAttributeBool(a, QLatin1String("val"), invertIfNegative);
            else if (reader.name() == QLatin1String("marker"))
                marker.read(reader);
            else if (reader.name() == QLatin1String("explosion"))
                parseAttributeInt(a, QLatin1String("val"), explosion);
            else if (reader.name() == QLatin1String("bubble3D"))
                parseAttributeBool(a, QLatin1String("val"), bubble3D);
            else if (reader.name() == QLatin1String("spPr"))
                shape.read(reader);
            else if (reader.name() == QLatin1String("pictureOptions")) {
                PictureOptions p;
                p.read(reader);
                pictureOptions = p;
            }
        }
        else if (token == QXmlStreamReader::EndElement && reader.name() == name)
            break;
    }
}

void DataPoint::write(QXmlStreamWriter &writer) const
{
    writer.writeStartElement(QLatin1String("c:dPt"));
    writer.writeEmptyElement(QLatin1String("c:idx"));
    writer.writeAttribute(QLatin1String("val"), QString::number(index));

    if (invertIfNegative.has_value()) {
        writer.writeEmptyElement(QLatin1String("c:invertIfNegative"));
        writer.writeAttribute(QLatin1String("val"), toST_Boolean(invertIfNegative.value()));
    }
    if (bubble3D.has_value()) {
        writer.writeEmptyElement(QLatin1String("c:bubble3D"));
        writer.writeAttribute(QLatin1String("val"), toST_Boolean(bubble3D.value()));
    }
    if (marker.isValid()) marker.write(writer);
    if (explosion.has_value()) {
        writer.writeEmptyElement(QLatin1String("c:explosion"));
        writer.writeAttribute(QLatin1String("val"), QString::number(explosion.value()));
    }
    if (shape.isValid()) shape.write(writer, QLatin1String("spPr"));
    if (pictureOptions.has_value()) pictureOptions->write(writer, QLatin1String("c:pictureOptions"));

    writer.writeEndElement();
}

DataSource::DataSource() : type{Type::None}
{

}

DataSource::DataSource(DataSource::Type type) : type{type}
{

}

DataSource::DataSource(const QVector<double> &numberData)
    : type{Type::NumberLiterals}, numberData{numberData}
{

}

DataSource::DataSource(const QStringList &stringData)
    : type{Type::StringLiterals}, stringData{stringData.toVector()}
{

}

DataSource::DataSource(DataSource::Type type, const QString &reference)
    : type{type}, reference{reference}
{
    // type checking
    if (!reference.isEmpty() && (type == Type::NumberReference || type == Type::StringReference
                                 || type == Type::MultiLevel))
        type = Type::None;
}

bool DataSource::isValid() const
{
    if (type == Type::None) return false;
    if ((type == Type::NumberReference || type == Type::StringReference
        || type == Type::MultiLevel ) && reference.isEmpty()) return false;
    if (type == Type::NumberLiterals && numberData.isEmpty()) return false;
    if (type == Type::StringLiterals && stringData.isEmpty()) return false;
    return true;
}

bool DataSource::operator==(const DataSource &other) const
{
    if (type != other.type) return false;
    if (reference != other.reference) return false;
    if (numberData != other.numberData) return false;
    if (stringData != other.stringData) return false;
    if (formatCode != other.formatCode) return false;
    return true;
}

bool DataSource::operator!=(const DataSource &other) const
{
    return !operator==(other);
}

void DataSource::read(QXmlStreamReader &reader)
{
    const auto &name = reader.name();
    if (type == Type::None) {
        if (name == QLatin1String("numRef")) type = Type::NumberReference;
        if (name == QLatin1String("numLit")) type = Type::NumberLiterals;
        if (name == QLatin1String("strRef")) type = Type::StringReference;
        if (name == QLatin1String("strLit")) type = Type::StringLiterals;
        if (name == QLatin1String("multiLvlStrRef")) type = Type::MultiLevel;
    }

    while (!reader.atEnd()) {
        auto token = reader.readNext();
        if (token == QXmlStreamReader::StartElement) {
            if (name == QLatin1String("numRef")) {
                if (reader.name() == QLatin1String("f")) reference = reader.readElementText();
                else if (reader.name() == QLatin1String("numCashe")) {
                    //TODO: add reading of cashe
                }
            }
            else if (name == QLatin1String("numLit")) {
                if (reader.name() == QLatin1String("formatCode")) formatCode = reader.readElementText();
                else if (reader.name() == QLatin1String("ptCount")) {
                    numberData.resize(reader.attributes().value(QLatin1String("val")).toInt());
                }
                else if (reader.name() == QLatin1String("pt")) {
                    // TODO: add reading and writing of individual formatCodes
                    int idx = reader.attributes().value(QLatin1String("idx")).toInt();
                    numberData[idx] = reader.readElementText().toDouble();
                }
            }
            else if (name == QLatin1String("strRef")) {
                if (reader.name() == QLatin1String("f")) reference = reader.readElementText();
                else if (reader.name() == QLatin1String("strCashe")) {
                    //TODO: add reading of cashe
                }
            }
            else if (name == QLatin1String("strLit")) {
                if (reader.name() == QLatin1String("formatCode")) formatCode = reader.readElementText();
                else if (reader.name() == QLatin1String("ptCount")) {
                    stringData.resize(reader.attributes().value(QLatin1String("val")).toInt());
                }
                else if (reader.name() == QLatin1String("pt")) {
                    // TODO: add reading and writing of individual formatCodes
                    int idx = reader.attributes().value(QLatin1String("idx")).toInt();
                    reader.readNextStartElement();
                    stringData[idx] = reader.readElementText();
                }
            }
            else if (name == QLatin1String("multiLvlStrRef")) {
                if (reader.name() == QLatin1String("f")) reference = reader.readElementText();
                else if (reader.name() == QLatin1String("multiLvlStrCache")) {
                    //TODO: add reading of cashe
                }
            }
        }
        else if (token == QXmlStreamReader::EndElement && reader.name() == name)
            break;
    }
}

void DataSource::write(QXmlStreamWriter &writer, const QString &name) const
{
    if (type == Type::None) {
        return;
    }

    writer.writeStartElement(name);
    if (type == Type::NumberReference && !reference.isEmpty()) {
        //numRef
        writer.writeStartElement(QLatin1String("c:numRef"));
        writer.writeTextElement(QLatin1String("c:f"), reference);
        //TODO: writing of cashe
        writer.writeEndElement();
    }
    if (type == Type::NumberLiterals && !numberData.isEmpty()) {
        //numLit
        writer.writeStartElement(QLatin1String("c:numLit"));
        if (!formatCode.isEmpty()) writer.writeTextElement(QLatin1String("c:formatCode"), formatCode);
        writeEmptyElement(writer, QLatin1String("ptCount"), int(numberData.size()));
        for (int idx = 0; idx < numberData.size(); ++idx) {
            writer.writeStartElement(QLatin1String("c:pt"));
            writer.writeAttribute(QLatin1String("idx"), QString::number(idx));
            writer.writeTextElement(QLatin1String("c:v"), QString::number(numberData[idx]));
            writer.writeEndElement();
        }
        writer.writeEndElement();
    }
    if (type == Type::StringReference && !reference.isEmpty()) {
        //strRef
        writer.writeStartElement(QLatin1String("c:strRef"));
        writer.writeTextElement(QLatin1String("c:f"), reference);
        //TODO: writing of cashe
        writer.writeEndElement();
    }
    if (type == Type::StringLiterals && !stringData.isEmpty()) {
        //strLit
        writer.writeStartElement(QLatin1String("c:numLit"));
        writeEmptyElement(writer, QLatin1String("ptCount"), int(stringData.size()));
        for (int idx = 0; idx < stringData.size(); ++idx) {
            writer.writeStartElement(QLatin1String("c:pt"));
            writer.writeAttribute(QLatin1String("idx"), QString::number(idx));
            writer.writeTextElement(QLatin1String("c:v"), stringData[idx]);
            writer.writeEndElement();
        }
        writer.writeEndElement();
    }
    if (type == Type::MultiLevel && !reference.isEmpty()) {
        writer.writeStartElement(QLatin1String("c:multiLvlStrRef"));
        writer.writeTextElement(QLatin1String("c:f"), reference);
        //TODO: writing of cashe
        writer.writeEndElement();
    }


    writer.writeEndElement();
}

bool ErrorBars::isValid() const
{
    //TODO: rewrite for qshareddata
    return true;
}

void ErrorBars::read(QXmlStreamReader &reader)
{
    const auto &name = reader.name();

    while (!reader.atEnd()) {
        auto token = reader.readNext();
        if (token == QXmlStreamReader::StartElement) {
            const auto &a = reader.attributes();
            if (name == QLatin1String("errDir")) {
                Direction d; fromString(a.value(QLatin1String("val")).toString(), d);
                direction = d;
            }
            else if (name == QLatin1String("errBarType")) {
                Type t; fromString(a.value(QLatin1String("val")).toString(), t);
                barType = t;
            }
            else if (name == QLatin1String("errValType")) {
                ValueType t; fromString(a.value(QLatin1String("val")).toString(), t);
                valueType = t;
            }
            else if (name == QLatin1String("noEndCap")) {
                parseAttributeBool(a, QLatin1String("val"), noEndCap);
            }
            else if (name == QLatin1String("plus")) {
                reader.readNextStartElement();
                plus.read(reader);
            }
            else if (name == QLatin1String("minus")) {
                reader.readNextStartElement();
                minus.read(reader);
            }
            else if (name == QLatin1String("val")) {
                value = a.value(QLatin1String("val")).toDouble();
            }
            else if (name == QLatin1String("spPr")) {
                shape.read(reader);
            }
        }
        else if (token == QXmlStreamReader::EndElement && reader.name() == name)
            break;
    }
}

void ErrorBars::write(QXmlStreamWriter &writer, const QString &name) const
{
    QString s;
    writer.writeStartElement(name);
    if (direction.has_value()) {
        writer.writeEmptyElement(QLatin1String("c:errDir"));
        toString(direction.value(), s);
        writer.writeAttribute(QLatin1String("val"), s);
    }
    writer.writeEmptyElement(QLatin1String("c:errBarType"));
    toString(barType, s); writer.writeAttribute(QLatin1String("val"), s);

    writer.writeEmptyElement(QLatin1String("c:errValType"));
    toString(valueType, s); writer.writeAttribute(QLatin1String("val"), s);

    writeEmptyElement(writer, QLatin1String("c:noEndCap"), noEndCap);

    if (plus.isValid()) plus.write(writer, QLatin1String("c:plus"));
    if (minus.isValid()) minus.write(writer, QLatin1String("c:minus"));
    if (value.has_value()) writeEmptyElement(writer, QLatin1String("c:value"), value);
    if (shape.isValid()) shape.write(writer, QLatin1String("a:spPr"));
    writer.writeEndElement();
}

bool Series::operator ==(const Series &other) const
{
    return d == other.d;
}

bool Series::operator !=(const Series &other) const
{
    return d != other.d;
}

QT_END_NAMESPACE_XLSX




