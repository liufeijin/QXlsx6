#include "xlsxautofilter.h"

#include <QDebug>
#include "xlsxmain.h"

namespace QXlsx {

struct AutoFilterColumn
{
    void write(QXmlStreamWriter &writer, const QString &name, int index) const;
    void read(QXmlStreamReader &reader);
    bool operator==(const AutoFilterColumn &other) const;
    bool isValid() const;

    void readFilters(QXmlStreamReader &reader);
    void readCustomFilters(QXmlStreamReader &reader);


    std::optional<bool> filterButtonHidden; //default = false
    std::optional<bool> showFilterOptions; //default = true

    Filter filter;

};

struct SortParameters
{
    void write(QXmlStreamWriter &writer, const QString &name) const;
    void read(QXmlStreamReader &reader);
    bool operator==(const SortParameters &other) const;
    bool isValid() const;
};

class AutoFilterPrivate : public QSharedData
{
public:
    AutoFilterPrivate();
    AutoFilterPrivate(const AutoFilterPrivate& other);
    ~AutoFilterPrivate();
    bool operator==(const AutoFilterPrivate& other) const;

    CellRange range;
    QMap<int, AutoFilterColumn> columns;
    SortParameters sort;
    ExtensionList extLst;
};

AutoFilterPrivate::AutoFilterPrivate()
{

}

AutoFilterPrivate::AutoFilterPrivate(const AutoFilterPrivate &other) : QSharedData(other),
    range{other.range}, columns{other.columns}, sort{other.sort}, extLst{other.extLst}
{

}

AutoFilterPrivate::~AutoFilterPrivate()
{

}

bool AutoFilterPrivate::operator==(const AutoFilterPrivate &other) const
{
    if (range != other.range) return false;
    if (columns != other.columns) return false;
    if (!(sort == other.sort)) return false;
    if (extLst != other.extLst) return false;

    return true;
}

AutoFilter::AutoFilter()
{
    //The d pointer is initialized with a null pointer
}

AutoFilter::AutoFilter(const CellRange &range)
{
    d = new AutoFilterPrivate;
    d->range = range;
}

AutoFilter::AutoFilter(const AutoFilter &other) : d{other.d}
{

}

AutoFilter &AutoFilter::operator=(const AutoFilter &other)
{
    if (*this != other) d = other.d;
    return *this;
}

AutoFilter::~AutoFilter()
{

}

bool AutoFilter::isValid() const
{
    if (d)
        return true;
    return false;
}

CellRange AutoFilter::range() const
{
    if (d) return d->range;
    return {};
}

void AutoFilter::setRange(const CellRange &range)
{
    if (!d) d = new AutoFilterPrivate;
    d->range = range;
}

bool AutoFilter::showFilterButton(int column) const
{
    if (d) return !d->columns.value(column).filterButtonHidden.value_or(false);
    return true;
}

void AutoFilter::setShowFilterButton(int column, bool show)
{
    if (!d) d = new AutoFilterPrivate;
    d->columns[column].filterButtonHidden = !show;
}
void AutoFilter::setShowFilterButton(bool show)
{
    if (!d) d = new AutoFilterPrivate;
    for (int col = d->range.firstColumn(); col <= d->range.lastColumn(); ++col)
        d->columns[col].filterButtonHidden = !show;
}

bool AutoFilter::showFilterOptions(int column) const
{
    if (d) return d->columns.value(column).showFilterOptions.value_or(true);
    return true;
}
void AutoFilter::setShowFilterOptions(int column, bool show)
{
    if (!d) d = new AutoFilterPrivate;
    d->columns[column].showFilterOptions = show;
}
void AutoFilter::setShowFilterOptions(bool show)
{
    if (!d) d = new AutoFilterPrivate;
    for (int col = d->range.firstColumn(); col <= d->range.lastColumn(); ++col)
        d->columns[col].showFilterOptions = show;
}

void AutoFilter::setFilterByTopN(int column, double value, std::optional<double> filterBy, bool usePercents)
{
    if (!d) d = new AutoFilterPrivate;
    if (column < 0) return;
    AutoFilterColumn c;
    c.filter.type = Filter::Type::Top10;
    c.filter.usePercents = usePercents;
    c.filter.top10val = value;
    c.filter.top10filterVal = filterBy;
    d->columns.insert(column, c);
}

void AutoFilter::setFilterByBottomN(int column, double value, std::optional<double> filterBy, bool usePercents)
{
    if (!d) d = new AutoFilterPrivate;
    if (column < 0) return;
    AutoFilterColumn c;
    c.filter.type = Filter::Type::Top10;
    c.filter.usePercents = usePercents;
    c.filter.top10val = value;
    c.filter.top10filterVal = filterBy;
    c.filter.top = false;
    d->columns.insert(column, c);
}

void AutoFilter::setFilterByValues(int column, const QVariantList &values)
{
    if (!d) d = new AutoFilterPrivate;
    if (column < 0 || values.isEmpty()) return;
    AutoFilterColumn c;
    c.filter.type = Filter::Type::Values;
    c.filter.filters = values;
    d->columns.insert(column, c);
}

QVariantList AutoFilter::filterValues(int column) const
{
    if (!d) return {};
    if (d->columns.contains(column))
        return d->columns.value(column).filter.filters;
    return {};
}

void AutoFilter::setPredicate(int column, Filter::Predicate op1, const QVariant &val1, Filter::Operation op, Filter::Predicate op2, const QVariant &val2)
{
    if (!d) d = new AutoFilterPrivate;
    if (column < 0) return;
    AutoFilterColumn c;
    c.filter.type = Filter::Type::Custom;
    c.filter.val1 = val1;
    c.filter.val2 = val2;
    c.filter.op1 = op1;
    c.filter.op2 = op2;
    c.filter.op = op;
    d->columns.insert(column, c);
}

void AutoFilter::setDynamicFilter(int column, Filter::DynamicFilterType type)
{
    if (!d) d = new AutoFilterPrivate;
    if (column < 0) return;
    AutoFilterColumn c;
    c.filter.type = Filter::Type::Dynamic;
    c.filter.dynamicType = type;
    d->columns.insert(column, c);
}

Filter::DynamicFilterType AutoFilter::dynamicFilterType(int column) const
{
    if (!d) return Filter::DynamicFilterType::Invalid;
    return d->columns.value(column).filter.dynamicType;
}

Filter::Type AutoFilter::filterType(int column) const
{
    if (!d) return Filter::Type::Invalid;
    return d->columns.value(column).filter.type;
}

void AutoFilter::removeFilter(int column)
{
    if (!d) return;
    d->columns.remove(column);
}

Filter AutoFilter::filter(int column) const
{
    if (d) return d->columns.value(column).filter;
    return {};
}

Filter &AutoFilter::filter(int column)
{
    if (!d) d = new AutoFilterPrivate;
    auto &col = d->columns[column];
    return col.filter;
}

void AutoFilter::setFilter(int column, const Filter &filter)
{
    if (!d) d = new AutoFilterPrivate;
    auto &col = d->columns[column];
    col.filter = filter;
}

void AutoFilter::write(QXmlStreamWriter &writer, const QString &name) const
{
    if (!isValid()) return;

    writer.writeStartElement(name);

    if (d->range.isValid()) writer.writeAttribute(QLatin1String("ref"), d->range.toString());

    for (auto it = d->columns.constBegin(); it != d->columns.constEnd(); ++it)
        it.value().write(writer, QLatin1String("filterColumn"), it.key());
    d->sort.write(writer, QLatin1String("sortState"));
    d->extLst.write(writer, QLatin1String("extLst"));
    writer.writeEndElement();
}

void AutoFilter::read(QXmlStreamReader &reader)
{
    if (!d) d = new AutoFilterPrivate;
    const auto &name = reader.name();
    QString s;
    parseAttributeString(reader.attributes(), QLatin1String("ref"), s);
    d->range = CellRange(s);

    int index = 0;
    while (!reader.atEnd()) {
        const auto token = reader.readNext();
        if (token == QXmlStreamReader::StartElement) {
            if (reader.name() == QLatin1String("filterColumn")) {
                AutoFilterColumn c;
                c.read(reader);
                d->columns.insert(index++, c);
            }
            else if (reader.name() == QLatin1String("sortState")) {
                d->sort.read(reader);
            }
            else if (reader.name() == QLatin1String("extLst")) {
                d->extLst.read(reader);
            }
            else reader.skipCurrentElement();
        }
        else if (token == QXmlStreamReader::EndElement && reader.name() == name)
            break;
    }
}

bool AutoFilter::operator==(const AutoFilter &other) const
{
    if (d == other.d) return true;
    if (!d || !other.d) return false;
    return *this->d.constData() == *other.d.constData();
}

bool AutoFilter::operator!=(const AutoFilter &other) const
{
    return !operator==(other);
}

QXlsx::AutoFilter::operator QVariant() const
{
    const auto& cref
#if QT_VERSION >= 0x060000 // Qt 6.0 or over
        = QMetaType::fromType<AutoFilter>();
#else
        = qMetaTypeId<AutoFilter>() ;
#endif
    return QVariant(cref, this);
}

#ifndef QT_NO_DEBUG_STREAM
QDebug operator<<(QDebug dbg, const AutoFilter &f)
{
    QDebugStateSaver saver(dbg);

    dbg.nospace() << "QXlsx::AutoFilter(" ;
    dbg.nospace() << f.d->range.toString() << ", ";
    //TODO
    dbg.nospace() << ")";
    return dbg;
}
#endif

void AutoFilterColumn::write(QXmlStreamWriter &writer, const QString &name, int index) const
{
    if (!isValid()) return;

    writer.writeStartElement(name);
    writeAttribute(writer, QLatin1String("colId"), index);
    writeAttribute(writer, QLatin1String("hiddenButton"), this->filterButtonHidden);
    writeAttribute(writer, QLatin1String("showButton"), this->showFilterOptions);
    switch (filter.type) {
        case Filter::Type::Values: {
            if (!filter.filters.isEmpty()) {
                writer.writeStartElement(QLatin1String("filters"));
                for (auto &val: filter.filters)
                    writeEmptyElement(writer, QLatin1String("filter"), val.toString());
                writer.writeEndElement();
            }
            break;
        }
        case Filter::Type::Custom: {
            if (filter.val1.isValid() || filter.val2.isValid()) {
                writer.writeStartElement(QLatin1String("customFilters"));
                if (filter.op.value_or(Filter::Operation::Or) == Filter::Operation::And)
                    writeAttribute(writer, QLatin1String("and"), true);
                if (filter.val1.isValid()) {
                    writer.writeEmptyElement(QLatin1String("customFilter"));
                    writer.writeAttribute(QLatin1String("val"), filter.val1.toString());
                    QString s;
                    switch (filter.op1.value_or(Filter::Predicate::Equal)) {
                        case Filter::Predicate::LessThan: s = "lessThan"; break;
                        case Filter::Predicate::LessThanOrEqual: s = "lessThanOrEqual"; break;
                        case Filter::Predicate::GreaterThanOrEqual: s = "greaterThanOrEqual"; break;
                        case Filter::Predicate::GreaterThan: s = "greaterThan"; break;
                        case Filter::Predicate::NotEqual: s = "notEqual"; break;
                        default: break;
                    }
                    writeAttribute(writer, QLatin1String("operator"), s);
                }
                if (filter.val2.isValid()) {
                    writer.writeEmptyElement(QLatin1String("customFilter"));
                    writer.writeAttribute(QLatin1String("val"), filter.val2.toString());
                    QString s;
                    switch (filter.op2.value_or(Filter::Predicate::Equal)) {
                        case Filter::Predicate::LessThan: s = "lessThan"; break;
                        case Filter::Predicate::LessThanOrEqual: s = "lessThanOrEqual"; break;
                        case Filter::Predicate::GreaterThanOrEqual: s = "greaterThanOrEqual"; break;
                        case Filter::Predicate::GreaterThan: s = "greaterThan"; break;
                        case Filter::Predicate::NotEqual: s = "notEqual"; break;
                        default: break;
                    }
                    writeAttribute(writer, QLatin1String("operator"), s);
                }
                writer.writeEndElement();
            }
            break;
        }
        case Filter::Type::Dynamic: {
            if (filter.dynamicType != Filter::DynamicFilterType::Invalid) {
                writer.writeEmptyElement(QLatin1String("dynamicFilter"));
                writeAttribute(writer, QLatin1String("type"), Filter::toString(filter.dynamicType));
                writeAttribute(writer, QLatin1String("val"), filter.val);
                if (filter.dateTime.has_value()) {
                    double val = datetimeToNumber(filter.dateTime.value());
                    writeAttribute(writer, QLatin1String("valIso"), val);
                }
                if (filter.maxDateTime.has_value()) {
                    double val = datetimeToNumber(filter.maxDateTime.value());
                    writeAttribute(writer, QLatin1String("maxValIso"), val);
                }
            }
            break;
        }
        case Filter::Type::Top10: {
            if (filter.top10val.has_value()) {
                writer.writeEmptyElement(QLatin1String("top10"));
                writeAttribute(writer, QLatin1String("top"), filter.top);
                writeAttribute(writer, QLatin1String("percent"), filter.usePercents);
                writeAttribute(writer, QLatin1String("val"), filter.top10val);
                writeAttribute(writer, QLatin1String("filterVal"), filter.top10filterVal);
            }
            break;
        }
        default: break;
    }

    writer.writeEndElement();
}

void AutoFilterColumn::read(QXmlStreamReader &reader)
{
    const auto &name = reader.name();
    const auto &a = reader.attributes();
    parseAttributeBool(a, QLatin1String("hiddenButton"), filterButtonHidden);
    parseAttributeBool(a, QLatin1String("showButton"), showFilterOptions);
    int idx = 0;
    parseAttributeInt(a, QLatin1String("colId"), idx);

    while (!reader.atEnd()) {
        auto token = reader.readNext();
        if (token == QXmlStreamReader::StartElement) {
            const auto &a = reader.attributes();
            if (reader.name() == QLatin1String("filters")) {
                filter.type = Filter::Type::Values;
                readFilters(reader);
            }
            else if (reader.name() == QLatin1String("top10")) {
                filter.type = Filter::Type::Top10;
                parseAttributeBool(a, QLatin1String("top"), filter.top);
                parseAttributeBool(a, QLatin1String("percent"), filter.usePercents);
                parseAttributeDouble(a, QLatin1String("val"), filter.top10val);
                parseAttributeDouble(a, QLatin1String("filterVal"), filter.top10filterVal);
            }
            else if (reader.name() == QLatin1String("customFilters")) {
                filter.type = Filter::Type::Custom;
                readCustomFilters(reader);
            }
            else if (reader.name() == QLatin1String("dynamicFilter")) {
                filter.type = Filter::Type::Dynamic;

                Filter::fromString(a.value(QLatin1String("type")).toString(), filter.dynamicType);
                parseAttributeDouble(a, QLatin1String("val"), filter.val);
                if (a.hasAttribute(QLatin1String("valIso"))) {
                    auto dt = datetimeFromNumber(a.value(QLatin1String("valIso")).toDouble());
                    filter.dateTime = dt.toDateTime();
                }
                if (a.hasAttribute(QLatin1String("maxValIso"))) {
                    auto dt = datetimeFromNumber(a.value(QLatin1String("maxValIso")).toDouble());
                    filter.maxDateTime = dt.toDateTime();
                }
            }
            else if (reader.name() == QLatin1String("colorFilter")) {
                //TODO:
                reader.skipCurrentElement();
            }
            else if (reader.name() == QLatin1String("iconFilter")) {
                //TODO:
                reader.skipCurrentElement();
            }
            else
                reader.skipCurrentElement();
        }
        else if (token == QXmlStreamReader::EndDocument && reader.name() == name)
            break;
    }
}

void AutoFilterColumn::readFilters(QXmlStreamReader &reader)
{
    const auto &name = reader.name();

    while (!reader.atEnd()) {
        auto token = reader.readNext();
        if (token == QXmlStreamReader::StartElement) {
            if (reader.name() == QLatin1String("filter")) {
                QString val = reader.attributes().value(QLatin1String("val")).toString();
                //TODO: parse val and convert it to int, double, date-time etc.
                //Right now filters are stored as strings
                filter.filters.append(val);
            }
            else if (reader.name() == QLatin1String("dateGroupItem")) {
                //TODO: dateGroupItem
                reader.skipCurrentElement();
            }
        }
        else if (token == QXmlStreamReader::EndDocument && reader.name() == name)
            break;
    }
}

void AutoFilterColumn::readCustomFilters(QXmlStreamReader &reader)
{
    const auto &name = reader.name();
    bool op = false;
    parseAttributeBool(reader.attributes(), QLatin1String("and"), op);
    if (op) filter.op = Filter::Operation::And;

    int idx = 0;
    while (!reader.atEnd()) {
        auto token = reader.readNext();
        if (token == QXmlStreamReader::StartElement) {
            if (reader.name() == QLatin1String("customFilter")) {
                auto opStr = reader.attributes().value(QLatin1String("operator"));
                auto op = Filter::Predicate::Equal;
                if (opStr == QLatin1String("lessThan")) op = Filter::Predicate::LessThan;
                else if (opStr == QLatin1String("lessThan")) op = Filter::Predicate::LessThan;
                else if (opStr == QLatin1String("lessThanOrEqual")) op = Filter::Predicate::LessThanOrEqual;
                else if (opStr == QLatin1String("notEqual")) op = Filter::Predicate::NotEqual;
                else if (opStr == QLatin1String("greaterThanOrEqual")) op = Filter::Predicate::GreaterThanOrEqual;
                else if (opStr == QLatin1String("greaterThan")) op = Filter::Predicate::GreaterThan;
                if (idx == 0) {
                    filter.op1 = op;
                    filter.val1 = reader.attributes().value(QLatin1String("val")).toString();
                }
                else {
                    filter.op2 = op;
                    filter.val2 = reader.attributes().value(QLatin1String("val")).toString();
                }
                idx++;
            }
        }
        else if (token == QXmlStreamReader::EndDocument && reader.name() == name)
            break;
    }
}

bool AutoFilterColumn::operator==(const AutoFilterColumn &other) const
{
    if (filterButtonHidden != other.filterButtonHidden) return false;
    if (showFilterOptions != other.showFilterOptions) return false;
    if (filter != other.filter) return false;
    return true;
}

bool AutoFilterColumn::isValid() const
{
    //TODO
    switch (filter.type) {
        case Filter::Type::Invalid: return false;
        case Filter::Type::Values: return !filter.filters.isEmpty();
        case Filter::Type::Custom: return (filter.val1.isValid() || filter.val2.isValid());
        case Filter::Type::Dynamic: return (filter.dynamicType != Filter::DynamicFilterType::Invalid);
        case Filter::Type::Top10: return filter.top10val.has_value();
        default:
            break;
    }

    return true;
}

void SortParameters::write(QXmlStreamWriter &writer, const QString &name) const
{
    if (!isValid()) return;
    //TODO
}

void SortParameters::read(QXmlStreamReader &reader)
{

}

bool SortParameters::operator==(const SortParameters &other) const
{
    //TODO
    return true;
}

bool SortParameters::isValid() const
{
    //TODO
    return true;
}

bool Filter::operator==(const Filter &other) const
{
    if (type != other.type) return false;
    if (filters != other.filters) return false;
    if (val1 != other.val1) return false;
    if (val2 != other.val2) return false;
    if (op1 != other.op1) return false;
    if (op2 != other.op2) return false;
    if (op != other.op) return false;
    if (dynamicType != other.dynamicType) return false;
    if (maxDateTime != other.maxDateTime) return false;
    if (dateTime != other.dateTime) return false;
    if (val != other.val) return false;
    if (top != other.top) return false;
    if (usePercents != other.usePercents) return false;
    if (top10val != other.top10val) return false;
    if (top10filterVal != other.top10filterVal) return false;

    //TODO: icon filter

    //TODO: color filter

    return true;
}

bool Filter::operator!=(const Filter &other) const
{
    return !operator==(other);
}

}
