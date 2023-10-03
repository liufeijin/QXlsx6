// xlsxdatavalidation.cpp

#include <QtGlobal>
#include <QXmlStreamReader>
#include <QXmlStreamWriter>

#include "xlsxdatavalidation.h"
#include "xlsxcellrange.h"

namespace QXlsx {

class DataValidationPrivate : public QSharedData
{
public:
    DataValidationPrivate();
    DataValidationPrivate(const DataValidationPrivate &other);
    ~DataValidationPrivate();
    bool operator==(const DataValidationPrivate &other) const;

    std::optional<DataValidation::Type> validationType;// = DataValidation::Type::None;
    std::optional<DataValidation::Predicate> validationOperator;// = DataValidation::Predicate::Between;
    std::optional<DataValidation::Error> errorStyle;// = DataValidation::Error::Stop;
    std::optional<bool> allowBlank; //=false
    std::optional<bool> isPromptMessageVisible; //=false
    std::optional<bool> isErrorMessageVisible; //=false
    std::optional<bool> isDropDownVisible; //=false
    QString formula1;
    QString formula2;
    QString errorMessage;
    QString errorMessageTitle;
    QString promptMessage;
    QString promptMessageTitle;
    QList<CellRange> ranges;
    std::optional<DataValidation::ImeMode> imeMode;
};

DataValidationPrivate::DataValidationPrivate()
{
}

DataValidationPrivate::DataValidationPrivate(const DataValidationPrivate &other)
    : QSharedData(other)
    , validationType{other.validationType}
    , validationOperator{other.validationOperator}
    , errorStyle{other.errorStyle}
    , allowBlank{other.allowBlank}
    , isPromptMessageVisible{other.isPromptMessageVisible}
    , isErrorMessageVisible{other.isErrorMessageVisible}
    , formula1{other.formula1}
    , formula2{other.formula2}
    , errorMessage{other.errorMessage}
    , errorMessageTitle{other.errorMessageTitle}
    , promptMessage{other.promptMessage}
    , promptMessageTitle{other.promptMessageTitle}
    , ranges{other.ranges}
{
}

DataValidationPrivate::~DataValidationPrivate()
{

}

bool DataValidationPrivate::operator==(const DataValidationPrivate &other) const
{
    if (validationType != other.validationType) return false;
    if (validationOperator != other.validationOperator) return false;
    if (errorStyle != other.errorStyle) return false;
    if (allowBlank != other.allowBlank) return false;
    if (isPromptMessageVisible != other.isPromptMessageVisible) return false;
    if (isErrorMessageVisible != other.isErrorMessageVisible) return false;
    if (formula1 != other.formula1) return false;
    if (formula2 != other.formula2) return false;
    if (errorMessage != other.errorMessage) return false;
    if (errorMessageTitle != other.errorMessageTitle) return false;
    if (promptMessage != other.promptMessage) return false;
    if (promptMessageTitle != other.promptMessageTitle) return false;
    if (ranges != other.ranges) return false;
    return true;
}

DataValidation::DataValidation(Type type,
                               const QString &formula1,
                               std::optional<Predicate> predicate,
                               const QString &formula2)
    : d(new DataValidationPrivate())
{
    d->validationType = type;
    d->formula1 = formula1;
    d->formula2 = formula2;
    d->validationOperator = predicate;
}

DataValidation::DataValidation(const CellRange &allowableValues)
    : d(new DataValidationPrivate())
{
    d->validationType = Type::List;
    d->formula1 = allowableValues.toString(true, true);
}

DataValidation::DataValidation(const QTime &time1, std::optional<Predicate> predicate, const QTime &time2)
 : d(new DataValidationPrivate())
{
    d->validationType = Type::Time;
    d->formula1 = QString::number(timeToNumber(time1));
    d->validationOperator = predicate;
    if (time2.isValid()) d->formula2 = timeToNumber(time2);
}

DataValidation::DataValidation(const QDate &date1, std::optional<Predicate> predicate, const QDate &date2)
: d(new DataValidationPrivate())
{
    d->validationType = Type::Date;
    d->formula1 = dateToNumber(date1);
    d->validationOperator = predicate;
    if (date2.isValid()) d->formula2 = dateToNumber(date2);
}

DataValidation::DataValidation(int len1, std::optional<Predicate> predicate, std::optional<int> len2)
: d(new DataValidationPrivate())
{
    d->validationType = Type::TextLength;
    if (len1 >= 0)
        d->formula1 = QString::number(len1);
    d->validationOperator = predicate;
    if (len2.has_value() && len2.value() >= 0) d->formula2 = QString::number(len2.value());
}

DataValidation::DataValidation()
{

}

DataValidation::DataValidation(const DataValidation &other)
    : d(other.d)
{

}

DataValidation &DataValidation::operator=(const DataValidation &other)
{
    if (*this != other)
        this->d = other.d;
    return *this;
}

bool DataValidation::operator==(const DataValidation &other) const
{
    if (d == other.d) return true;
    if (!d || !other.d) return false;
    return *this->d.constData() == *other.d.constData();
}

bool DataValidation::operator!=(const DataValidation &other) const
{
    return !operator==(other);
}

DataValidation::operator QVariant() const
{
    const auto& cref
#if QT_VERSION >= 0x060000 // Qt 6.0 or over
        = QMetaType::fromType<DataValidation>();
#else
        = qMetaTypeId<DataValidation>() ;
#endif
    return QVariant(cref, this);
}

bool DataValidation::isValid() const
{
    if (!d) return false;
    return !d->ranges.isEmpty();
}


DataValidation::~DataValidation()
{
}

std::optional<DataValidation::Type> DataValidation::type() const
{
    if (d) return d->validationType;
    return Type::None;
}

std::optional<DataValidation::Predicate> DataValidation::predicate() const
{
    if (d) return d->validationOperator;
    return Predicate::Between; //TODO: check default value
}

std::optional<DataValidation::Error> DataValidation::errorStyle() const
{
    if (d) return d->errorStyle;
    return Error::Stop;
}

QString DataValidation::formula1() const
{
    if (d) return d->formula1;
    return {};
}

QString DataValidation::formula2() const
{
    if (d) return d->formula2;
    return {};
}

std::optional<bool> DataValidation::allowBlank() const
{
    if (d) return d->allowBlank;
    return false;
}

QString DataValidation::errorMessage() const
{
    if (d) return d->errorMessage;
    return {};
}

QString DataValidation::errorTitle() const
{
    if (d) return d->errorMessageTitle;
    return {};
}

QString DataValidation::promptMessage() const
{
    if (d) return d->promptMessage;
    return {};
}

QString DataValidation::promptTitle() const
{
    if (d) return d->promptMessageTitle;
    return {};
}

std::optional<bool> DataValidation::isPromptMessageVisible() const
{
    if (d) return d->isPromptMessageVisible;
    return false;
}

std::optional<bool> DataValidation::isErrorMessageVisible() const
{
    if (d) return d->isErrorMessageVisible;
    return false;
}

QList<CellRange> DataValidation::ranges() const
{
    if (d) return d->ranges;
    return {};
}

void DataValidation::setType(Type type)
{
    if (!d) d = new DataValidationPrivate();
    d->validationType = type;
}

void DataValidation::setPredicate(Predicate op)
{
    if (!d) d = new DataValidationPrivate();
    d->validationOperator = op;
}

void DataValidation::setErrorStyle(DataValidation::Error errorStyle)
{
    if (!d) d = new DataValidationPrivate();
    d->errorStyle = errorStyle;
}

void DataValidation::setFormula1(const QString &formula)
{
    if (!d) d = new DataValidationPrivate();
    if (formula.startsWith(QLatin1Char('=')))
        d->formula1 = formula.mid(1);
    else
        d->formula1 = formula;
}

void DataValidation::setFormula2(const QString &formula)
{
    if (!d) d = new DataValidationPrivate();
    if (formula.startsWith(QLatin1Char('=')))
        d->formula2 = formula.mid(1);
    else
        d->formula2 = formula;
}

void DataValidation::setErrorMessage(const QString &error, const QString &title)
{
    if (!d) d = new DataValidationPrivate();
    d->errorMessage = error;
    d->errorMessageTitle = title;
}

void DataValidation::setPromptMessage(const QString &prompt, const QString &title)
{
    if (!d) d = new DataValidationPrivate();
    d->promptMessage = prompt;
    d->promptMessageTitle = title;
}

std::optional<bool> DataValidation::isDropDownVisible() const
{
    if (d) return d->isDropDownVisible;
    return false;
}

void DataValidation::setDropDownVisible(bool visible)
{
    if (!d) d = new DataValidationPrivate();
    d->isDropDownVisible = visible;
}

std::optional<DataValidation::ImeMode> DataValidation::imeMode() const
{
    if (d) return d->imeMode;
    return ImeMode::NoControl;
}

void DataValidation::setImeMode(ImeMode mode)
{
    if (!d) d = new DataValidationPrivate();
    d->imeMode = mode;
}

void DataValidation::setAllowBlank(bool enable)
{
    if (!d) d = new DataValidationPrivate();
    d->allowBlank = enable;
}

void DataValidation::setPromptMessageVisible(bool visible)
{
    if (!d) d = new DataValidationPrivate();
    d->isPromptMessageVisible = visible;
}

void DataValidation::setErrorMessageVisible(bool visible)
{
    if (!d) d = new DataValidationPrivate();
    d->isErrorMessageVisible = visible;
}

bool DataValidation::addCell(const CellReference &cell)
{
    return addRange(CellRange(cell, cell));
}

bool DataValidation::addCell(int row, int column)
{
    return addRange(CellRange(row, column, row, column));
}

bool DataValidation::addRange(int firstRow, int firstColumn, int lastRow, int lastColumn)
{
    return addRange(CellRange(firstRow, firstColumn, lastRow, lastColumn));
}

bool DataValidation::removeRange(const CellRange &range)
{
    if (!d) return false;
    return d->ranges.removeOne(range);
}

bool DataValidation::removeRange(int firstRow, int firstColumn, int lastRow, int lastColumn)
{
    return removeRange(CellRange(firstRow, firstColumn, lastRow, lastColumn));
}

bool DataValidation::addRange(const CellRange &range)
{
    if (!range.isValid()) return false;
    if (!d) d = new DataValidationPrivate();
    if (d->ranges.contains(range)) return false;
    d->ranges.append(range);
    return true;
}

/*!
 * \internal
 */
void DataValidation::write(QXmlStreamWriter &writer) const
{
    //TODO: check writing empty dataValidation
    if (!d) return;

    writer.writeStartElement(QLatin1String("dataValidation"));
    if (d->validationType.has_value())
        writer.writeAttribute(QLatin1String("type"), toString(d->validationType.value()));
    if (d->errorStyle.has_value())
        writer.writeAttribute(QLatin1String("errorStyle"), toString(d->errorStyle.value()));
    if (d->imeMode.has_value())
        writeAttribute(writer, QLatin1String("imeMode"), toString(d->imeMode.value()));
    if (d->validationOperator.has_value())
        writer.writeAttribute(QLatin1String("operator"), toString(d->validationOperator.value()));
    writeAttribute(writer, QLatin1String("allowBlank"), d->allowBlank);
    writeAttribute(writer, QLatin1String("showDropDown"), d->isDropDownVisible);
    writeAttribute(writer, QLatin1String("showInputMessage"), d->isPromptMessageVisible);
    writeAttribute(writer, QLatin1String("showErrorMessage"), d->isErrorMessageVisible);
    writeAttribute(writer, QLatin1String("errorTitle"), d->errorMessageTitle);
    writeAttribute(writer, QLatin1String("error"), d->errorMessage);
    writeAttribute(writer, QLatin1String("promptTitle"), d->promptMessageTitle);
    writeAttribute(writer, QLatin1String("prompt"), d->promptMessage);

    QStringList sqref;
    for (const CellRange &range : qAsConst(d->ranges))
        sqref.append(range.toString());
    writeAttribute(writer, QLatin1String("sqref"), sqref.join(QLatin1String(" ")));

    if (!formula1().isEmpty())
        writer.writeTextElement(QLatin1String("formula1"), formula1());
    if (!formula2().isEmpty())
        writer.writeTextElement(QLatin1String("formula2"), formula2());

    writer.writeEndElement(); //dataValidation
}

/*!
 * \internal
 */
void DataValidation::read(QXmlStreamReader &reader)
{
    Q_ASSERT(reader.name() == QLatin1String("dataValidation"));
    if (!d) d = new DataValidationPrivate();

    const auto &a = reader.attributes();

    const auto sqrefParts = a.value(QLatin1String("sqref")).toString().split(QLatin1Char(' '));
    for (const QString &range : sqrefParts)
        addRange(range);

    if (a.hasAttribute(QLatin1String("type"))) {
        Type t;
        fromString(a.value(QLatin1String("type")).toString(), t);
        d->validationType = t;
    }
    if (a.hasAttribute(QLatin1String("errorStyle"))) {
        Error t;
        fromString(a.value(QLatin1String("errorStyle")).toString(), t);
        d->errorStyle = t;
    }
    if (a.hasAttribute(QLatin1String("imeMode"))) {
        ImeMode im;
        QString es = a.value(QLatin1String("imeMode")).toString();
        fromString(es, im); d->imeMode = im;
    }
    if (a.hasAttribute(QLatin1String("operator"))) {
        Predicate t;
        fromString(a.value(QLatin1String("operator")).toString(), t);
        d->validationOperator = t;
    }
    parseAttributeBool(a, QLatin1String("allowBlank"), d->allowBlank);
    parseAttributeBool(a, QLatin1String("showInputMessage"), d->isPromptMessageVisible);
    parseAttributeBool(a, QLatin1String("showErrorMessage"), d->isErrorMessageVisible);
    parseAttributeString(a, QLatin1String("errorTitle"), d->errorMessageTitle);
    parseAttributeString(a, QLatin1String("error"), d->errorMessage);
    parseAttributeString(a, QLatin1String("promptTitle"), d->promptMessageTitle);
    parseAttributeString(a, QLatin1String("prompt"), d->promptMessage);

    //find the end
    while(!(reader.name() == QLatin1String("dataValidation") && reader.tokenType() == QXmlStreamReader::EndElement)) {
        reader.readNextStartElement();
        if (reader.tokenType() == QXmlStreamReader::StartElement) {
            if (reader.name() == QLatin1String("formula1")) {
                setFormula1(reader.readElementText());
            } else if (reader.name() == QLatin1String("formula2")) {
                setFormula2(reader.readElementText());
            }
        }
    }
}

}
