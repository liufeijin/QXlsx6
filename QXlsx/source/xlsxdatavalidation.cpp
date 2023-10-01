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
    DataValidationPrivate(DataValidation::Type type, DataValidation::Operator op,
                          const QString &formula1, const QString &formula2, bool allowBlank);
    DataValidationPrivate(const DataValidationPrivate &other);
    ~DataValidationPrivate();
    bool operator==(const DataValidationPrivate &other) const;

    DataValidation::Type validationType = DataValidation::Type::None;
    DataValidation::Operator validationOperator = DataValidation::Operator::Between;
    DataValidation::Error errorStyle = DataValidation::Error::Stop;
    bool allowBlank = false;
    bool isPromptMessageVisible = true;
    bool isErrorMessageVisible = true;
    QString formula1;
    QString formula2;
    QString errorMessage;
    QString errorMessageTitle;
    QString promptMessage;
    QString promptMessageTitle;
    QList<CellRange> ranges;
};

DataValidationPrivate::DataValidationPrivate()
{
}

DataValidationPrivate::DataValidationPrivate(DataValidation::Type type,
                                             DataValidation::Operator op,
                                             const QString &formula1,
                                             const QString &formula2,
                                             bool allowBlank)
    : validationType(type)
    , validationOperator(op)
    , allowBlank(allowBlank)
    , formula1(formula1)
    , formula2(formula2)
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
    if ( allowBlank != other.allowBlank) return false;
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

DataValidation::DataValidation(Type type, Operator op, const QString &formula1, const QString &formula2, bool allowBlank)
    : d(new DataValidationPrivate(type, op, formula1, formula2, allowBlank))
{
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

bool DataValidation::isValid() const
{
    if (!d) return false;
    return true;
}


DataValidation::~DataValidation()
{
}

/*!
    Returns the validation type.
 */
DataValidation::Type DataValidation::type() const
{
    if (d) return d->validationType;
    return Type::None;
}

/*!
    Returns the validation operator.
 */
DataValidation::Operator DataValidation::validationOperator() const
{
    if (d) return d->validationOperator;
    return Operator::Between; //TODO: check default value
}

DataValidation::Error DataValidation::errorStyle() const
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

bool DataValidation::allowBlank() const
{
    if (d) return d->allowBlank;
    return false; //TODO: check default value
}

QString DataValidation::errorMessage() const
{
    if (d) return d->errorMessage;
    return {};
}

QString DataValidation::errorMessageTitle() const
{
    if (d) return d->errorMessageTitle;
    return {};
}

QString DataValidation::promptMessage() const
{
    if (d) return d->promptMessage;
    return {};
}

QString DataValidation::promptMessageTitle() const
{
    if (d) return d->promptMessageTitle;
    return {};
}

bool DataValidation::isPromptMessageVisible() const
{
    if (d) return d->isPromptMessageVisible;
    return true; //TODO: check default value
}

bool DataValidation::isErrorMessageVisible() const
{
    if (d) return d->isErrorMessageVisible;
    return true;
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

void DataValidation::setValidationOperator(Operator op)
{
    if (!d) d = new DataValidationPrivate();
    d->validationOperator = op;
}

void DataValidation::setErrorStyle(DataValidation::Error es)
{
    if (!d) d = new DataValidationPrivate();
    d->errorStyle = es;
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

void DataValidation::addCell(const CellReference &cell)
{
    if (!d) d = new DataValidationPrivate();
    d->ranges.append(CellRange(cell, cell));
}

void DataValidation::addCell(int row, int col)
{
    if (!d) d = new DataValidationPrivate();
    d->ranges.append(CellRange(row, col, row, col));
}

void DataValidation::addRange(int firstRow, int firstCol, int lastRow, int lastCol)
{
    if (!d) d = new DataValidationPrivate();
    d->ranges.append(CellRange(firstRow, firstCol, lastRow, lastCol));
}

void DataValidation::addRange(const CellRange &range)
{
    if (!d) d = new DataValidationPrivate();
    d->ranges.append(range);
}

/*!
 * \internal
 */
void DataValidation::write(QXmlStreamWriter &writer) const
{
    //TODO: check writing empty dataValidation
    if (!d) return;

    writer.writeStartElement(QLatin1String("dataValidation"));
    if (d->validationType != Type::None)
        writer.writeAttribute(QLatin1String("type"), toString(d->validationType));
    if (d->errorStyle != Error::Stop)
        writer.writeAttribute(QLatin1String("errorStyle"), toString(d->errorStyle));
    if (d->validationOperator != DataValidation::Operator::Between)
        writer.writeAttribute(QLatin1String("operator"), toString(d->validationOperator));
    if (d->allowBlank)
        writeAttribute(writer, QLatin1String("allowBlank"), d->allowBlank);
    //        if (dropDownVisible())
    //            writer.writeAttribute(QStringLiteral("showDropDown"), QStringLiteral("1"));
    if (d->isPromptMessageVisible)
        writeAttribute(writer, QLatin1String("showInputMessage"), d->isPromptMessageVisible);
    if (d->isErrorMessageVisible)
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
        QString t = a.value(QLatin1String("type")).toString();
        fromString(t, d->validationType);
    }
    if (a.hasAttribute(QLatin1String("errorStyle"))) {
        QString es = a.value(QLatin1String("errorStyle")).toString();
        fromString(es, d->errorStyle);
    }
    if (a.hasAttribute(QLatin1String("operator"))) {
        QString op = a.value(QLatin1String("operator")).toString();
        fromString(op, d->validationOperator);
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
