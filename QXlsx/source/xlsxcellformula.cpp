// xlsxcellformula.cpp

#include <QtGlobal>
#include <QObject>
#include <QString>
#include <QXmlStreamReader>
#include <QXmlStreamWriter>
#include <QDebug>

#include "xlsxcellformula.h"
#include "xlsxutility_p.h"
#include "xlsxcellrange.h"

namespace QXlsx {

class CellFormulaPrivate : public QSharedData
{
public:
    CellFormulaPrivate() {}
    CellFormulaPrivate(const QString &formula, const CellRange &reference, CellFormula::Type type);
    CellFormulaPrivate(const CellFormulaPrivate &other);
    ~CellFormulaPrivate();
    bool operator == (const CellFormulaPrivate &other) const;

    QString formula; //formula contents
    std::optional<CellFormula::Type> type;
    CellRange reference;
    std::optional<bool> ca; //Calculate Cell
    std::optional<int> si;  //Shared group index
};

CellFormulaPrivate::CellFormulaPrivate(const QString &formula_, const CellRange &ref_, CellFormula::Type type_)
    :formula(formula_), type{type_}, reference(ref_)
{
    //Remove the formula '=' sign if exists
    if (formula.startsWith(QLatin1String("=")))
        formula.remove(0,1);
    else if (formula.startsWith(QLatin1String("{=")) && formula.endsWith(QLatin1String("}")))
        formula = formula.mid(2, formula.length()-3);
}

CellFormulaPrivate::CellFormulaPrivate(const CellFormulaPrivate &other)
    : QSharedData(other)
    , formula(other.formula), type(other.type), reference(other.reference)
    , ca(other.ca), si(other.si)
{

}

CellFormulaPrivate::~CellFormulaPrivate()
{

}

bool CellFormulaPrivate::operator ==(const CellFormulaPrivate &other) const
{
    if (formula != other.formula) return false;
    if (type != other.type) return false;
    if (reference != other.reference) return false;
    if (ca != other.ca) return false;
    if (si != other.si) return false;
    return true;
}

CellFormula::CellFormula()
{

}

CellFormula::CellFormula(const char *formula, Type type)
    :d(new CellFormulaPrivate(QString::fromLatin1(formula), CellRange(), type))
{

}

CellFormula::CellFormula(const QString &formula, Type type)
    :d(new CellFormulaPrivate(formula, CellRange(), type))
{

}

CellFormula::CellFormula(const QString &formula, const CellRange &range, Type type)
    :d(new CellFormulaPrivate(formula, range, type))
{

}

CellFormula::CellFormula(const CellFormula &other)
    :d(other.d)
{
}

CellFormula &CellFormula::operator =(const CellFormula &other)
{
    if (*this != other)
        d = other.d;
    return *this;
}

CellFormula::~CellFormula()
{

}

std::optional<CellFormula::Type> CellFormula::type() const
{
    return d ? d->type : std::nullopt;
}

QString CellFormula::text() const
{
    return d ? d->formula : QString();
}

CellRange CellFormula::reference() const
{
    return d ? d->reference : CellRange();
}

void CellFormula::setReference(const CellRange &range)
{
    if (!d) d = new CellFormulaPrivate;
    d->reference = range;
}

bool CellFormula::isValid() const
{
    return d;
}

std::optional<int> CellFormula::sharedIndex() const
{
    return d && d->type.value_or(Type::Normal) == Type::Shared ? d->si : std::nullopt;
}

void CellFormula::setSharedIndex(int sharedIndex)
{
    if (!d) d = new CellFormulaPrivate;
    d->si = sharedIndex;
}

std::optional<bool> CellFormula::needsRecalculation() const
{
    if (d) return d->ca;
    return {};
}

void CellFormula::setNeedsRecalculation(bool recalculate)
{
    if (!d) d = new CellFormulaPrivate;
    d->ca = recalculate;
}

bool CellFormula::saveToXml(QXmlStreamWriter &writer) const
{
    if (!d) return false;
    QString t;
    switch (d->type.value_or(Type::Normal))
    {
    case CellFormula::Type::Array:
        t = QStringLiteral("array");
        break;
    case CellFormula::Type::Shared:
        t = QStringLiteral("shared");
        break;
    case CellFormula::Type::Normal:
        t = QStringLiteral("normal");
        break;
    case CellFormula::Type::DataTable:
        t = QStringLiteral("dataTable");
        break;
    default: // undefined type
        return false;
        break;
    }

    writer.writeStartElement(QStringLiteral("f"));

    if (!t.isEmpty())
        writer.writeAttribute(QStringLiteral("t"), t); // write type(t)

    if ( d->type.value_or(Type::Normal) == CellFormula::Type::Shared ||
         d->type.value_or(Type::Normal) == CellFormula::Type::Array ||
         d->type.value_or(Type::Normal) == CellFormula::Type::DataTable ) {
        if (d->reference.isValid())
            writer.writeAttribute(QStringLiteral("ref"), d->reference.toString());
    }

    writeAttribute(writer, QLatin1String("ca"), d->ca);

    if (d->type.value_or(Type::Normal) == CellFormula::Type::Shared)
        writeAttribute(writer, QLatin1String("si"), d->si);

    if (!d->formula.isEmpty()) {
        QString strFormula = d->formula;
        writer.writeCharacters(strFormula); // write formula
    }

    writer.writeEndElement(); // f

    return true;
}

bool CellFormula::loadFromXml(QXmlStreamReader &reader)
{
    Q_ASSERT(reader.name() == QLatin1String("f"));
    if (!d)
        d = new CellFormulaPrivate(QString(), CellRange(), Type::Normal);

    QXmlStreamAttributes attributes = reader.attributes();
    QString typeString = attributes.value(QLatin1String("t")).toString();

    if (typeString == QLatin1String("array")) d->type = Type::Array;
    else if (typeString == QLatin1String("shared")) d->type = Type::Shared;
    else if (typeString == QLatin1String("normal")) d->type = Type::Normal;
    else if (typeString == QLatin1String("dataTable")) d->type = Type::DataTable;

    if (attributes.hasAttribute(QLatin1String("ref")))
        d->reference = CellRange(attributes.value(QLatin1String("ref")).toString());

    parseAttributeBool(attributes, QLatin1String("ca"), d->ca);
    parseAttributeInt(attributes, QLatin1String("si"), d->si);

    d->formula = reader.readElementText(); // read formula

    return true;
}

bool CellFormula::operator ==(const CellFormula &formula) const
{
    if (d == formula.d) return true;
    if (!d || !formula.d) return false;
    return *this->d.constData() == *formula.d.constData();
}

bool CellFormula::operator !=(const CellFormula &formula) const
{
    return !operator==(formula);
}

QXlsx::CellFormula::operator QVariant() const
{
    const auto& cref
#if QT_VERSION >= 0x060000 // Qt 6.0 or over
        = QMetaType::fromType<CellFormula>();
#else
        = qMetaTypeId<CellFormula>() ;
#endif
    return QVariant(cref, this);
}

}
