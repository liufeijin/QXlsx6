// xlsxconditionalformatting.cpp

#include <QtGlobal>
#include <QXmlStreamReader>
#include <QXmlStreamWriter>
#include <QDebug>

#include "xlsxconditionalformatting.h"
#include "xlsxconditionalformatting_p.h"
#include "xlsxworksheet.h"
#include "xlsxcellrange.h"
#include "xlsxstyles_p.h"

namespace QXlsx {

ConditionalFormattingPrivate::ConditionalFormattingPrivate()
{

}

ConditionalFormattingPrivate::ConditionalFormattingPrivate(const ConditionalFormattingPrivate &other)
    :QSharedData(other)
{

}

ConditionalFormattingPrivate::~ConditionalFormattingPrivate()
{

}

void ConditionalFormattingPrivate::writeCfVo(QXmlStreamWriter &writer, const XlsxCfVoData &cfvo) const
{
    writer.writeEmptyElement(QStringLiteral("cfvo"));
    QString type;
    switch(cfvo.type) {
    case ConditionalFormatting::ValueObjectType::Formula: type=QStringLiteral("formula"); break;
    case ConditionalFormatting::ValueObjectType::Max: type=QStringLiteral("max"); break;
    case ConditionalFormatting::ValueObjectType::Min: type=QStringLiteral("min"); break;
    case ConditionalFormatting::ValueObjectType::Num: type=QStringLiteral("num"); break;
    case ConditionalFormatting::ValueObjectType::Percent: type=QStringLiteral("percent"); break;
    case ConditionalFormatting::ValueObjectType::Percentile: type=QStringLiteral("percentile"); break;
    default: break;
    }
    writer.writeAttribute(QStringLiteral("type"), type);
    writer.writeAttribute(QStringLiteral("val"), cfvo.value);
    if (!cfvo.gte)
        writer.writeAttribute(QStringLiteral("gte"), QStringLiteral("0"));
}

ConditionalFormatting::ConditionalFormatting()
    :d(new ConditionalFormattingPrivate())
{

}

ConditionalFormatting::ConditionalFormatting(const ConditionalFormatting &other)
    :d(other.d)
{

}

ConditionalFormatting &ConditionalFormatting::operator=(const ConditionalFormatting &other)
{
    this->d = other.d;
    return *this;
}


ConditionalFormatting::~ConditionalFormatting()
{
}

bool ConditionalFormatting::addHighlightCellsRule(Type type, const QString &formula1, const QString &formula2, const Format &format, bool stopIfTrue)
{
    if (format.isEmpty())
        return false;

    bool skipFormula = false;

    auto cfRule = std::make_shared<XlsxCfRuleData>();
    switch (type) {
    case Type::LessThan:
        cfRule->attrs[XlsxCfRuleData::A_type] = QStringLiteral("cellIs");
        cfRule->attrs[XlsxCfRuleData::A_operator] = QStringLiteral("lessThan");
        break;
    case Type::LessThanOrEqual:
        cfRule->attrs[XlsxCfRuleData::A_type] = QStringLiteral("cellIs");
        cfRule->attrs[XlsxCfRuleData::A_operator] = QStringLiteral("lessThanOrEqual");
        break;
    case Type::Equal:
        cfRule->attrs[XlsxCfRuleData::A_type] = QStringLiteral("cellIs");
        cfRule->attrs[XlsxCfRuleData::A_operator] = QStringLiteral("equal");
        break;
    case Type::NotEqual:
        cfRule->attrs[XlsxCfRuleData::A_type] = QStringLiteral("cellIs");
        cfRule->attrs[XlsxCfRuleData::A_operator] = QStringLiteral("notEqual");
        break;
    case Type::GreaterThanOrEqual:
        cfRule->attrs[XlsxCfRuleData::A_type] = QStringLiteral("cellIs");
        cfRule->attrs[XlsxCfRuleData::A_operator] = QStringLiteral("greaterThanOrEqual");
        break;
    case Type::GreaterThan:
        cfRule->attrs[XlsxCfRuleData::A_type] = QStringLiteral("cellIs");
        cfRule->attrs[XlsxCfRuleData::A_operator] = QStringLiteral("greaterThan");
        break;
    case Type::Between:
        cfRule->attrs[XlsxCfRuleData::A_type] = QStringLiteral("cellIs");
        cfRule->attrs[XlsxCfRuleData::A_operator] = QStringLiteral("between");
        break;
    case Type::NotBetween:
        cfRule->attrs[XlsxCfRuleData::A_type] = QStringLiteral("cellIs");
        cfRule->attrs[XlsxCfRuleData::A_operator] = QStringLiteral("notBetween");
        break;
    case Type::ContainsText:
        cfRule->attrs[XlsxCfRuleData::A_type] = QStringLiteral("containsText");
        cfRule->attrs[XlsxCfRuleData::A_operator] = QStringLiteral("containsText");
        //<formula>NOT(ISERROR(SEARCH("a",B6)))</formula>
        cfRule->attrs[XlsxCfRuleData::A_formula1_temp] = QStringLiteral("NOT(ISERROR(SEARCH(\"%1\",%2)))").arg(formula1);
        cfRule->attrs[XlsxCfRuleData::A_text] = formula1;
        skipFormula = true;
        break;
    case Type::NotContainsText:
        cfRule->attrs[XlsxCfRuleData::A_type] = QStringLiteral("notContainsText");
        cfRule->attrs[XlsxCfRuleData::A_operator] = QStringLiteral("notContains");
        cfRule->attrs[XlsxCfRuleData::A_formula1_temp] = QStringLiteral("ISERROR(SEARCH(\"%2\",%1))").arg(formula1);
        cfRule->attrs[XlsxCfRuleData::A_text] = formula1;
        skipFormula = true;
        break;
    case Type::BeginsWith:
        cfRule->attrs[XlsxCfRuleData::A_type] = QStringLiteral("beginsWith");
        cfRule->attrs[XlsxCfRuleData::A_operator] = QStringLiteral("beginsWith");
        cfRule->attrs[XlsxCfRuleData::A_formula1_temp] = QStringLiteral("LEFT(%2,LEN(\"%1\"))=\"%1\"").arg(formula1);
        cfRule->attrs[XlsxCfRuleData::A_text] = formula1;
        skipFormula = true;
        break;
    case Type::EndsWith:
        cfRule->attrs[XlsxCfRuleData::A_type] = QStringLiteral("endsWith");
        cfRule->attrs[XlsxCfRuleData::A_operator] = QStringLiteral("endsWith");
        cfRule->attrs[XlsxCfRuleData::A_formula1_temp] = QStringLiteral("RIGHT(%2,LEN(\"%1\"))=\"%1\"").arg(formula1);
        cfRule->attrs[XlsxCfRuleData::A_text] = formula1;
        skipFormula = true;
        break;
    case Type::TimePeriod:
        cfRule->attrs[XlsxCfRuleData::A_type] = QStringLiteral("timePeriod");
        //Todo
        return false;
        break;
    case Type::Duplicate:
        cfRule->attrs[XlsxCfRuleData::A_type] = QStringLiteral("duplicateValues");
        skipFormula = true;
        break;
    case Type::Unique:
        cfRule->attrs[XlsxCfRuleData::A_type] = QStringLiteral("uniqueValues");
        skipFormula = true;
        break;
    case Type::Blanks:
        cfRule->attrs[XlsxCfRuleData::A_type] = QStringLiteral("containsBlanks");
        cfRule->attrs[XlsxCfRuleData::A_formula1_temp] = QStringLiteral("LEN(TRIM(%1))=0");
        skipFormula = true;
        break;
    case Type::NoBlanks:
        cfRule->attrs[XlsxCfRuleData::A_type] = QStringLiteral("notContainsBlanks");
        cfRule->attrs[XlsxCfRuleData::A_formula1_temp] = QStringLiteral("LEN(TRIM(%1))>0");
        skipFormula = true;
        break;
    case Type::Errors:
        cfRule->attrs[XlsxCfRuleData::A_type] = QStringLiteral("containsErrors");
        cfRule->attrs[XlsxCfRuleData::A_formula1_temp] = QStringLiteral("ISERROR(%1)");
        skipFormula = true;
        break;
    case Type::NoErrors:
        cfRule->attrs[XlsxCfRuleData::A_type] = QStringLiteral("notContainsErrors");
        cfRule->attrs[XlsxCfRuleData::A_formula1_temp] = QStringLiteral("NOT(ISERROR(%1))");
        skipFormula = true;
        break;
    case Type::Top:
        cfRule->attrs[XlsxCfRuleData::A_type] = QStringLiteral("top10");
        cfRule->attrs[XlsxCfRuleData::A_rank] = !formula1.isEmpty() ? formula1 : QStringLiteral("10");
        skipFormula = true;
        break;
    case Type::TopPercent:
        cfRule->attrs[XlsxCfRuleData::A_type] = QStringLiteral("top10");
        cfRule->attrs[XlsxCfRuleData::A_percent] = QStringLiteral("1");
        cfRule->attrs[XlsxCfRuleData::A_rank] = !formula1.isEmpty() ? formula1 : QStringLiteral("10");
        skipFormula = true;
        break;
    case Type::Bottom:
        cfRule->attrs[XlsxCfRuleData::A_type] = QStringLiteral("top10");
        cfRule->attrs[XlsxCfRuleData::A_bottom] = QStringLiteral("1");
        cfRule->attrs[XlsxCfRuleData::A_rank] = !formula1.isEmpty() ? formula1 : QStringLiteral("10");
        skipFormula = true;
        break;
    case Type::BottomPercent:
        cfRule->attrs[XlsxCfRuleData::A_type] = QStringLiteral("top10");
        cfRule->attrs[XlsxCfRuleData::A_bottom] = QStringLiteral("1");
        cfRule->attrs[XlsxCfRuleData::A_percent] = QStringLiteral("1");
        cfRule->attrs[XlsxCfRuleData::A_rank] = !formula1.isEmpty() ? formula1 : QStringLiteral("10");
        skipFormula = true;
        break;
    case Type::AboveAverage:
        cfRule->attrs[XlsxCfRuleData::A_type] = QStringLiteral("aboveAverage");
        skipFormula = true;
        break;
    case Type::AboveOrEqualAverage:
        cfRule->attrs[XlsxCfRuleData::A_type] = QStringLiteral("aboveAverage");
        cfRule->attrs[XlsxCfRuleData::A_equalAverage] = QStringLiteral("1");
        skipFormula = true;
        break;
    case Type::AboveStdDev1:
        cfRule->attrs[XlsxCfRuleData::A_type] = QStringLiteral("aboveAverage");
        cfRule->attrs[XlsxCfRuleData::A_stdDev] = QStringLiteral("1");
        break;
    case Type::AboveStdDev2:
        cfRule->attrs[XlsxCfRuleData::A_type] = QStringLiteral("aboveAverage");
        cfRule->attrs[XlsxCfRuleData::A_stdDev] = QStringLiteral("2");
        break;
    case Type::AboveStdDev3:
        cfRule->attrs[XlsxCfRuleData::A_type] = QStringLiteral("aboveAverage");
        cfRule->attrs[XlsxCfRuleData::A_stdDev] = QStringLiteral("3");
        break;
    case Type::BelowAverage:
        cfRule->attrs[XlsxCfRuleData::A_type] = QStringLiteral("aboveAverage");
        cfRule->attrs[XlsxCfRuleData::A_aboveAverage] = QStringLiteral("0");
        break;
    case Type::BelowOrEqualAverage:
        cfRule->attrs[XlsxCfRuleData::A_type] = QStringLiteral("aboveAverage");
        cfRule->attrs[XlsxCfRuleData::A_aboveAverage] = QStringLiteral("0");
        cfRule->attrs[XlsxCfRuleData::A_equalAverage] = QStringLiteral("1");
        break;
    case Type::BelowStdDev1:
        cfRule->attrs[XlsxCfRuleData::A_type] = QStringLiteral("aboveAverage");
        cfRule->attrs[XlsxCfRuleData::A_aboveAverage] = QStringLiteral("0");
        cfRule->attrs[XlsxCfRuleData::A_stdDev] = QStringLiteral("1");
        break;
    case Type::BelowStdDev2:
        cfRule->attrs[XlsxCfRuleData::A_type] = QStringLiteral("aboveAverage");
        cfRule->attrs[XlsxCfRuleData::A_aboveAverage] = QStringLiteral("0");
        cfRule->attrs[XlsxCfRuleData::A_stdDev] = QStringLiteral("2");
        break;
    case Type::BelowStdDev3:
        cfRule->attrs[XlsxCfRuleData::A_type] = QStringLiteral("aboveAverage");
        cfRule->attrs[XlsxCfRuleData::A_aboveAverage] = QStringLiteral("0");
        cfRule->attrs[XlsxCfRuleData::A_stdDev] = QStringLiteral("3");
        break;
    case Type::Expression:
        cfRule->attrs[XlsxCfRuleData::A_type] = QStringLiteral("expression");
        break;
    }

    cfRule->dxfFormat = format;
    if (stopIfTrue)
        cfRule->attrs[XlsxCfRuleData::A_stopIfTrue] = true;
    if (!skipFormula) {
        if (!formula1.isEmpty())
            cfRule->attrs[XlsxCfRuleData::A_formula1] = formula1.startsWith(QLatin1String("=")) ? formula1.mid(1) : formula1;
        if (!formula2.isEmpty())
            cfRule->attrs[XlsxCfRuleData::A_formula2] = formula2.startsWith(QLatin1String("=")) ? formula2.mid(1) : formula2;
    }
    d->appendRule(cfRule);
    return true;
}

bool ConditionalFormatting::addHighlightCellsRule(Type type, const QString &formula1, const QString &formula2, PredefinedFormat format, bool stopIfTrue)
{
    return addHighlightCellsRule(type, formula1, formula2, predefinedFormat(format), stopIfTrue);
}

bool ConditionalFormatting::addHighlightCellsRule(Type type, const Format &format, bool stopIfTrue)
{
    return addHighlightCellsRule(type, QString(), QString(), format, stopIfTrue);
}

bool ConditionalFormatting::addHighlightCellsRule(Type type, PredefinedFormat format, bool stopIfTrue)
{
    return addHighlightCellsRule(type, QString(), QString(), predefinedFormat(format), stopIfTrue);
}

bool ConditionalFormatting::addHighlightCellsRule(Type type, const QString &formula, const Format &format, bool stopIfTrue)
{
    return addHighlightCellsRule(type, formula, QString(), format, stopIfTrue);
}

bool ConditionalFormatting::addHighlightCellsRule(Type type, const QString &formula, PredefinedFormat format, bool stopIfTrue)
{
    return addHighlightCellsRule(type, formula, QString(), predefinedFormat(format), stopIfTrue);
}

bool ConditionalFormatting::addDataBarRule(const QColor &color, ValueObjectType type1, const QString &val1, ValueObjectType type2, const QString &val2, bool showData, bool stopIfTrue)
{
    auto cfRule = std::make_shared<XlsxCfRuleData>();

    cfRule->attrs[XlsxCfRuleData::A_type] = QStringLiteral("dataBar");
    cfRule->attrs[XlsxCfRuleData::A_color1] = Color(color);
    if (stopIfTrue)
        cfRule->attrs[XlsxCfRuleData::A_stopIfTrue] = true;
    if (!showData)
        cfRule->attrs[XlsxCfRuleData::A_hideData] = true;

    XlsxCfVoData cfvo1(type1, val1);
    XlsxCfVoData cfvo2(type2, val2);
    cfRule->attrs[XlsxCfRuleData::A_cfvo1] = QVariant::fromValue(cfvo1);
    cfRule->attrs[XlsxCfRuleData::A_cfvo2] = QVariant::fromValue(cfvo2);

    d->appendRule(cfRule);
    return true;
}

bool ConditionalFormatting::addDataBarRule(const QColor &color, bool showData, bool stopIfTrue)
{
    return addDataBarRule(color, ConditionalFormatting::ValueObjectType::Min, QStringLiteral("0"),
                          ConditionalFormatting::ValueObjectType::Max, QStringLiteral("0"),
                          showData, stopIfTrue);
}

bool ConditionalFormatting::add2ColorScaleRule(const QColor &minColor, const QColor &maxColor, bool stopIfTrue)
{
    ValueObjectType type1 = ValueObjectType::Min;
    ValueObjectType type2 = ValueObjectType::Max;
    QString val1 = QStringLiteral("0");
    QString val2 = QStringLiteral("0");

    auto cfRule = std::make_shared<XlsxCfRuleData>();

    cfRule->attrs[XlsxCfRuleData::A_type] = QStringLiteral("colorScale");
    cfRule->attrs[XlsxCfRuleData::A_color1] = Color(minColor);
    cfRule->attrs[XlsxCfRuleData::A_color2] = Color(maxColor);
    if (stopIfTrue)
        cfRule->attrs[XlsxCfRuleData::A_stopIfTrue] = true;

    XlsxCfVoData cfvo1(type1, val1);
    XlsxCfVoData cfvo2(type2, val2);
    cfRule->attrs[XlsxCfRuleData::A_cfvo1] = QVariant::fromValue(cfvo1);
    cfRule->attrs[XlsxCfRuleData::A_cfvo2] = QVariant::fromValue(cfvo2);

    d->appendRule(cfRule);
    return true;
}

bool ConditionalFormatting::add3ColorScaleRule(const QColor &minColor, const QColor &midColor, const QColor &maxColor, bool stopIfTrue)
{
    ValueObjectType type1 = ValueObjectType::Min;
    ValueObjectType type2 = ValueObjectType::Percent;
    ValueObjectType type3 = ValueObjectType::Max;
    QString val1 = QStringLiteral("0");
    QString val2 = QStringLiteral("50");
    QString val3 = QStringLiteral("0");

    auto cfRule = std::make_shared<XlsxCfRuleData>();

    cfRule->attrs[XlsxCfRuleData::A_type] = QStringLiteral("colorScale");
    cfRule->attrs[XlsxCfRuleData::A_color1] = Color(minColor);
    cfRule->attrs[XlsxCfRuleData::A_color2] = Color(midColor);
    cfRule->attrs[XlsxCfRuleData::A_color3] = Color(maxColor);

    if (stopIfTrue)
        cfRule->attrs[XlsxCfRuleData::A_stopIfTrue] = true;

    XlsxCfVoData cfvo1(type1, val1);
    XlsxCfVoData cfvo2(type2, val2);
    XlsxCfVoData cfvo3(type3, val3);
    cfRule->attrs[XlsxCfRuleData::A_cfvo1] = QVariant::fromValue(cfvo1);
    cfRule->attrs[XlsxCfRuleData::A_cfvo2] = QVariant::fromValue(cfvo2);
    cfRule->attrs[XlsxCfRuleData::A_cfvo3] = QVariant::fromValue(cfvo3);

    d->appendRule(cfRule);
    return true;
}

int ConditionalFormatting::rulesCount() const
{
    return d->cfRules.size();
}

void ConditionalFormatting::clearRules()
{
    d->cfRules.clear();
}

bool ConditionalFormatting::removeRule(int ruleIndex)
{
    if (ruleIndex < 0 || ruleIndex >= d->cfRules.size()) return false;
    d->cfRules.removeAt(ruleIndex);
    return true;
}

bool ConditionalFormatting::setRulesPriority(int priority)
{
    if (d->cfRules.isEmpty()) return false;
    if (priority < 1) return false;
    for (auto &rule: d->cfRules) rule->priority = priority;
    return true;
}

bool ConditionalFormatting::setRulePriority(int ruleIndex, int priority)
{
    if (ruleIndex < 0 || ruleIndex >= d->cfRules.size()) return false;
    if (priority < 1) return false;
    d->cfRules[ruleIndex]->priority = priority;
    return true;
}

void ConditionalFormatting::setAutoDecrementPriority(bool autodecrement)
{
    d->autodecrement = autodecrement;
}

void ConditionalFormatting::updateRulesPriorities(bool firstRuleHasMaximumPriority)
{
    for (int index = 0; index < d->cfRules.size(); ++index) {
        d->cfRules[index]->priority = firstRuleHasMaximumPriority ? index+1 : d->cfRules.size() - index;
    }
}

std::optional<int> ConditionalFormatting::rulePriority(int ruleIndex) const
{
    if (ruleIndex < 0 || ruleIndex >= d->cfRules.size()) return std::nullopt;
    return d->cfRules.at(ruleIndex)->priority;
}

/*!
    Returns the ranges on which the validation will be applied.
 */
QList<CellRange> ConditionalFormatting::ranges() const
{
    return d->ranges;
}

/*!
    Add the \a cell on which the conditional formatting will apply to.
 */
void ConditionalFormatting::addCell(const CellReference &cell)
{
    d->ranges.append(CellRange(cell, cell));
}

/*!
    \overload
    Add the cell(\a row, \a col) on which the conditional formatting will apply to.
 */
void ConditionalFormatting::addCell(int row, int col)
{
    d->ranges.append(CellRange(row, col, row, col));
}

void ConditionalFormatting::addRange(int firstRow, int firstCol, int lastRow, int lastCol)
{
    d->ranges.append(CellRange(firstRow, firstCol, lastRow, lastCol));
}

void ConditionalFormatting::addRange(const CellRange &range)
{
    d->ranges.append(range);
}

void ConditionalFormatting::setRange(const CellRange &range)
{
    d->ranges.clear();
    d->ranges.append(range);
}

void ConditionalFormatting::clearRanges()
{
    d->ranges.clear();
}

bool ConditionalFormattingPrivate::readCfRule(QXmlStreamReader &reader, XlsxCfRuleData *rule, Styles *styles)
{
    Q_ASSERT(reader.name() == QLatin1String("cfRule"));
    QXmlStreamAttributes attrs = reader.attributes();
    if (attrs.hasAttribute(QLatin1String("type")))
        rule->attrs[XlsxCfRuleData::A_type] = attrs.value(QLatin1String("type")).toString();
    if (attrs.hasAttribute(QLatin1String("dxfId"))) {
        int id = attrs.value(QLatin1String("dxfId")).toInt();
        if (styles)
            rule->dxfFormat = styles->dxfFormat(id);
        else
            rule->dxfFormat.setDxfIndex(id);
    }
    rule->priority = attrs.value(QLatin1String("priority")).toInt();
    if (attrs.value(QLatin1String("stopIfTrue")) == QLatin1String("1")) {
        //default is false
        rule->attrs[XlsxCfRuleData::A_stopIfTrue] = QLatin1String("1");
    }
    if (attrs.value(QLatin1String("aboveAverage")) == QLatin1String("0")) {
        //default is true
        rule->attrs[XlsxCfRuleData::A_aboveAverage] = QLatin1String("0");
    }
    if (attrs.value(QLatin1String("percent")) == QLatin1String("1")) {
        //default is false
        rule->attrs[XlsxCfRuleData::A_percent] = QLatin1String("1");
    }
    if (attrs.value(QLatin1String("bottom")) == QLatin1String("1")) {
        //default is false
        rule->attrs[XlsxCfRuleData::A_bottom] = QLatin1String("1");
    }
    if (attrs.hasAttribute(QLatin1String("operator")))
        rule->attrs[XlsxCfRuleData::A_operator] = attrs.value(QLatin1String("operator")).toString();

    if (attrs.hasAttribute(QLatin1String("text")))
        rule->attrs[XlsxCfRuleData::A_text] = attrs.value(QLatin1String("text")).toString();

    if (attrs.hasAttribute(QLatin1String("timePeriod")))
        rule->attrs[XlsxCfRuleData::A_timePeriod] = attrs.value(QLatin1String("timePeriod")).toString();

    if (attrs.hasAttribute(QLatin1String("rank")))
        rule->attrs[XlsxCfRuleData::A_rank] = attrs.value(QLatin1String("rank")).toString();

    if (attrs.hasAttribute(QLatin1String("stdDev")))
        rule->attrs[XlsxCfRuleData::A_stdDev] = attrs.value(QLatin1String("stdDev")).toString();

    if (attrs.value(QLatin1String("equalAverage")) == QLatin1String("1")) {
        //default is false
        rule->attrs[XlsxCfRuleData::A_equalAverage] = QLatin1String("1");
    }

    while (!reader.atEnd()) {
        reader.readNextStartElement();
        if (reader.tokenType() == QXmlStreamReader::StartElement) {
            if (reader.name() == QLatin1String("formula")) {
                const QString f = reader.readElementText();
                if (!rule->attrs.contains(XlsxCfRuleData::A_formula1))
                    rule->attrs[XlsxCfRuleData::A_formula1] = f;
                else if (!rule->attrs.contains(XlsxCfRuleData::A_formula2))
                    rule->attrs[XlsxCfRuleData::A_formula2] = f;
                else if (!rule->attrs.contains(XlsxCfRuleData::A_formula3))
                    rule->attrs[XlsxCfRuleData::A_formula3] = f;
            } else if (reader.name() == QLatin1String("dataBar")) {
                readCfDataBar(reader, rule);
            } else if (reader.name() == QLatin1String("colorScale")) {
                readCfColorScale(reader, rule);
            }
        }
        if (reader.tokenType() == QXmlStreamReader::EndElement
                && reader.name() == QStringLiteral("conditionalFormatting")) {
            break;
        }
    }
    return true;
}

bool ConditionalFormattingPrivate::readCfDataBar(QXmlStreamReader &reader, XlsxCfRuleData *rule)
{
    Q_ASSERT(reader.name() == QLatin1String("dataBar"));
    QXmlStreamAttributes attrs = reader.attributes();
    if (attrs.value(QLatin1String("showValue")) == QLatin1String("0"))
        rule->attrs[XlsxCfRuleData::A_hideData] = QStringLiteral("1");

    while (!reader.atEnd()) {
        reader.readNextStartElement();
        if (reader.tokenType() == QXmlStreamReader::StartElement) {
            if (reader.name() == QLatin1String("cfvo")) {
                XlsxCfVoData data;
                readCfVo(reader, data);
                if (!rule->attrs.contains(XlsxCfRuleData::A_cfvo1))
                    rule->attrs[XlsxCfRuleData::A_cfvo1] = QVariant::fromValue(data);
                else
                    rule->attrs[XlsxCfRuleData::A_cfvo2] = QVariant::fromValue(data);
            } else if (reader.name() == QLatin1String("color")) {
                Color color;
                color.read(reader);
                rule->attrs[XlsxCfRuleData::A_color1] = color;
            }
        }
        if (reader.tokenType() == QXmlStreamReader::EndElement
                && reader.name() == QStringLiteral("dataBar")) {
            break;
        }
    }

    return true;
}

bool ConditionalFormattingPrivate::readCfColorScale(QXmlStreamReader &reader, XlsxCfRuleData *rule)
{
    Q_ASSERT(reader.name() == QLatin1String("colorScale"));

    while (!reader.atEnd()) {
        reader.readNextStartElement();
        if (reader.tokenType() == QXmlStreamReader::StartElement) {
            if (reader.name() == QLatin1String("cfvo")) {
                XlsxCfVoData data;
                readCfVo(reader, data);
                if (!rule->attrs.contains(XlsxCfRuleData::A_cfvo1))
                    rule->attrs[XlsxCfRuleData::A_cfvo1] = QVariant::fromValue(data);
                else if (!rule->attrs.contains(XlsxCfRuleData::A_cfvo2))
                    rule->attrs[XlsxCfRuleData::A_cfvo2] = QVariant::fromValue(data);
                else
                    rule->attrs[XlsxCfRuleData::A_cfvo3] = QVariant::fromValue(data);
            } else if (reader.name() == QLatin1String("color")) {
                Color color;
                color.read(reader);
                if (!rule->attrs.contains(XlsxCfRuleData::A_color1))
                    rule->attrs[XlsxCfRuleData::A_color1] = color;
                else if (!rule->attrs.contains(XlsxCfRuleData::A_color2))
                    rule->attrs[XlsxCfRuleData::A_color2] = color;
                else
                    rule->attrs[XlsxCfRuleData::A_color3] = color;
            }
        }
        if (reader.tokenType() == QXmlStreamReader::EndElement
                && reader.name() == QStringLiteral("colorScale")) {
            break;
        }
    }

    return true;
}

void ConditionalFormattingPrivate::appendRule(std::shared_ptr<XlsxCfRuleData> &rule)
{
    if (autodecrement) {
        auto maxEl = std::max_element(cfRules.cbegin(), cfRules.cend(),
                                      [](auto left, auto right){return left->priority < right->priority;});
        if (maxEl != cfRules.cend())
            rule->priority = maxEl->get()->priority+1;
        //else cfRule->priority = 1;
    }
    cfRules.append(rule);
}

bool ConditionalFormattingPrivate::readCfVo(QXmlStreamReader &reader, XlsxCfVoData &cfvo)
{
    Q_ASSERT(reader.name() == QStringLiteral("cfvo"));

    QXmlStreamAttributes attrs = reader.attributes();

    QString type = attrs.value(QLatin1String("type")).toString();
    ConditionalFormatting::ValueObjectType t;
    if (type == QLatin1String("formula"))
        t = ConditionalFormatting::ValueObjectType::Formula;
    else if (type == QLatin1String("max"))
        t = ConditionalFormatting::ValueObjectType::Max;
    else if (type == QLatin1String("min"))
        t = ConditionalFormatting::ValueObjectType::Min;
    else if (type == QLatin1String("num"))
        t = ConditionalFormatting::ValueObjectType::Num;
    else if (type == QLatin1String("percent"))
        t = ConditionalFormatting::ValueObjectType::Percent;
    else //if (type == QLatin1String("percentile"))
        t = ConditionalFormatting::ValueObjectType::Percentile;

    cfvo.type = t;
    cfvo.value = attrs.value(QLatin1String("val")).toString();
    if (attrs.value(QLatin1String("gte")) == QLatin1String("0")) {
        //default is true
        cfvo.gte = false;
    }
    return true;
}

bool ConditionalFormatting::loadFromXml(QXmlStreamReader &reader, Styles *styles)
{
    Q_ASSERT(reader.name() == QStringLiteral("conditionalFormatting"));

    d->ranges.clear();
    d->cfRules.clear();
    QXmlStreamAttributes attrs = reader.attributes();
    const QString sqref = attrs.value(QLatin1String("sqref")).toString();
    const auto sqrefParts = sqref.split(QLatin1Char(' '));
    for (const QString &range : sqrefParts) {
        this->addRange(range);
    }

    while (!reader.atEnd()) {
        reader.readNextStartElement();
        if (reader.tokenType() == QXmlStreamReader::StartElement) {
            if (reader.name() == QLatin1String("cfRule")) {
                auto cfRule = std::make_shared<XlsxCfRuleData>();
                d->readCfRule(reader, cfRule.get(), styles);
                d->cfRules.append(cfRule);
            }
        }
        if (reader.tokenType() == QXmlStreamReader::EndElement
                && reader.name() == QStringLiteral("conditionalFormatting")) {
            break;
        }
    }


    return true;
}

Format ConditionalFormatting::predefinedFormat(PredefinedFormat format)
{
    Format f;
    switch (format) {
    case PredefinedFormat::Format1:
        f.setFontColor(QColor(0xFF9C0006));
        f.setPatternForegroundColor(QColor(0xFFFFC7CE));
        break;
    case PredefinedFormat::Format2:
        f.setFontColor(QColor(0xFF9C6500));
        f.setPatternForegroundColor(QColor(0xFFFFEB9C));
        break;
    case PredefinedFormat::Format3:
        f.setFontColor(QColor(0xFF006100));
        f.setPatternForegroundColor(QColor(0xFFC6EFCE));
        break;
    case PredefinedFormat::Format4:
        f.setPatternForegroundColor(QColor(0xFFFFC7CE));
        break;
    case PredefinedFormat::Format5:
        f.setFontColor(QColor(0xFF9C0006));
        break;
    case PredefinedFormat::Format6:
        f.setBorderColor(QColor(0xFF9C0006));
        break;
    }
    return f;
}

bool ConditionalFormatting::saveToXml(QXmlStreamWriter &writer) const
{
    writer.writeStartElement(QStringLiteral("conditionalFormatting"));
    QStringList sqref;
    const auto rangeList = ranges();
    for (const CellRange &range : rangeList) {
        sqref.append(range.toString());
    }
    writer.writeAttribute(QStringLiteral("sqref"), sqref.join(QLatin1String(" ")));

    for (int i=0; i<d->cfRules.size(); ++i) {
        const std::shared_ptr<XlsxCfRuleData> &rule = d->cfRules[i];
        writer.writeStartElement(QStringLiteral("cfRule"));
        writer.writeAttribute(QStringLiteral("type"), rule->attrs[XlsxCfRuleData::A_type].toString());
        if (rule->dxfFormat.dxfIndexValid())
            writer.writeAttribute(QStringLiteral("dxfId"), QString::number(rule->dxfFormat.dxfIndex()));
        writer.writeAttribute(QStringLiteral("priority"), QString::number(rule->priority));

        auto it = rule->attrs.constFind(XlsxCfRuleData::A_stopIfTrue);
        if (it != rule->attrs.constEnd())
            writer.writeAttribute(QStringLiteral("stopIfTrue"), it.value().toString());

        it = rule->attrs.constFind(XlsxCfRuleData::A_aboveAverage);
        if (it != rule->attrs.constEnd())
            writer.writeAttribute(QStringLiteral("aboveAverage"), it.value().toString());

        it = rule->attrs.constFind(XlsxCfRuleData::A_percent);
        if (it != rule->attrs.constEnd())
            writer.writeAttribute(QStringLiteral("percent"), it.value().toString());

        it = rule->attrs.constFind(XlsxCfRuleData::A_bottom);
        if (it != rule->attrs.constEnd())
            writer.writeAttribute(QStringLiteral("bottom"), it.value().toString());

        it = rule->attrs.constFind(XlsxCfRuleData::A_operator);
        if (it != rule->attrs.constEnd())
            writer.writeAttribute(QStringLiteral("operator"), it.value().toString());

        it = rule->attrs.constFind(XlsxCfRuleData::A_text);
        if (it != rule->attrs.constEnd())
            writer.writeAttribute(QStringLiteral("text"), it.value().toString());

        it = rule->attrs.constFind(XlsxCfRuleData::A_timePeriod);
        if (it != rule->attrs.constEnd())
            writer.writeAttribute(QStringLiteral("timePeriod"), it.value().toString());

        it = rule->attrs.constFind(XlsxCfRuleData::A_rank);
        if (it != rule->attrs.constEnd())
            writer.writeAttribute(QStringLiteral("rank"), it.value().toString());

        it = rule->attrs.constFind(XlsxCfRuleData::A_stdDev);
        if (it != rule->attrs.constEnd())
            writer.writeAttribute(QStringLiteral("stdDev"), it.value().toString());

        it = rule->attrs.constFind(XlsxCfRuleData::A_equalAverage);
        if (it != rule->attrs.constEnd())
            writer.writeAttribute(QStringLiteral("equalAverage"), it.value().toString());

        if (rule->attrs[XlsxCfRuleData::A_type] == QLatin1String("dataBar")) {
            writer.writeStartElement(QStringLiteral("dataBar"));
            if (rule->attrs.contains(XlsxCfRuleData::A_hideData))
                writer.writeAttribute(QStringLiteral("showValue"), QStringLiteral("0"));
            d->writeCfVo(writer, rule->attrs[XlsxCfRuleData::A_cfvo1].value<XlsxCfVoData>());
            d->writeCfVo(writer, rule->attrs[XlsxCfRuleData::A_cfvo2].value<XlsxCfVoData>());
            rule->attrs[XlsxCfRuleData::A_color1].value<Color>().write(writer, QLatin1String("color"));
            writer.writeEndElement();//dataBar
        } else if (rule->attrs[XlsxCfRuleData::A_type] == QLatin1String("colorScale")) {
            writer.writeStartElement(QStringLiteral("colorScale"));
            d->writeCfVo(writer, rule->attrs[XlsxCfRuleData::A_cfvo1].value<XlsxCfVoData>());
            d->writeCfVo(writer, rule->attrs[XlsxCfRuleData::A_cfvo2].value<XlsxCfVoData>());

            it = rule->attrs.constFind(XlsxCfRuleData::A_cfvo3);
            if (it != rule->attrs.constEnd())
                d->writeCfVo(writer, it.value().value<XlsxCfVoData>());

            rule->attrs[XlsxCfRuleData::A_color1].value<Color>().write(writer, QLatin1String("color"));
            rule->attrs[XlsxCfRuleData::A_color2].value<Color>().write(writer, QLatin1String("color"));

            it = rule->attrs.constFind(XlsxCfRuleData::A_color3);
            if (it != rule->attrs.constEnd())
                it.value().value<Color>().write(writer, QLatin1String("color"));

            writer.writeEndElement();//colorScale
        }


        it = rule->attrs.constFind(XlsxCfRuleData::A_formula1_temp); // <formula>NOT(ISERROR(SEARCH("a",B6)))</formula>
        if (it != rule->attrs.constEnd() && !ranges().isEmpty()) {
            auto startCell = ranges().first().topLeft().toString();
            writer.writeTextElement(QStringLiteral("formula"), it.value().toString().arg(startCell));
        } else if ((it = rule->attrs.constFind(XlsxCfRuleData::A_formula1)) != rule->attrs.constEnd()) {
            writer.writeTextElement(QStringLiteral("formula"), it.value().toString());
        }

        it = rule->attrs.constFind(XlsxCfRuleData::A_formula2);
        if (it != rule->attrs.constEnd())
            writer.writeTextElement(QStringLiteral("formula"), it.value().toString());

        it = rule->attrs.constFind(XlsxCfRuleData::A_formula3);
        if (it != rule->attrs.constEnd())
            writer.writeTextElement(QStringLiteral("formula"), it.value().toString());

        writer.writeEndElement(); //cfRule
    }

    writer.writeEndElement(); //conditionalFormatting
    return true;
}

}
