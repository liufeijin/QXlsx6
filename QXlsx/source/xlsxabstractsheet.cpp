// xlsxabstractsheet.cpp

#include <QtGlobal>

#include "xlsxabstractsheet.h"
#include "xlsxabstractsheet_p.h"
#include "xlsxworkbook.h"
#include "xlsxutility_p.h"

namespace QXlsx {

AbstractSheetPrivate::AbstractSheetPrivate(AbstractSheet *p, AbstractSheet::CreateFlag flag)
    : AbstractOOXmlFilePrivate(p, flag)
{
    type = AbstractSheet::Type::Worksheet;
    sheetState = AbstractSheet::Visibility::Visible;
}

AbstractSheetPrivate::~AbstractSheetPrivate()
{
}



/*!
 * \internal
 */
AbstractSheet::AbstractSheet(const QString &name, int id, Workbook *workbook, AbstractSheetPrivate *d) :
    AbstractOOXmlFile(d)
{
    d_func()->name = name;
    d_func()->id = id;
    d_func()->workbook = workbook;
}


QString AbstractSheet::name() const
{
    Q_D(const AbstractSheet);
    return d->name;
}

/*!
 * \internal
 */
void AbstractSheet::setName(const QString &sheetName)
{
    Q_D(AbstractSheet);
    d->name = sheetName;
}

/*!
 * Returns the type of the sheet.
 */
AbstractSheet::Type AbstractSheet::type() const
{
    Q_D(const AbstractSheet);
    return d->type;
}

/*!
 * \internal
 */
void AbstractSheet::setType(Type type)
{
    Q_D(AbstractSheet);
    d->type = type;
}

AbstractSheet::Visibility AbstractSheet::visibility() const
{
    Q_D(const AbstractSheet);
    return d->sheetState;
}

void AbstractSheet::setVisibility(Visibility visibility)
{
    Q_D(AbstractSheet);
    d->sheetState = visibility;
}

bool AbstractSheet::isHidden() const
{
    Q_D(const AbstractSheet);
    return d->sheetState != Visibility::Visible;
}

bool AbstractSheet::isVisible() const
{
    return !isHidden();
}

void AbstractSheet::setHidden(bool hidden)
{
    Q_D(AbstractSheet);
    if (!hidden) d->sheetState = Visibility::Visible;
    else if (d->sheetState == Visibility::Visible) d->sheetState = Visibility::Hidden;
}

void AbstractSheet::setVisible(bool visible)
{
    Q_D(AbstractSheet);
    if (visible) d->sheetState = Visibility::Visible;
    else if (d->sheetState == Visibility::Visible) d->sheetState = Visibility::Hidden;
}

HeaderFooter AbstractSheet::headerFooter() const
{
    Q_D(const AbstractSheet);
    return d->headerFooter;
}

HeaderFooter &AbstractSheet::headerFooter()
{
    Q_D(AbstractSheet);
    return d->headerFooter;
}

QMap<PageMargins::Position, double> AbstractSheet::pageMarginsInches() const
{
    Q_D(const AbstractSheet);
    return d->pageMargins.marginsInches();
}

QMap<PageMargins::Position, double> AbstractSheet::pageMarginsMm() const
{
    Q_D(const AbstractSheet);
    return d->pageMargins.marginsMm();
}

void AbstractSheet::setPageMarginsInches(double left, double top, double right, double bottom, double header, double footer)
{
    Q_D(AbstractSheet);
    d->pageMargins.setMarginsInches(left, top, right, bottom, header, footer);
}

void AbstractSheet::setPageMarginsMm(double left, double top, double right, double bottom, double header, double footer)
{
    Q_D(AbstractSheet);
    d->pageMargins.setMarginsMm(left, top, right, bottom, header, footer);
}

PageMargins AbstractSheet::pageMargins() const
{
    Q_D(const AbstractSheet);
    return d->pageMargins;
}

PageMargins &AbstractSheet::pageMargins()
{
    Q_D(AbstractSheet);
    return d->pageMargins;
}

void AbstractSheet::setPageMargins(const PageMargins &margins)
{
    Q_D(AbstractSheet);
    d->pageMargins = margins;
}

/*!
 * \internal
 */
int AbstractSheet::id() const
{
    Q_D(const AbstractSheet);
    return d->id;
}

/*!
 * \internal
 */
Drawing *AbstractSheet::drawing() const
{
    Q_D(const AbstractSheet);
    return d->drawing.get();
}

/*!
 * Return the workbook
 */
Workbook *AbstractSheet::workbook() const
{
    Q_D(const AbstractSheet);
    return d->workbook;
}

bool HeaderFooter::isValid() const
{
    if (differentOddEven.has_value()) return true;
    if (differentFirst.has_value()) return true;
    if (scaleWithDoc.has_value()) return true;
    if (alignWithMargins.has_value()) return true;

    if (!oddHeader.isEmpty()) return true;
    if (!oddFooter.isEmpty()) return true;
    if (!evenHeader.isEmpty() && differentOddEven.value_or(false)) return true;
    if (!evenFooter.isEmpty() && differentOddEven.value_or(false)) return true;
    if (!firstHeader.isEmpty() && differentFirst.value_or(false)) return true;
    if (!firstFooter.isEmpty() && differentFirst.value_or(false)) return true;
    return false;
}

void HeaderFooter::write(QXmlStreamWriter &writer) const
{
    if (isValid()) {
        writer.writeStartElement(QStringLiteral("headerFooter"));
        writeAttribute(writer, QLatin1String("differentOddEven"), differentOddEven);
        writeAttribute(writer, QLatin1String("differentFirst"), differentFirst);
        writeAttribute(writer, QLatin1String("alignWithMargins"), alignWithMargins);
        writeAttribute(writer, QLatin1String("scaleWithDoc"), scaleWithDoc);
        if (!oddHeader.isEmpty())
            writer.writeTextElement(QStringLiteral("oddHeader"), oddHeader);
        if (!oddFooter.isEmpty())
            writer.writeTextElement(QStringLiteral("oddFooter"), oddFooter);
        if (!evenHeader.isEmpty() && differentOddEven.value_or(false))
            writer.writeTextElement(QStringLiteral("evenHeader"), evenHeader);
        if (!evenFooter.isEmpty() && differentOddEven.value_or(false))
            writer.writeTextElement(QStringLiteral("evenFooter"), evenFooter);
        if (!firstHeader.isEmpty() && differentFirst.value_or(false))
            writer.writeTextElement(QStringLiteral("firstHeader"), firstHeader);
        if (!firstFooter.isEmpty() && differentFirst.value_or(false))
            writer.writeTextElement(QStringLiteral("firstFooter"), firstFooter);
        writer.writeEndElement();
    }
}

void HeaderFooter::read(QXmlStreamReader &reader)
{
    const auto &name = reader.name();
    const auto &a = reader.attributes();
    parseAttributeBool(a, QLatin1String("differentOddEven"), differentOddEven);
    parseAttributeBool(a, QLatin1String("differentFirst"), differentFirst);
    parseAttributeBool(a, QLatin1String("scaleWithDoc"), scaleWithDoc);
    parseAttributeBool(a, QLatin1String("differentOddEven"), alignWithMargins);
    while (!reader.atEnd()) {
        auto token = reader.readNext();
        if (token == QXmlStreamReader::StartElement) {
            if (reader.name() == QLatin1String("oddHeader")) oddHeader = reader.readElementText();
            else if (reader.name() == QLatin1String("oddFooter")) oddFooter = reader.readElementText();
            else if (reader.name() == QLatin1String("evenHeader")) evenHeader = reader.readElementText();
            else if (reader.name() == QLatin1String("evenFooter")) evenFooter = reader.readElementText();
            else if (reader.name() == QLatin1String("firstHeader")) firstHeader = reader.readElementText();
            else if (reader.name() == QLatin1String("firstFooter")) firstFooter = reader.readElementText();
            else reader.skipCurrentElement();
        }
        else if (token == QXmlStreamReader::EndElement && reader.name() == name)
            break;
    }
}

PageMargins::PageMargins(double left, double top, double right, double bottom,
            double header, double footer)
{
    mMargins.insert(Position::Left, left);
    mMargins.insert(Position::Top, top);
    mMargins.insert(Position::Right, right);
    mMargins.insert(Position::Bottom, bottom);
    mMargins.insert(Position::Header, header);
    mMargins.insert(Position::Footer, footer);
}

void PageMargins::setMarginsInches(double left, double top, double right, double bottom, double header, double footer)
{
    mMargins.insert(Position::Left, left);
    mMargins.insert(Position::Top, top);
    mMargins.insert(Position::Right, right);
    mMargins.insert(Position::Bottom, bottom);
    mMargins.insert(Position::Header, header);
    mMargins.insert(Position::Footer, footer);
}

void PageMargins::setMarginInches(Position pos, double value)
{
    mMargins.insert(pos, value);
}
void PageMargins::setLeftMarginInches(double value)
{
    mMargins.insert(Position::Left, value);
}
void PageMargins::setRightMarginInches(double value)
{
    mMargins.insert(Position::Right, value);
}
void PageMargins::setTopMarginInches(double value)
{
    mMargins.insert(Position::Top, value);
}
void PageMargins::setBottomMarginInches(double value)
{
    mMargins.insert(Position::Bottom, value);
}
void PageMargins::setHeaderMarginInches(double value)
{
    mMargins.insert(Position::Header, value);
}
void PageMargins::setFooterMarginInches(double value)
{
    mMargins.insert(Position::Footer, value);
}

void PageMargins::setMarginsMm(double left, double top, double right, double bottom, double header, double footer)
{
    mMargins.insert(Position::Left, left / 25.4);
    mMargins.insert(Position::Top, top / 25.4);
    mMargins.insert(Position::Right, right / 25.4);
    mMargins.insert(Position::Bottom, bottom / 25.4);
    mMargins.insert(Position::Header, header / 25.4);
    mMargins.insert(Position::Footer, footer / 25.4);
}
void PageMargins::setMarginMm(Position pos, double value)
{
    mMargins.insert(pos, value / 25.4);
}
void PageMargins::setLeftMarginMm(double value)
{
    mMargins.insert(Position::Left, value / 25.4);
}
void PageMargins::setRightMarginMm(double value)
{
    mMargins.insert(Position::Right, value / 25.4);
}
void PageMargins::setTopMarginMm(double value)
{
    mMargins.insert(Position::Top, value / 25.4);
}
void PageMargins::setBottomMarginMm(double value)
{
    mMargins.insert(Position::Bottom, value / 25.4);
}
void PageMargins::setHeaderMarginMm(double value)
{
    mMargins.insert(Position::Header, value / 25.4);
}
void PageMargins::setFooterMarginMm(double value)
{
    mMargins.insert(Position::Footer, value / 25.4);
}
QMap<PageMargins::Position, double> PageMargins::marginsInches() const
{
    return mMargins;
}
QMap<PageMargins::Position, double> PageMargins::marginsMm() const
{
    auto mm = mMargins;
    for (auto &m: mm) m*=25.4;
    return mm;
}
double PageMargins::marginInches(Position pos) const
{
    return mMargins.value(pos, 0.0);
}
double PageMargins::leftMarginInches() const
{
    return mMargins.value(Position::Left, 0.0);
}
double PageMargins::rightMarginInches() const
{
    return mMargins.value(Position::Right, 0.0);
}
double PageMargins::topMarginInches() const
{
    return mMargins.value(Position::Top, 0.0);
}
double PageMargins::bottomMarginInches() const
{
    return mMargins.value(Position::Bottom, 0.0);
}
double PageMargins::headerMarginInches() const
{
    return mMargins.value(Position::Header, 0.0);
}
double PageMargins::footerMarginInches() const
{
    return mMargins.value(Position::Footer, 0.0);
}
double PageMargins::marginMm(Position pos) const
{
    return mMargins.value(pos, 0.0) * 25.4;
}
double PageMargins::leftMarginMm() const
{
    return mMargins.value(Position::Left, 0.0) * 25.4;
}
double PageMargins::rightMarginMm() const
{
    return mMargins.value(Position::Right, 0.0) * 25.4;
}
double PageMargins::topMarginMm() const
{
    return mMargins.value(Position::Top, 0.0) * 25.4;
}
double PageMargins::bottomMarginMm() const
{
    return mMargins.value(Position::Bottom, 0.0) * 25.4;
}
double PageMargins::headerMarginMm() const
{
    return mMargins.value(Position::Header, 0.0) * 25.4;
}
double PageMargins::footerMarginMm() const
{
    return mMargins.value(Position::Footer, 0.0) * 25.4;
}

bool PageMargins::hasMargin(Position pos) const
{
    return mMargins.contains(pos);
}
bool PageMargins::isValid() const
{
    return !mMargins.isEmpty();
}
void PageMargins::write(QXmlStreamWriter &writer) const
{
    if (isValid()) {
        writer.writeEmptyElement(QLatin1String("pageMargins"));
        writer.writeAttribute(QLatin1String("left"), QString::number(mMargins.value(Position::Left, 0.0), 'f'));
        writer.writeAttribute(QLatin1String("right"), QString::number(mMargins.value(Position::Right, 0.0), 'f'));
        writer.writeAttribute(QLatin1String("top"), QString::number(mMargins.value(Position::Top, 0.0), 'f'));
        writer.writeAttribute(QLatin1String("bottom"), QString::number(mMargins.value(Position::Bottom, 0.0), 'f'));
        writer.writeAttribute(QLatin1String("header"), QString::number(mMargins.value(Position::Header, 0.0), 'f'));
        writer.writeAttribute(QLatin1String("footer"), QString::number(mMargins.value(Position::Footer, 0.0), 'f'));
    }
}

void PageMargins::read(QXmlStreamReader &reader)
{
    const auto &a = reader.attributes();
    if (a.hasAttribute(QLatin1String("left"))) mMargins.insert(Position::Left, a.value(QLatin1String("left")).toDouble());
    if (a.hasAttribute(QLatin1String("right"))) mMargins.insert(Position::Right, a.value(QLatin1String("right")).toDouble());
    if (a.hasAttribute(QLatin1String("top"))) mMargins.insert(Position::Top, a.value(QLatin1String("top")).toDouble());
    if (a.hasAttribute(QLatin1String("bottom"))) mMargins.insert(Position::Bottom, a.value(QLatin1String("bottom")).toDouble());
    if (a.hasAttribute(QLatin1String("header"))) mMargins.insert(Position::Header, a.value(QLatin1String("header")).toDouble());
    if (a.hasAttribute(QLatin1String("footer"))) mMargins.insert(Position::Footer, a.value(QLatin1String("footer")).toDouble());
}

}
