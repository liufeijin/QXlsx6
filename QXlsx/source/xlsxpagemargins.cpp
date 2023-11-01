#include "xlsxpagemargins.h"

namespace QXlsx {
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

PageMargins PageMargins::defaultPageMargins()
{
    return {};
}

PageMargins PageMargins::pageMarginsMm(double left, double top, double right, double bottom, double header, double footer)
{
    PageMargins p;
    p.setMarginsMm(left, top, right, bottom, header, footer);
    return p;
}

PageMargins PageMargins::pageMarginsInches(double left, double top, double right, double bottom, double header, double footer)
{
    PageMargins p;
    p.setMarginsInches(left, top, right, bottom, header, footer);
    return p;
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
