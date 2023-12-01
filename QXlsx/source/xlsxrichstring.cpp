// xlsxrichstring.cpp

#include <QtGlobal>
#include <QDebug>
#include <QTextDocument>
#include <QTextFragment>

#include "xlsxrichstring.h"
#include "xlsxrichstring_p.h"
#include "xlsxformat_p.h"
#include "xlsxutility_p.h"
#include "xlsxcolor.h"

namespace QXlsx {

RichStringPrivate::RichStringPrivate()
    :_dirty(true)
{

}

RichStringPrivate::RichStringPrivate(const RichStringPrivate &other)
    :QSharedData(other), fragmentTexts(other.fragmentTexts)
    ,fragmentFormats(other.fragmentFormats)
    , _idKey(other.idKey()), _dirty(other._dirty)
{

}

RichStringPrivate::~RichStringPrivate()
{

}

RichString::RichString()
    :d(new RichStringPrivate)
{
}

RichString::RichString(const QString& text)
    :d(new RichStringPrivate)
{
    addFragment(text, Format());
}

RichString::RichString(const RichString &other)
    :d(other.d)
{

}

RichString::~RichString()
{

}

RichString &RichString::operator =(const RichString &other)
{
    if (*this != other) this->d = other.d;
    return *this;
}

RichString::operator QVariant() const
{
    const auto& cref
#if QT_VERSION >= 0x060000 // Qt 6.0 or over
        = QMetaType::fromType<RichString>();
#else
        = qMetaTypeId<RichString>() ;
#endif
    return QVariant(cref, this);
}

void RichString::writeFormat(QXmlStreamWriter &writer, const Format &format)
{
    if (!format.hasFontData())
        return;

    if (!format.fontName().isEmpty()) {
        writer.writeEmptyElement(QStringLiteral("rFont"));
        writer.writeAttribute(QStringLiteral("val"), format.fontName());
    }
    if (format.hasProperty(FormatPrivate::P_Font_Charset))
        writeEmptyElement(writer, QLatin1String("charset"), format.property(FormatPrivate::P_Font_Charset).toInt());
    if (format.hasProperty(FormatPrivate::P_Font_Family)) {
        writer.writeEmptyElement(QStringLiteral("family"));
        writer.writeAttribute(QStringLiteral("val"), QString::number(format.intProperty(FormatPrivate::P_Font_Family)));
    }
    if (format.fontBold())
        writer.writeEmptyElement(QStringLiteral("b"));
    if (format.fontItalic())
        writer.writeEmptyElement(QStringLiteral("i"));
    if (format.fontStrikeOut())
        writer.writeEmptyElement(QStringLiteral("strike"));
    if (format.fontOutline())
        writer.writeEmptyElement(QStringLiteral("outline"));
    if (format.boolProperty(FormatPrivate::P_Font_Shadow))
        writer.writeEmptyElement(QStringLiteral("shadow"));
    if (format.hasProperty(FormatPrivate::P_Font_Condense)) {
        writer.writeEmptyElement(QStringLiteral("condense"));
        writeAttribute(writer, QLatin1String("val"), format.property(FormatPrivate::P_Font_Condense).toInt());
    }
    if (format.hasProperty(FormatPrivate::P_Font_Extend)) {
        writer.writeEmptyElement(QStringLiteral("extend"));
        writeAttribute(writer, QLatin1String("val"), format.property(FormatPrivate::P_Font_Extend).toInt());
    }
    if (format.hasProperty(FormatPrivate::P_Font_Color)) {
        Color color = format.property(FormatPrivate::P_Font_Color).value<Color>();
        color.write(writer, QStringLiteral("color"));
    }
    if (format.hasProperty(FormatPrivate::P_Font_Size)) {
        writer.writeEmptyElement(QStringLiteral("sz"));
        writer.writeAttribute(QStringLiteral("val"), QString::number(format.fontSize()));
    }
    if (format.hasProperty(FormatPrivate::P_Font_Underline)) {
        Format::FontUnderline u = format.fontUnderline();
        if (u != Format::FontUnderlineNone) {
            writer.writeEmptyElement(QStringLiteral("u"));
            if (u== Format::FontUnderlineDouble)
                writer.writeAttribute(QStringLiteral("val"), QStringLiteral("double"));
            else if (u == Format::FontUnderlineSingleAccounting)
                writer.writeAttribute(QStringLiteral("val"), QStringLiteral("singleAccounting"));
            else if (u == Format::FontUnderlineDoubleAccounting)
                writer.writeAttribute(QStringLiteral("val"), QStringLiteral("doubleAccounting"));
        }
    }
    if (format.hasProperty(FormatPrivate::P_Font_Script)) {
        Format::FontScript s = format.fontScript();
        if (s != Format::FontScriptNormal) {
            writer.writeEmptyElement(QStringLiteral("vertAlign"));
            if (s == Format::FontScriptSuper)
                writer.writeAttribute(QStringLiteral("val"), QStringLiteral("superscript"));
            else
                writer.writeAttribute(QStringLiteral("val"), QStringLiteral("subscript"));
        }
    }
    if (format.hasProperty(FormatPrivate::P_Font_Scheme)) {
        writer.writeEmptyElement(QStringLiteral("scheme"));
        writer.writeAttribute(QStringLiteral("val"), format.stringProperty(FormatPrivate::P_Font_Scheme));
    }
}

void RichString::readRichStringPart(QXmlStreamReader &reader)
{
    Q_ASSERT(reader.name() == QLatin1String("r"));

    QString text;
    Format format;
    while (!reader.atEnd() && !(reader.name() == QLatin1String("r") && reader.tokenType() == QXmlStreamReader::EndElement)) {
        reader.readNextStartElement();
        if (reader.tokenType() == QXmlStreamReader::StartElement) {
            if (reader.name() == QLatin1String("rPr")) {
                format = readFormat(reader);
            } else if (reader.name() == QLatin1String("t")) {
                text = reader.readElementText();
            }
        }
    }
    addFragment(text, format);
}

void RichString::readPlainStringPart(QXmlStreamReader &reader)
{
    Q_ASSERT(reader.name() == QLatin1String("t"));

    //QXmlStreamAttributes attributes = reader.attributes();

    // NOTICE: CHECK POINT
    QString text = reader.readElementText();
    addFragment(text, Format());
}

Format RichString::readFormat(QXmlStreamReader &reader) const
{
    Q_ASSERT(reader.name() == QLatin1String("rPr"));
    Format format;
    while (!reader.atEnd() && !(reader.name() == QLatin1String("rPr") && reader.tokenType() == QXmlStreamReader::EndElement)) {
        reader.readNextStartElement();
        if (reader.tokenType() == QXmlStreamReader::StartElement) {
            QXmlStreamAttributes attributes = reader.attributes();
            if (reader.name() == QLatin1String("rFont")) {
                format.setFontName(attributes.value(QLatin1String("val")).toString());
            } else if (reader.name() == QLatin1String("charset")) {
                format.setProperty(FormatPrivate::P_Font_Charset, attributes.value(QLatin1String("val")).toInt());
            } else if (reader.name() == QLatin1String("family")) {
                format.setProperty(FormatPrivate::P_Font_Family, attributes.value(QLatin1String("val")).toInt());
            } else if (reader.name() == QLatin1String("b")) {
                format.setFontBold(true);
            } else if (reader.name() == QLatin1String("i")) {
                format.setFontItalic(true);
            } else if (reader.name() == QLatin1String("strike")) {
                format.setFontStrikeOut(true);
            } else if (reader.name() == QLatin1String("outline")) {
                format.setFontOutline(true);
            } else if (reader.name() == QLatin1String("shadow")) {
                format.setProperty(FormatPrivate::P_Font_Shadow, true);
            } else if (reader.name() == QLatin1String("condense")) {
                format.setProperty(FormatPrivate::P_Font_Condense, attributes.value(QLatin1String("val")).toInt());
            } else if (reader.name() == QLatin1String("extend")) {
                format.setProperty(FormatPrivate::P_Font_Extend, attributes.value(QLatin1String("val")).toInt());
            } else if (reader.name() == QLatin1String("color")) {
                Color color;
                color.read(reader);
                format.setProperty(FormatPrivate::P_Font_Color, color);
            } else if (reader.name() == QLatin1String("sz")) {
                format.setFontSize(attributes.value(QLatin1String("val")).toDouble());
            } else if (reader.name() == QLatin1String("u")) {
                QString value = attributes.value(QLatin1String("val")).toString();
                if (value == QLatin1String("double"))
                    format.setFontUnderline(Format::FontUnderlineDouble);
                else if (value == QLatin1String("doubleAccounting"))
                    format.setFontUnderline(Format::FontUnderlineDoubleAccounting);
                else if (value == QLatin1String("singleAccounting"))
                    format.setFontUnderline(Format::FontUnderlineSingleAccounting);
                else
                    format.setFontUnderline(Format::FontUnderlineSingle);
            } else if (reader.name() == QLatin1String("vertAlign")) {
                QString value = attributes.value(QLatin1String("val")).toString();
                if (value == QLatin1String("superscript"))
                    format.setFontScript(Format::FontScriptSuper);
                else if (value == QLatin1String("subscript"))
                    format.setFontScript(Format::FontScriptSub);
            } else if (reader.name() == QLatin1String("scheme")) {
                format.setProperty(FormatPrivate::P_Font_Scheme, attributes.value(QLatin1String("val")).toString());
            }
        }
    }
    return format;
}

void RichString::write(QXmlStreamWriter &writer, const QString &name) const
{
    writer.writeStartElement(name); //is or si
    if (isRichString()) {
        //Rich text string
        for (int i=0; i<fragmentCount(); ++i) {
            writer.writeStartElement(QStringLiteral("r"));
            if (fragmentFormat(i).hasFontData()) {
                writer.writeStartElement(QStringLiteral("rPr"));
                writeFormat(writer, fragmentFormat(i));
                writer.writeEndElement();// rPr
            }
            writer.writeStartElement(QStringLiteral("t"));
            if (isSpaceReserveNeeded(fragmentText(i)))
                writer.writeAttribute(QStringLiteral("xml:space"), QStringLiteral("preserve"));
            writer.writeCharacters(fragmentText(i));
            writer.writeEndElement();// t
            writer.writeEndElement(); //r
        }
    } else {
        writer.writeStartElement(QStringLiteral("t"));
        QString pString = toPlainString();
        if (isSpaceReserveNeeded(pString))
            writer.writeAttribute(QStringLiteral("xml:space"), QStringLiteral("preserve"));
        writer.writeCharacters(pString);
        writer.writeEndElement();//t
    }
    writer.writeEndElement();//si
}

void RichString::read(QXmlStreamReader &reader, const QString &name)
{
    while (!reader.atEnd() && !(reader.name() == name && reader.tokenType() == QXmlStreamReader::EndElement)) {
        reader.readNextStartElement();
        if (reader.tokenType() == QXmlStreamReader::StartElement) {
            if (reader.name() == QLatin1String("r"))
                readRichStringPart(reader);
            else if (reader.name() == QLatin1String("t"))
                readPlainStringPart(reader);
        }
    }
}

/*!
    Returns true if this is rich text string.
 */
bool RichString::isRichString() const
{
    if (fragmentCount() > 1) //Is this enough??
        return true;
    return false;
}

/*!
    Returns true is this is an Null string.
 */
bool RichString::isNull() const
{
    return d->fragmentTexts.size() == 0;
}

bool RichString::isEmpty() const
{
    for (const auto& str : qAsConst(d->fragmentTexts)) {
        if (!str.isEmpty())
            return false;
    }

    return true;
}

/*!
    Converts to plain text string.
*/
QString RichString::toPlainString() const
{
    return d->fragmentTexts.join(QString());
}

/*!
  Converts to html string
*/
QString RichString::toHtml() const
{
    //: Todo
    return QString();
}

/*!
  Replaces the entire contents of the document
  with the given HTML-formatted text in the \a text string
*/
void RichString::setHtml(const QString &text)
{
    QTextDocument doc;
    doc.setHtml(text);
    QTextBlock block = doc.firstBlock();
    QTextBlock::iterator it;
    for (it = block.begin(); !(it.atEnd()); ++it) {
        QTextFragment textFragment = it.fragment();
        if (textFragment.isValid()) {
            Format fmt;
            fmt.setFont(textFragment.charFormat().font());
            fmt.setFontColor(textFragment.charFormat().foreground().color());
            addFragment(textFragment.text(), fmt);
        }
    }
}

/*!
    Returns fragment count.
 */
int RichString::fragmentCount() const
{
    return d->fragmentTexts.size();
}

/*!
    Appends a fragment with the given \a text and \a format.
 */
void RichString::addFragment(const QString &text, const Format &format)
{
    d->fragmentTexts.append(text);
    d->fragmentFormats.append(format);
    d->_dirty = true;
}

/*!
    Returns fragment text at the position \a index.
 */
QString RichString::fragmentText(int index) const
{
    if (index < 0 || index >= fragmentCount())
        return QString();

    return d->fragmentTexts[index];
}

/*!
    Returns fragment format at the position \a index.
 */
Format RichString::fragmentFormat(int index) const
{
    if (index < 0 || index >= fragmentCount())
        return Format();

    return d->fragmentFormats[index];
}

/*!
 * \internal
 */
QByteArray RichStringPrivate::idKey() const
{
    if (_dirty) {
        RichStringPrivate *rs = const_cast<RichStringPrivate *>(this);
        QByteArray bytes;
        if (fragmentTexts.size() == 1) {
            bytes = fragmentTexts[0].toUtf8();
        } else {
            //Generate a hash value base on QByteArray ?
            bytes.append("@@QtXlsxRichString=");
            for (int i=0; i<fragmentTexts.size(); ++i) {
                bytes.append("@Text");
                bytes.append(fragmentTexts[i].toUtf8());
                bytes.append("@Format");
                if (fragmentFormats[i].hasFontData())
                    bytes.append(fragmentFormats[i].fontKey());
            }
        }
        rs->_idKey = bytes;
        rs->_dirty = false;
    }

    return _idKey;
}

/*!
    Returns true if this string \a rs1 is equal to string \a rs2;
    otherwise returns false.
 */
bool operator==(const RichString &rs1, const RichString &rs2)
{
    if (rs1.fragmentCount() != rs2.fragmentCount())
        return false;

    return rs1.d->idKey() == rs2.d->idKey();
}

/*!
    Returns true if this string \a rs1 is not equal to string \a rs2;
    otherwise returns false.
 */
bool operator!=(const RichString &rs1, const RichString &rs2)
{
    if (rs1.fragmentCount() != rs2.fragmentCount())
        return true;

    return rs1.d->idKey() != rs2.d->idKey();
}

/*!
 * \internal
 */
bool operator<(const RichString &rs1, const RichString &rs2)
{
    return rs1.d->idKey() < rs2.d->idKey();
}

/*!
    \overload
    Returns true if this string \a rs1 is equal to string \a rs2;
    otherwise returns false.
 */
bool operator ==(const RichString &rs1, const QString &rs2)
{
    if (rs1.fragmentCount() == 1 && rs1.fragmentText(0) == rs2) //format == 0
        return true;

    return false;
}

/*!
    \overload
    Returns true if this string \a rs1 is not equal to string \a rs2;
    otherwise returns false.
 */
bool operator !=(const RichString &rs1, const QString &rs2)
{
    if (rs1.fragmentCount() == 1 && rs1.fragmentText(0) == rs2) //format == 0
        return false;

    return true;
}

/*!
    \overload
    Returns true if this string \a rs1 is equal to string \a rs2;
    otherwise returns false.
 */
bool operator ==(const QString &rs1, const RichString &rs2)
{
    return rs2 == rs1;
}

/*!
    \overload
    Returns true if this string \a rs1 is not equal to string \a rs2;
    otherwise returns false.
 */
bool operator !=(const QString &rs1, const RichString &rs2)
{
    return rs2 != rs1;
}

uint qHash(const RichString &rs, uint seed) Q_DECL_NOTHROW
{
   return qHash(rs.d->idKey(), seed);
}

#ifndef QT_NO_DEBUG_STREAM
QDebug operator<<(QDebug dbg, const RichString &rs)
{
    dbg.nospace() << "QXlsx::RichString(" << rs.d->fragmentTexts << ")";
    return dbg.space();
}
#endif

}
