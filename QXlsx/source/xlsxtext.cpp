#include <QtGlobal>
#include <QDebug>
#include <QTextDocument>
#include <QTextFragment>
#include <QTemporaryFile>

#include "xlsxtext.h"
#include "xlsxtext_p.h"
#include "xlsxformat_p.h"
#include "xlsxmain.h"
#include "xlsxutility_p.h"

QT_BEGIN_NAMESPACE_XLSX

class TextPrivate : public QSharedData
{
public:
    TextPrivate();
    TextPrivate(const TextPrivate &other);
    ~TextPrivate();

    Text::Type type = Text::Type::None;

    QList<Paragraph> paragraphs; //for RichString
    QString reference; // for Reference
    QString plainText; //for PlainString
    QStringList cashe;

    TextProperties textProperties; //element, required
    ListStyleProperties defaultParagraphProperties; //element, optional

    bool operator ==(const TextPrivate &other) const;
};



TextPrivate::TextPrivate()
{

}

TextPrivate::TextPrivate(const TextPrivate &other) : QSharedData(other),
    type{other.type}, paragraphs{other.paragraphs},
    reference{other.reference}, textProperties{other.textProperties},
    defaultParagraphProperties{other.defaultParagraphProperties}
{

}

TextPrivate::~TextPrivate()
{

}

bool TextPrivate::operator==(const TextPrivate &other) const
{
    if (type != other.type)
        return false;
    if (paragraphs != other.paragraphs)
        return false;
    if (reference != other.reference)
        return false;
    if (textProperties != other.textProperties)
        return false;
    if (defaultParagraphProperties != other.defaultParagraphProperties)
        return false;
    return true;
}

Text::Text()
{
}

Text::Text(const QString &text)
{
    d = new TextPrivate;
    d->type = Type::PlainText;
    addFragment(text, CharacterProperties());
}

Text::Text(Text::Type type)
{
    d = new TextPrivate;
    d->type = type;
}

Text::Text(const Text &other) : d(other.d)
{

}

Text::~Text()
{

}

void Text::setType(Type type)
{
    if (!d) d = new TextPrivate;
    d->type = type;

//    //adjust formatting
//    if (type != Type::RichText) {
//        d->fragmentFormats.clear();
//        QString s = d->fragmentTexts.join(QString());
//        d->fragmentTexts.clear();
//        d->fragmentTexts << s;
//    }
}

Text::Type Text::type() const
{
    return d->type;
}

void Text::setPlainString(const QString &s)
{
    if (!d) d = new TextPrivate;
    d->type = Type::PlainText;
//    d->paragraphs.clear();
//    d->reference.clear();
    d->plainText = s;
}

QString Text::toPlainString() const
{
    if (d)
        return d->plainText;
    return QString();
}

void Text::setHtml(const QString &text, bool detectHtml)
{
    if (!d) d = new TextPrivate;
    d->type = Type::RichText;
    d->paragraphs.clear();
    d->reference.clear();

    if (detectHtml) {
        //TODO: rewrite using our own parser
        QTextDocument doc;
        doc.setHtml(text);

        for (int i=0; i<doc.blockCount(); ++i) {
            auto block = doc.findBlockByNumber(i);
            //One Paragraph for each block
            Paragraph p;
            p.paragraphProperties = ParagraphProperties::from(block);

            for (auto it = block.begin(); !(it.atEnd()); ++it) {
                QTextFragment textFragment = it.fragment();
                if (textFragment.isValid()) {
                    QString s = textFragment.text();
                    if (s == QChar::LineSeparator) {
                        TextRun r(TextRun::Type::LineBreak);
                        p.textRuns << r;
                    }
                    else {
                        auto pr = CharacterProperties::from(textFragment.charFormat());
                        TextRun r(TextRun::Type::Regular, s, pr);
                        p.textRuns << r;
                    }
                }
            }
            d->paragraphs << p;
        }
    }
    else {
        Paragraph p;
        p.textRuns << TextRun(TextRun::Type::Regular, text, CharacterProperties());
        d->paragraphs << p;
    }
}

QString Text::toHtml() const
{
    //TODO: write
    return {};
}

void Text::setStringReference(const QString &ref)
{
    if (!d) d = new TextPrivate;
    d->type = Type::StringRef;
    d->reference = ref;
}

void Text::setStringCashe(const QStringList &cashe)
{
    if (!d) d = new TextPrivate;
    if (d->type != Type::StringRef) return;
    d->cashe = cashe;
}

QString Text::stringReference() const
{
    if (d && d->type == Type::StringRef)
        return d->reference;

    return {};
}

QStringList Text::stringCashe() const
{
    if (d && d->type == Type::StringRef)
        return d->cashe;
    return {};
}

bool Text::isRichString() const
{
    if (!d) return false;
    return d->type == Type::RichText;
}

bool Text::isPlainString() const
{
    if (!d) return false;
    return d->type == Type::PlainText;
}

bool Text::isStringReference() const
{
    if (!d) return false;
    return d->type == Type::StringRef;
}

//bool Text::isNull() const
//{
//    if (!d) return true;
//    switch (d->type) {
//        case Type::None: return true;
//        case Type::PlainText: return d->plainText.isNull();
//        case Type::RichText: return !d->textProperties.isValid() && d->paragraphs.isEmpty();
//        case Type::StringRef: return d->reference.isNull();
//    }

//    return true;
//}

//bool Text::isEmpty() const
//{
//    if (!d) return true;
//    switch (d->type) {
//        case Type::None: return true;
//        case Type::PlainText: return d->plainText.isEmpty();
//        case Type::RichText: return d->paragraphs.isEmpty();
//        case Type::StringRef: return d->reference.isEmpty();
//    }

//    return true;
//}

bool Text::isValid() const
{
    if (!d) return false;
    return true;
}

void Text::addFragment(const QString &text, const CharacterProperties &format)
{
    if (!d) d = new TextPrivate;
    if (d->paragraphs.isEmpty()) d->paragraphs << Paragraph();
    d->paragraphs.last().textRuns << TextRun(TextRun::Type::Regular, text, format);
}

QString Text::fragmentText(int index) const
{
    if (d) {
        if (d->paragraphs.isEmpty()) return {};
        if (index < 0 || index >= d->paragraphs.last().textRuns.count()) return {};
        return d->paragraphs.last().textRuns[index].text;
    }
    return QString();
}

CharacterProperties Text::fragmentFormat(int index) const
{
    if (d) {
        if (d->paragraphs.isEmpty()) return {};
        if (index < 0 || index >= d->paragraphs.last().textRuns.count()) return {};
        return d->paragraphs.last().textRuns[index].characterProperties.value_or(CharacterProperties());
    }
    return {};
}

void Text::read(QXmlStreamReader &reader, bool diveInto)
{
    if (!d) d = new TextPrivate;

    const auto &name = reader.name();

    if (diveInto) {
        while (!reader.atEnd()) {
            auto token = reader.readNext();
            if (token == QXmlStreamReader::StartElement) {
                if (reader.name() == QLatin1String("strRef")) {
                    readStringReference(reader);
                } else if (reader.name() == QLatin1String("rich")) {
                    readRichString(reader);
                } else if (reader.name() == QLatin1String("v")) {
                    readPlainString(reader);
                }
                else reader.skipCurrentElement();
            }
            else if (token == QXmlStreamReader::EndElement && reader.name() == name)
                break;
        }
    } else {
        if (reader.name() == QLatin1String("txPr"))
            readRichString(reader);
    }
}

void Text::write(QXmlStreamWriter &writer, const QString &name, bool diveInto) const
{
    if (!d) return;
//    if (d->type == Type::None) return;
    if (diveInto) {
        writer.writeStartElement(name);

        switch (d->type) {
        case Type::StringRef:
            writeStringReference(writer, QLatin1String("c:strRef"));
            break;
        case Type::PlainText:
            writePlainString(writer, "c:v");
            break;
        case Type::RichText:
            writeRichString(writer, "c:rich");
            break;
        default:
            break;
        }

        writer.writeEndElement(); //name
    } else {
        writeRichString(writer, name);
    }
}

Text &Text::operator=(const Text &other)
{
    this->d = other.d;
    return *this;
}

bool Text::operator ==(const Text &other) const
{
    if (d == other.d) return true;
    if (!d || !other.d) return false;
    return *this->d.constData() == *other.d.constData();
}

bool Text::operator !=(const Text &other) const
{
    return !(operator==(other));
}

void Text::readStringReference(QXmlStreamReader &reader)
{
    if (!d) d = new TextPrivate;
    d->type = Type::StringRef;

    const auto &name = reader.name();
    int count = 0;
    int idx = -1;
    QVector<QString> list;
    while (!reader.atEnd()) {
        auto token = reader.readNext();
        if (token == QXmlStreamReader::StartElement) {
            if (reader.name() == QLatin1String("f")) {
                d->reference = reader.readElementText();
            }
            else if (reader.name() == QLatin1String("ptCount")) {
                count = reader.attributes().value(QLatin1String("val")).toInt();
                list.resize(count);
            }
            else if (reader.name() == QLatin1String("pt")) {
                idx = reader.attributes().value(QLatin1String("idx")).toInt();
            }
            else if (reader.name() == QLatin1String("v")) {
                QString s = reader.readElementText();
                if (idx >=0 && idx < list.size()) list[idx] = s;
                idx = -1;
            }
            else reader.skipCurrentElement();
        }
        else if (token == QXmlStreamReader::EndElement && reader.name() == name)
            break;
    }
    d->cashe.clear();
    std::copy(list.constBegin(), list.constEnd(), std::back_inserter(d->cashe));
}

void Text::readRichString(QXmlStreamReader &reader)
{
//    <xsd:complexType name="CT_TextBody">
//        <xsd:sequence>
//          <xsd:element name="bodyPr" type="CT_TextBodyProperties" minOccurs="1" maxOccurs="1"/>
//          <xsd:element name="lstStyle" type="CT_TextListStyle" minOccurs="0" maxOccurs="1"/>
//          <xsd:element name="p" type="CT_TextParagraph" minOccurs="1" maxOccurs="unbounded"/>
//        </xsd:sequence>
//      </xsd:complexType>
    if (!d) d = new TextPrivate;
    d->type = Type::RichText;

    const auto &name = reader.name();
    while (!reader.atEnd()) {
        auto token = reader.readNext();
        if (token == QXmlStreamReader::StartElement) {
            if (reader.name() == QLatin1String("bodyPr")) {
                d->textProperties.read(reader);
            }
            else if (reader.name() == QLatin1String("lstStyle")) {
                d->defaultParagraphProperties.read(reader);
            }
            else if (reader.name() == QLatin1String("p")) {
                readParagraph(reader);
            }
            else reader.skipCurrentElement();
        }
        else if (token == QXmlStreamReader::EndElement && reader.name() == name)
            break;
    }

}

void Text::readPlainString(QXmlStreamReader &reader)
{
    if (!d) {
        d = new TextPrivate;
        d->type = Type::PlainText;
    }
    Q_ASSERT(reader.name() == QLatin1String("v"));

    d->plainText = reader.readElementText();
}

void Text::readParagraph(QXmlStreamReader &reader)
{
    if (!d) d = new TextPrivate;
    Paragraph p;
    p.read(reader);
    d->paragraphs << p;
}

void Text::writeParagraphs(QXmlStreamWriter &writer) const
{
    for (const auto &p: d->paragraphs)
        p.write(writer, QLatin1String("a:p"));
}

void Text::writeStringReference(QXmlStreamWriter &writer, const QString &name) const
{
    if (!d || d->reference.isEmpty())
        return;

    writer.writeStartElement(name);
    writer.writeTextElement(QLatin1String("c:f"), d->reference);
    if (!d->cashe.isEmpty()) {
        writer.writeStartElement(QLatin1String("c:strCashe"));
        writer.writeEmptyElement(QLatin1String("c:ptCount"));
        writer.writeAttribute(QLatin1String("val"), QString::number(d->cashe.size()));
        for (const QString &s: d->cashe) {
            writer.writeStartElement(QLatin1String("c:pt"));
            writer.writeTextElement(QLatin1String("c:v"), s);
            writer.writeEndElement(); //c:pt
        }
        writer.writeEndElement(); //c:strCashe
    }
    writer.writeEndElement(); //c:strRef
}

void Text::writeRichString(QXmlStreamWriter &writer, const QString &name) const
{
    if (!d) return;

    writer.writeStartElement(name);
    d->textProperties.write(writer);
    d->defaultParagraphProperties.write(writer, QLatin1String("a:lstStyle"));

    writeParagraphs(writer);

    writer.writeEndElement(); // name
}

void Text::writePlainString(QXmlStreamWriter &writer, const QString &name) const
{
    if (!d) return;
    if (d->plainText.isEmpty()) return;

    writer.writeTextElement(name, d->plainText);
}

QXlsx::Text::operator QVariant() const
{
    const auto& cref
#if QT_VERSION >= 0x060000 // Qt 6.0 or over
        = QMetaType::fromType<Text>();
#else
        = qMetaTypeId<Text>() ;
#endif
    return QVariant(cref, this);
}

bool TextProperties::isValid() const
{
    //TODO: переписать
    if (spcFirstLastPara.has_value()) return true;
    if (verticalOverflow.has_value()) return true;
    if (horizontalOverflow.has_value()) return true;
    if (verticalOrientation.has_value()) return true;
    if (rotation.has_value()) return true;
    if (wrap.has_value()) return true;
    if (leftInset.isValid()) return true;
    if (rightInset.isValid()) return true;
    if (topInset.isValid()) return true;
    if (bottomInset.isValid()) return true;
    if (columnCount.has_value()) return true;
    if (columnSpace.isValid()) return true;
    if (columnsRtl.has_value()) return true;
    if (fromWordArt.has_value()) return true;
    if (anchor.has_value()) return true;
    if (anchorCentering.has_value()) return true;
    if (forceAntiAlias.has_value()) return true;
    if (upright.has_value()) return true;
    if (compatibleLineSpacing.has_value()) return true;

    return false;
}

void TextProperties::read(QXmlStreamReader &reader)
{
    const auto &name = reader.name();
    const auto &a = reader.attributes();
    parseAttribute(a, QLatin1String("rot"), rotation);
    parseAttributeBool(a, QLatin1String("spcFirstLastPara"), spcFirstLastPara);

    if (a.hasAttribute(QLatin1String("vertOverflow"))) {
        const auto &val = a.value(QLatin1String("vertOverflow"));
        if (val == QLatin1String("overflow")) verticalOverflow = TextProperties::OverflowType::Overflow;
        if (val == QLatin1String("ellipsis")) verticalOverflow = TextProperties::OverflowType::Ellipsis;
        if (val == QLatin1String("clip")) verticalOverflow = TextProperties::OverflowType::Clip;
    }
    if (a.hasAttribute(QLatin1String("horzOverflow"))) {
        const auto &val = a.value(QLatin1String("horzOverflow"));
        if (val == QLatin1String("overflow")) horizontalOverflow = TextProperties::OverflowType::Overflow;
        if (val == QLatin1String("ellipsis")) horizontalOverflow = TextProperties::OverflowType::Ellipsis;
        if (val == QLatin1String("clip")) horizontalOverflow = TextProperties::OverflowType::Clip;
    }
    if (a.hasAttribute(QLatin1String("vert"))) {
        const auto &val = a.value(QLatin1String("vert"));
        if (val == QLatin1String("horz")) verticalOrientation = TextProperties::VerticalType::Horizontal;
        if (val == QLatin1String("vert")) verticalOrientation = TextProperties::VerticalType::Vertical;
        if (val == QLatin1String("vert270")) verticalOrientation = TextProperties::VerticalType::Vertical270;
        if (val == QLatin1String("wordArtVert")) verticalOrientation = TextProperties::VerticalType::WordArtVertical;
        if (val == QLatin1String("eaVert")) verticalOrientation = TextProperties::VerticalType::EastAsianVertical;
        if (val == QLatin1String("mongolianVert")) verticalOrientation = TextProperties::VerticalType::MongolianVertical;
        if (val == QLatin1String("wordArtVertRtl")) verticalOrientation = TextProperties::VerticalType::WordArtVerticalRtl;
    }
    if (a.hasAttribute(QLatin1String("wrap"))) {
        const auto &val = a.value(QLatin1String("wrap"));
        wrap = val == QLatin1String("square"); //"none" -> false
    }
    parseAttribute(a, QLatin1String("lIns"), leftInset);
    parseAttribute(a, QLatin1String("rIns"), rightInset);
    parseAttribute(a, QLatin1String("tIns"), topInset);
    parseAttribute(a, QLatin1String("bIns"), bottomInset);
    parseAttributeInt(a, QLatin1String("numCol"), columnCount);
    parseAttribute(a, QLatin1String("spcCol"), columnSpace);
    parseAttributeBool(a, QLatin1String("rtlCol"), columnsRtl);
    parseAttributeBool(a, QLatin1String("fromWordArt"), fromWordArt);
    parseAttributeBool(a, QLatin1String("anchorCtr"), anchorCentering);
    parseAttributeBool(a, QLatin1String("forceAA"), forceAntiAlias);
    parseAttributeBool(a, QLatin1String("upright"), upright);
    parseAttributeBool(a, QLatin1String("compatLnSpc"), compatibleLineSpacing);

    if (a.hasAttribute(QLatin1String("anchor"))) {
        const auto &val = a.value(QLatin1String("anchor"));
        if (val == QLatin1String("b")) anchor = TextProperties::Anchor::Bottom;
        if (val == QLatin1String("t")) anchor = TextProperties::Anchor::Top;
        if (val == QLatin1String("ctr")) anchor = TextProperties::Anchor::Center;
        if (val == QLatin1String("just")) anchor = TextProperties::Anchor::Justified;
        if (val == QLatin1String("dist")) anchor = TextProperties::Anchor::Distributed;
    }

    while (!reader.atEnd()) {
        auto token = reader.readNext();
        if (token == QXmlStreamReader::StartElement) {
            if (reader.name() == QLatin1String("prstTxWarp")) {
                textShape->read(reader);
            }
            else if (reader.name() == QLatin1String("noAutofit")) {
                textAutofit = TextProperties::TextAutofit::NoAutofit;
            }
            else if (reader.name() == QLatin1String("spAutoFit")) {
                textAutofit = TextProperties::TextAutofit::ShapeAutofit;
            }
            else if (reader.name() == QLatin1String("normAutofit")) {
                textAutofit = TextProperties::TextAutofit::NormalAutofit;
                const auto &a = reader.attributes();
                parseAttributePercent(a, QLatin1String("fontScale"), fontScale);
                parseAttributePercent(a, QLatin1String("lnSpcReduction"), lineSpaceReduction);
            }
            else if (reader.name() == QLatin1String("scene3d")) {
                scene3D->read(reader);
            }
            else if (reader.name() == QLatin1String("sp3d")) {
                text3D->read(reader);
            }
            else if (reader.name() == QLatin1String("flatTx")) {
                z = Coordinate::create(reader.attributes().value(QLatin1String("z")).toString());
            }
            else reader.skipCurrentElement();
        }
        else if (token == QXmlStreamReader::EndElement && reader.name() == name)
            break;
    }
}

void TextProperties::write(QXmlStreamWriter &writer) const
{
    if (!isValid()) {
        writer.writeEmptyElement(QLatin1String("a:bodyPr"));
        return;
    }
    writer.writeStartElement(QLatin1String("a:bodyPr"));

    writeAttribute(writer, QLatin1String("spcFirstLastPara"), spcFirstLastPara);
    writeAttribute(writer, QLatin1String("rtlCol"), columnsRtl);
    writeAttribute(writer, QLatin1String("fromWordArt"), fromWordArt);
    writeAttribute(writer, QLatin1String("anchorCtr"), anchorCentering);
    writeAttribute(writer, QLatin1String("forceAA"), forceAntiAlias);
    writeAttribute(writer, QLatin1String("upright"), upright);
    writeAttribute(writer, QLatin1String("compatLnSpc"), compatibleLineSpacing);
    if (rotation.has_value())
        writer.writeAttribute(QLatin1String("rot"), rotation->toString());
    if (verticalOverflow.has_value()) {
        switch (verticalOverflow.value()) {
            case TextProperties::OverflowType::Overflow:
                writer.writeAttribute(QLatin1String("vertOverflow"), QLatin1String("overflow"));
                break;
            case TextProperties::OverflowType::Ellipsis:
                writer.writeAttribute(QLatin1String("vertOverflow"), QLatin1String("ellipsis"));
                break;
            case TextProperties::OverflowType::Clip:
                writer.writeAttribute(QLatin1String("vertOverflow"), QLatin1String("clip"));
                break;
        }
    }
    if (horizontalOverflow.has_value()) {
        switch (horizontalOverflow.value()) {
            case TextProperties::OverflowType::Overflow:
                writer.writeAttribute(QLatin1String("horzOverflow"), QLatin1String("overflow"));
                break;
            case TextProperties::OverflowType::Ellipsis:
                writer.writeAttribute(QLatin1String("horzOverflow"), QLatin1String("ellipsis"));
                break;
            case TextProperties::OverflowType::Clip:
                writer.writeAttribute(QLatin1String("horzOverflow"), QLatin1String("clip"));
                break;
        }
    }
    if (verticalOrientation.has_value()) {
        switch (verticalOrientation.value()) {
            case TextProperties::VerticalType::Horizontal:
                writer.writeAttribute(QLatin1String("vert"), QLatin1String("horz"));
                break;
            case TextProperties::VerticalType::Vertical:
                writer.writeAttribute(QLatin1String("vert"), QLatin1String("vert"));
                break;
            case TextProperties::VerticalType::Vertical270:
                writer.writeAttribute(QLatin1String("vert"), QLatin1String("vert270"));
                break;
            case TextProperties::VerticalType::WordArtVertical:
                writer.writeAttribute(QLatin1String("vert"), QLatin1String("wordArtVert"));
                break;
            case TextProperties::VerticalType::EastAsianVertical:
                writer.writeAttribute(QLatin1String("vert"), QLatin1String("eaVert"));
                break;
            case TextProperties::VerticalType::MongolianVertical:
                writer.writeAttribute(QLatin1String("vert"), QLatin1String("mongolianVert"));
                break;
            case TextProperties::VerticalType::WordArtVerticalRtl:
                writer.writeAttribute(QLatin1String("vert"), QLatin1String("wordArtVertRtl"));
                break;
        }
    }
    if (wrap.has_value())
        writer.writeAttribute(QLatin1String("wrap"),
                              wrap.value() ? QLatin1String("square") : QLatin1String("none"));
    if (leftInset.isValid())
        writer.writeAttribute(QLatin1String("lIns"), leftInset.toString());
    if (rightInset.isValid())
        writer.writeAttribute(QLatin1String("rIns"), rightInset.toString());
    if (topInset.isValid())
        writer.writeAttribute(QLatin1String("tIns"), topInset.toString());
    if (bottomInset.isValid())
        writer.writeAttribute(QLatin1String("bIns"), bottomInset.toString());
    if (columnCount.has_value())
        writer.writeAttribute(QLatin1String("numCol"), QString::number(columnCount.value()));
    if (columnSpace.isValid())
        writer.writeAttribute(QLatin1String("spcCol"), columnSpace.toString());

    if (anchor.has_value()) {
        switch (anchor.value()) {
            case TextProperties::Anchor::Bottom:
                writer.writeAttribute(QLatin1String("anchor"), QLatin1String("b")); break;
            case TextProperties::Anchor::Top:
                writer.writeAttribute(QLatin1String("anchor"), QLatin1String("t")); break;
            case TextProperties::Anchor::Center:
                writer.writeAttribute(QLatin1String("anchor"), QLatin1String("ctr")); break;
            case TextProperties::Anchor::Justified:
                writer.writeAttribute(QLatin1String("anchor"), QLatin1String("just")); break;
            case TextProperties::Anchor::Distributed:
                writer.writeAttribute(QLatin1String("anchor"), QLatin1String("dist")); break;
        }
    }
    if (textShape.has_value())
        textShape.value().write(writer, QLatin1String("a:prstTxWarp"));

    if (textAutofit.has_value()) {
        switch (textAutofit.value()) {
            case TextProperties::TextAutofit::NoAutofit:
                writer.writeEmptyElement(QLatin1String("a:noAutofit")); break;
            case TextProperties::TextAutofit::NormalAutofit:
                writer.writeEmptyElement(QLatin1String("a:normAutofit"));
                if (fontScale.has_value())
                    writer.writeAttribute(QLatin1String("fontScale"), toST_Percent(fontScale.value()));
                if (lineSpaceReduction.has_value())
                    writer.writeAttribute(QLatin1String("lnSpcReduction"), toST_Percent(lineSpaceReduction.value()));
                break;
            case TextProperties::TextAutofit::ShapeAutofit:
                writer.writeEmptyElement(QLatin1String("a:spAutofit")); break;
        }
    }
    if (text3D.has_value())
        text3D.value().write(writer, QLatin1String("a:sp3d"));
    else if (z.isValid()) {
        writer.writeEmptyElement(QLatin1String("a:flatTx"));
        writer.writeAttribute(QLatin1String("z"), z.toString());
    }

    writer.writeEndElement();
}

bool TextProperties::operator ==(const TextProperties &other) const
{
    if (spcFirstLastPara != other.spcFirstLastPara) return false;
    if (verticalOverflow != other.verticalOverflow) return false;
    if (horizontalOverflow != other.horizontalOverflow) return false;
    if (verticalOrientation != other.verticalOrientation) return false;
    if (rotation != other.rotation) return false;
    if (wrap != other.wrap) return false;
    if (leftInset != other.leftInset) return false;
    if (rightInset != other.rightInset) return false;
    if (topInset != other.topInset) return false;
    if (bottomInset != other.bottomInset) return false;
    if (columnCount != other.columnCount) return false;
    if (columnSpace != other.columnSpace) return false;
    if (columnsRtl != other.columnsRtl) return false;
    if (fromWordArt != other.fromWordArt) return false;
    if (anchor != other.anchor) return false;
    if (anchorCentering != other.anchorCentering) return false;
    if (forceAntiAlias != other.forceAntiAlias) return false;
    if (upright != other.upright) return false;
    if (compatibleLineSpacing != other.compatibleLineSpacing) return false;
    if (textShape != other.textShape) return false;
    if (textAutofit != other.textAutofit) return false;
    if (fontScale != other.fontScale) return false;
    if (lineSpaceReduction != other.lineSpaceReduction) return false;
    if (scene3D != other.scene3D) return false;
    if (text3D != other.text3D) return false;
    if (z != other.z) return false;
    return true;
}
bool TextProperties::operator !=(const TextProperties &other) const
{
    return !operator==(other);
}

Size::Size()
{

}

Size::Size(uint points)
{
    val = QVariant::fromValue<uint>(points);
}

Size::Size(double percents)
{
    val = QVariant::fromValue<double>(percents);
}

int Size::toPoints() const
{
    return val.toUInt();
}

double Size::toPercents() const
{
    return val.toDouble();
}

QString Size::toString() const
{
    if (val.userType() == QMetaType::UInt) return QString::number(val.toUInt());
    if (val.userType() == QMetaType::Double) return QString::number(val.toDouble())+'%';
    return QString();
}


bool Size::isValid() const
{
    return val.isValid();
}

bool Size::isPoints() const
{
    return val.userType() == QMetaType::UInt;
}

bool Size::isPercents() const
{
    return val.userType() == QMetaType::Double;
}

void Size::setPoints(int points)
{
    val = QVariant::fromValue<uint>(points);
}

void Size::setPercents(double percents)
{
    val = QVariant::fromValue<double>(percents);
}

void Size::read(QXmlStreamReader &reader)
{
    //    <xsd:complexType name="CT_TextSpacing">
    //      <xsd:choice>
    //        <xsd:element name="spcPct" type="CT_TextSpacingPercent"/>
    //        <xsd:element name="spcPts" type="CT_TextSpacingPoint"/>
    //      </xsd:choice>
    //    </xsd:complexType>

    //    <xsd:complexType name="CT_TextSpacingPoint">
    //      <xsd:attribute name="val" type="ST_TextSpacingPoint" use="required"/>
    //    </xsd:complexType>
    //    <xsd:simpleType name="ST_TextSpacingPoint">
    //        <xsd:restriction base="xsd:int">
    //          <xsd:minInclusive value="0"/>
    //          <xsd:maxInclusive value="158400"/>
    //        </xsd:restriction>
    //      </xsd:simpleType>
    //    <xsd:simpleType name="ST_TextSpacingPercentOrPercentString">
    //      <xsd:union memberTypes="s:ST_Percentage"/>
    //    </xsd:simpleType>
    //    <xsd:complexType name="CT_TextSpacingPercent">
    //      <xsd:attribute name="val" type="ST_TextSpacingPercentOrPercentString" use="required"/>
    //    </xsd:complexType>

    reader.readNextStartElement();
    if (reader.name() == QLatin1String("spcPct")) {
        std::optional<double> v;
        parseAttributePercent(reader.attributes(), QLatin1String("val"), v);
        if (v.has_value()) val = v.value();
    }
    if (reader.name() == QLatin1String("spcPts")) {
        int v;
        parseAttributeInt(reader.attributes(), QLatin1String("val"), v);
        val = v;
    }
}

void Size::write(QXmlStreamWriter &writer, const QString &name) const
{
    writer.writeStartElement(name);
    if (val.userType() == QMetaType::Int) {
        writer.writeEmptyElement(QLatin1String("spcPts"));
        writer.writeAttribute(QLatin1String("val"), QString::number(val.toInt()));
    }
    if (val.userType() == QMetaType::Double) {
        writer.writeEmptyElement(QLatin1String("spcPct"));
        writer.writeAttribute(QLatin1String("val"), toST_Percent(val.toDouble()));
    }
    writer.writeEndElement();
}

void ParagraphProperties::read(QXmlStreamReader &reader)
{
    const auto &name = reader.name();

    const auto &a = reader.attributes();

    parseAttribute(a, QLatin1String("marL"), leftMargin);
    parseAttribute(a, QLatin1String("marR"), rightMargin);
    parseAttributeInt(a, QLatin1String("lvl"), indentLevel);
    parseAttribute(a, QLatin1String("indent"), indent);

    if (a.hasAttribute(QLatin1String("algn"))) {
        TextAlign al; fromString(a.value(QLatin1String("algn")).toString(), al);
        align = al;
    }
    parseAttribute(a, QLatin1String("defTabSz"), defaultTabSize);
    parseAttributeBool(a, QLatin1String("rtl"), rtl);
    parseAttributeBool(a, QLatin1String("eaLnBrk"), eastAsianLineBreak);

    if (a.hasAttribute(QLatin1String("fontAlgn"))) {
        FontAlign f; fromString(a.value(QLatin1String("fontAlgn")).toString(), f);
        fontAlign = f;
    }
    parseAttributeBool(a, QLatin1String("latinLnBrk"), latinLineBreak);
    parseAttributeBool(a, QLatin1String("hangingPunct"), hangingPunctuation);

    while (!reader.atEnd()) {
        auto token = reader.readNext();
        if (token == QXmlStreamReader::StartElement) {
            if (reader.name() == QLatin1String("lnSpc")) {
                Size ls; ls.read(reader);
                lineSpacing = ls;
            }
            else if (reader.name() == QLatin1String("spcBef")) {
                Size ls; ls.read(reader);
                spacingBefore = ls;
            }
            else if (reader.name() == QLatin1String("spcAft")) {
                Size ls; ls.read(reader);
                spacingAfter = ls;
            }
            else if (reader.name() == QLatin1String("buClrTx")) {
                // bullet color follows text
            }
            else if (reader.name() == QLatin1String("buClr")) {
                reader.readNextStartElement();
                bulletColor = Color();
                bulletColor->read(reader);
            }
            else if (reader.name() == QLatin1String("buSzTx")) {
                // bullet size follows text
            }
            else if (reader.name() == QLatin1String("buSzPct")) {
                std::optional<double> v;
                parseAttributePercent(reader.attributes(), QLatin1String("val"), v);
                if (v.has_value()) bulletSize = Size(v.value());
            }
            else if (reader.name() == QLatin1String("buSzPts")) {
                uint v;
                parseAttributeUInt(reader.attributes(), QLatin1String("val"), v);
                bulletSize = Size(v);
            }
            else if (reader.name() == QLatin1String("buFontTx")) {
                // bullet font follows text
            }
            else if (reader.name() == QLatin1String("buFont")) {
                Font f; f.read(reader);
                bulletFont = f;
            }
            else if (reader.name() == QLatin1String("buNone")) {
                bulletType = BulletType::None;
            }
            else if (reader.name() == QLatin1String("buAutoNum")) {
                bulletType = BulletType::AutoNumber;
                const auto &a = reader.attributes();
                const QString &type = a.value(QLatin1String("type")).toString();
                fromString(type, bulletAutonumberType);
                if (a.hasAttribute(QLatin1String("startAt")))
                    bulletAutonumberStart = a.value(QLatin1String("startAt")).toInt();
            }
            else if (reader.name() == QLatin1String("buChar")) {
                bulletType = BulletType::Char;
                bulletChar = reader.attributes().value(QLatin1String("char")).toString();
            }
            else if (reader.name() == QLatin1String("buBlip")) {
                //TODO: picture bullet type
                bulletType = BulletType::Blip;
            }
            else if (reader.name() == QLatin1String("tabLst")) {
                readTabStops(reader);
            }
            else if (reader.name() == QLatin1String("defRPr")) {
                CharacterProperties p;
                p.read(reader);
                defaultTextCharacterProperties = p;
            }
            else reader.skipCurrentElement();
        }
        else if (token == QXmlStreamReader::EndElement && reader.name() == name)
            break;
    }
}

void ParagraphProperties::write(QXmlStreamWriter &writer, const QString &name) const
{
    if (!isValid()) {
        writer.writeEmptyElement(name);
        return;
    }

    writer.writeStartElement(name);
    writeAttribute(writer, QLatin1String("marL"), leftMargin.toString());
    writeAttribute(writer, QLatin1String("marR"), rightMargin.toString());
    writeAttribute(writer, QLatin1String("lvl"), indentLevel);
    writeAttribute(writer, QLatin1String("indent"), indent.toString());
    if (align.has_value()) {
        QString s; toString(align.value(), s); writer.writeAttribute(QLatin1String("algn"), s);
    }
    writeAttribute(writer, QLatin1String("defTabSz"), defaultTabSize.toString());
    writeAttribute(writer, QLatin1String("rtl"), rtl);
    writeAttribute(writer, QLatin1String("eaLnBrk"), eastAsianLineBreak);
    if (fontAlign.has_value()) {
        QString s; toString(fontAlign.value(), s); writer.writeAttribute(QLatin1String("fontAlgn"), s);
    }
    writeAttribute(writer, QLatin1String("latinLnBrk"), latinLineBreak);
    writeAttribute(writer, QLatin1String("hangingPunct"), hangingPunctuation);

    if (lineSpacing.has_value())
        lineSpacing->write(writer, QLatin1String("a:lnSpc"));
    if (spacingBefore.has_value())
        spacingBefore->write(writer, QLatin1String("a:spcBef"));
    if (spacingAfter.has_value())
        spacingAfter->write(writer, QLatin1String("a:spcAft"));
    if (bulletColor.has_value()) {
        if (bulletColor->isValid()) {
            writer.writeStartElement(QLatin1String("a:buClr"));
            bulletColor->write(writer);
            writer.writeEndElement();
        }
        else writer.writeEmptyElement(QLatin1String("a:buClrTx"));
    }
    if (bulletSize.has_value()) {
        if (bulletSize->isValid()) {
            if (bulletSize->isPoints()) writeEmptyElement(writer, QLatin1String("a:buSzPts"), bulletSize->toString());
            if (bulletSize->isPercents()) writeEmptyElement(writer, QLatin1String("a:buSzPct"), bulletSize->toString());
        }
        else writer.writeEmptyElement(QLatin1String("a:buSzTx"));
    }
    if (bulletFont.has_value()) {
        if (bulletFont->isValid()) bulletFont->write(writer, QLatin1String("a:buFont"));
        else writer.writeEmptyElement(QLatin1String("a:buFontTx"));
    }
    if (bulletType.has_value()) {
        switch (bulletType.value()) {
            case BulletType::None: writer.writeEmptyElement(QLatin1String("a:buNone")); break;
            case BulletType::Char: {
                writer.writeEmptyElement(QLatin1String("a:buChar"));
                writer.writeAttribute(QLatin1String("char"), bulletChar);
                break;
            }
            case BulletType::AutoNumber: {
                writer.writeEmptyElement(QLatin1String("a:buAutoNum"));
                QString s; toString(bulletAutonumberType, s);
                writer.writeAttribute(QLatin1String("type"), s);
                writeAttribute(writer, QLatin1String("startAt"), bulletAutonumberStart);
                break;
            }
            case BulletType::Blip: {
                //TODO: picture bullet type
                break;
            }
        }
    }
    writeTabStops(writer, QLatin1String("a:tabLst"));
    if (defaultTextCharacterProperties.has_value())
        defaultTextCharacterProperties->write(writer, QLatin1String("a:defRPr"));

    writer.writeEndElement();
}

bool ListStyleProperties::isEmpty() const
{
    return vals.isEmpty();
}

ParagraphProperties ListStyleProperties::defaultParagraphProperties() const
{
    if (vals.size()>0) return vals[0];
    return {};
}

ParagraphProperties &ListStyleProperties::defaultParagraphProperties()
{
    if (vals.isEmpty()) vals << ParagraphProperties();
    return vals[0];
}

void ListStyleProperties::setDefaultParagraphProperties(const ParagraphProperties &properties)
{
    if (vals.isEmpty()) vals << properties;
    else vals[0] = properties;
}

ParagraphProperties ListStyleProperties::value(int level) const
{
    if (vals.size()>level && level > 0) return vals[level];
    return {};
}

void ListStyleProperties::read(QXmlStreamReader &reader)
{
    const auto &name = reader.name();

    while (!reader.atEnd()) {
        auto token = reader.readNext();
        if (token == QXmlStreamReader::StartElement) {
            if (reader.name() == QLatin1String("defPPr")) {
                if (vals.size() < 1) vals.resize(1);
                ParagraphProperties p;
                p.read(reader);
                vals[0] = p;
            }
            else if (reader.name() == QLatin1String("lvl1pPr")) {
                if (vals.size() < 2) vals.resize(2);
                ParagraphProperties p;
                p.read(reader);
                vals[1] = p;
            }
            else if (reader.name() == QLatin1String("lvl2pPr")) {
                if (vals.size() < 3) vals.resize(3);
                ParagraphProperties p;
                p.read(reader);
                vals[2] = p;
            }
            else if (reader.name() == QLatin1String("lvl3pPr")) {
                if (vals.size() < 4) vals.resize(4);
                ParagraphProperties p;
                p.read(reader);
                vals[3] = p;
            }
            else if (reader.name() == QLatin1String("lvl4pPr")) {
                if (vals.size() < 5) vals.resize(5);
                ParagraphProperties p;
                p.read(reader);
                vals[4] = p;
            }
            else if (reader.name() == QLatin1String("lvl5pPr")) {
                if (vals.size() < 6) vals.resize(6);
                ParagraphProperties p;
                p.read(reader);
                vals[5] = p;
            }
            else if (reader.name() == QLatin1String("lvl6pPr")) {
                if (vals.size() < 7) vals.resize(7);
                ParagraphProperties p;
                p.read(reader);
                vals[6] = p;
            }
            else if (reader.name() == QLatin1String("lvl7pPr")) {
                if (vals.size() < 8) vals.resize(8);
                ParagraphProperties p;
                p.read(reader);
                vals[7] = p;
            }
            else if (reader.name() == QLatin1String("lvl8pPr")) {
                if (vals.size() < 9) vals.resize(9);
                ParagraphProperties p;
                p.read(reader);
                vals[8] = p;
            }
            else if (reader.name() == QLatin1String("lvl9pPr")) {
                if (vals.size() < 10) vals.resize(10);
                ParagraphProperties p;
                p.read(reader);
                vals[9] = p;
            }
            else reader.skipCurrentElement();
        }
        else if (token == QXmlStreamReader::EndElement && reader.name() == name)
            break;
    }
}

void ListStyleProperties::write(QXmlStreamWriter &writer, const QString &name) const
{
    if (vals.isEmpty()) {
        writer.writeEmptyElement(name);
        return;
    }

    writer.writeStartElement(name);
    if (vals.size()>0) vals[0].write(writer, QLatin1String("defPPr"));
    if (vals.size()>1) vals[1].write(writer, QLatin1String("lvl1pPr"));
    if (vals.size()>2) vals[2].write(writer, QLatin1String("lvl2pPr"));
    if (vals.size()>3) vals[3].write(writer, QLatin1String("lvl3pPr"));
    if (vals.size()>4) vals[4].write(writer, QLatin1String("lvl4pPr"));
    if (vals.size()>5) vals[5].write(writer, QLatin1String("lvl5pPr"));
    if (vals.size()>6) vals[6].write(writer, QLatin1String("lvl6pPr"));
    if (vals.size()>7) vals[7].write(writer, QLatin1String("lvl7pPr"));
    if (vals.size()>8) vals[8].write(writer, QLatin1String("lvl8pPr"));
    if (vals.size()>9) vals[9].write(writer, QLatin1String("lvl9pPr"));
    writer.writeEndElement();
}

bool ListStyleProperties::operator==(const ListStyleProperties &other) const
{
    return vals == other.vals;
}
bool ListStyleProperties::operator!=(const ListStyleProperties &other) const
{
    return vals != other.vals;
}

void ParagraphProperties::readTabStops(QXmlStreamReader &reader)
{
    const auto &name = reader.name();
    while (!reader.atEnd()) {
        auto token = reader.readNext();
        if (token == QXmlStreamReader::StartElement) {
            if (reader.name() == QLatin1String("tab")) {
                const auto &a = reader.attributes();
                QPair<Coordinate, TabAlign> p;
                if (a.hasAttribute(QLatin1String("pos"))) p.first = Coordinate::create(a.value(QLatin1String("pos")));
                if (a.hasAttribute(QLatin1String("algn"))) {
                    TabAlign t;
                    fromString(a.value(QLatin1String("algn")).toString(), t);
                    p.second = t;
                    tabStops << p;
                }
            }
            else reader.skipCurrentElement();
        }
        else if (token == QXmlStreamReader::EndElement && reader.name() == name)
            break;
    }
}

void ParagraphProperties::writeTabStops(QXmlStreamWriter &writer, const QString &name) const
{
    if (tabStops.isEmpty()) return;
    writer.writeStartElement(name); //
    for (const auto &p: tabStops) {
        writer.writeEmptyElement(QLatin1String("a:tab"));
        writer.writeAttribute(QLatin1String("pos"), p.first.toString());
        QString s;
        toString(p.second, s);
        writer.writeAttribute(QLatin1String("algn"), s);
    }
    writer.writeEndElement();
}

ParagraphProperties ParagraphProperties::from(const QTextBlock &block)
{
    static int i=1;
    qDebug()<<"block"<<i++;
    qDebug()<<"paragraph"<<block.blockFormat().properties();
    qDebug()<<"chars"<<block.charFormat().properties();
    for (auto g: block.textFormats())
        qDebug()<<g.format.properties();
    ParagraphProperties pr;
    const auto paragraphFormat = block.blockFormat();
    if (!paragraphFormat.isEmpty()) {
        if (paragraphFormat.hasProperty(QTextFormat::BlockLeftMargin))
            pr.leftMargin = Coordinate(paragraphFormat.leftMargin());
        if (paragraphFormat.hasProperty(QTextFormat::BlockRightMargin))
            pr.rightMargin = Coordinate(paragraphFormat.rightMargin());
        if (paragraphFormat.hasProperty(QTextFormat::TextIndent)) {
            Coordinate c; c.setPoints(paragraphFormat.textIndent());
            pr.indent = c;
            pr.indentLevel = paragraphFormat.indent();
        }
        if (paragraphFormat.hasProperty(QTextFormat::BlockAlignment)) {
            if (paragraphFormat.alignment().testFlag(Qt::AlignLeft)) pr.align = TextAlign::Left;
            if (paragraphFormat.alignment().testFlag(Qt::AlignRight)) pr.align = TextAlign::Right;
            if (paragraphFormat.alignment().testFlag(Qt::AlignHCenter)) pr.align = TextAlign::Center;
            if (paragraphFormat.alignment().testFlag(Qt::AlignJustify)) pr.align = TextAlign::Justify;
        }
        //Coordinate defaultTabSize;
        if (paragraphFormat.hasProperty(QTextFormat::LayoutDirection))
            pr.rtl = paragraphFormat.layoutDirection() == Qt::RightToLeft;
        //std::optional<bool> eastAsianLineBreak;
        //std::optional<bool> latinLineBreak;
        //    std::optional<bool> hangingPunctuation;
        //    std::optional<Size> lineSpacing;
        pr.spacingBefore = Size(paragraphFormat.topMargin());
        pr.spacingAfter = Size(paragraphFormat.bottomMargin());

        //    std::optional<Color> bulletColor;
        //    std::optional<Size> bulletSize;
        //    Font bulletFont;
        //    std::optional<BulletType> bulletType;
        //    BulletAutonumberType bulletAutonumberType; //makes sense only if bulletType == Autonumber
        //    std::optional<int> bulletAutonumberStart;
        //    QString bulletChar; //makes sense only if bulletType == Char

        //    std::optional<TextCharacterProperties> defaultTextCharacterProperties;

        if (paragraphFormat.hasProperty(QTextFormat::TabPositions)) {
            auto tabPositions = paragraphFormat.tabPositions();
            for (auto &tab : tabPositions) {
                TabAlign alg = TabAlign::Left;
                if (tab.type == QTextOption::RightTab) alg = TabAlign::Right;
                if (tab.type == QTextOption::CenterTab) alg = TabAlign::Center;
                auto c = Coordinate(tab.position);
                pr.tabStops << qMakePair(c, alg);
            }
        }
    }
    const auto &charFormat = block.charFormat();
    if (!charFormat.isEmpty()) {
        if (paragraphFormat.hasProperty(QTextFormat::TextVerticalAlignment))
            switch (charFormat.verticalAlignment()) {
                case QTextCharFormat::AlignNormal: pr.fontAlign = FontAlign::Auto;
                    break;
                case QTextCharFormat::AlignSuperScript:
                    break;
                case QTextCharFormat::AlignSubScript:
                    break;
                case QTextCharFormat::AlignMiddle:  pr.fontAlign = FontAlign::Center;
                    break;
                case QTextCharFormat::AlignTop: pr.fontAlign = FontAlign::Top;
                    break;
                case QTextCharFormat::AlignBottom:  pr.fontAlign = FontAlign::Bottom;
                    break;
                case QTextCharFormat::AlignBaseline:  pr.fontAlign = FontAlign::Base;
                    break;
            }
    }
    return pr;
}

bool ParagraphProperties::isValid() const
{
    //TODO:
//    if (d) return true;
//    return false;

    if (leftMargin.isValid()) return true;
    if (rightMargin.isValid()) return true;
    if (indent.isValid()) return true;
    if (indentLevel.has_value()) return true;
    if (align.has_value()) return true;
    if (defaultTabSize.isValid()) return true;
    if (rtl.has_value()) return true;
    if (eastAsianLineBreak.has_value()) return true;
    if (latinLineBreak.has_value()) return true;
    if (hangingPunctuation.has_value()) return true;
    if (fontAlign.has_value()) return true;
    if (lineSpacing.has_value()) return true;
    if (spacingBefore.has_value()) return true;
    if (spacingAfter.has_value()) return true;
    if (bulletColor.has_value()) return true;
    if (bulletSize.has_value()) return true;
    if (bulletFont && bulletFont->isValid()) return true;
    if (bulletType.has_value()) return true;
    if (bulletAutonumberStart.has_value()) return true;
    if (!bulletChar.isEmpty()) return true;
    if (defaultTextCharacterProperties.has_value()) return true;
    if (!tabStops.isEmpty()) return true;

    return true;
}

bool ParagraphProperties::operator==(const ParagraphProperties &other) const
{
    if (leftMargin != other.leftMargin) return false;
    if (rightMargin != other.rightMargin) return false;
    if (indent != other.indent) return false;
    if (indentLevel != other.indentLevel) return false;
    if (align != other.align) return false;
    if (defaultTabSize != other.defaultTabSize) return false;
    if (rtl != other.rtl) return false;
    if (eastAsianLineBreak != other.eastAsianLineBreak) return false;
    if (latinLineBreak != other.latinLineBreak) return false;
    if (hangingPunctuation != other.hangingPunctuation) return false;
    if (fontAlign != other.fontAlign) return false;
    if (lineSpacing != other.lineSpacing) return false;
    if (spacingBefore != other.spacingBefore) return false;
    if (spacingAfter != other.spacingAfter) return false;
    if (bulletColor != other.bulletColor) return false;
    if (bulletSize != other.bulletSize) return false;
    if (bulletFont != other.bulletFont) return false;
    if (bulletType != other.bulletType) return false;
    if (bulletAutonumberType != other.bulletAutonumberType) return false;
    if (bulletAutonumberStart != other.bulletAutonumberStart) return false;
    if (bulletChar != other.bulletChar) return false;
    if (defaultTextCharacterProperties != other.defaultTextCharacterProperties) return false;
    if (tabStops != other.tabStops) return false;
    return true;
}

bool ParagraphProperties::operator!=(const ParagraphProperties &other) const
{
    return !operator==(other);
}

void Paragraph::read(QXmlStreamReader &reader)
{
    const auto &name = reader.name();
    while (!reader.atEnd()) {
        auto token = reader.readNext();
        if (token == QXmlStreamReader::StartElement) {
            if (reader.name() == QLatin1String("pPr")) {
                ParagraphProperties p;
                p.read(reader);
                paragraphProperties = p;
            }
            else if (reader.name() == QLatin1String("endParaRPr")) {
                CharacterProperties p;
                p.read(reader);
                endParagraphDefaultCharacterProperties = p;
            }
            else if (reader.name() == QLatin1String("r")
                     || reader.name() == QLatin1String("br")
                     || reader.name() == QLatin1String("fld")) {
                TextRun t;
                t.read(reader);
                textRuns.append(t);
            }
            else reader.skipCurrentElement();
        }
        else if (token == QXmlStreamReader::EndElement && reader.name() == name)
            break;
    }
}

void Paragraph::write(QXmlStreamWriter &writer, const QString &name) const
{
    if (!isValid()) {
        writer.writeEmptyElement(name);
        return;
    }

    writer.writeStartElement(name);
    if (paragraphProperties.has_value()) paragraphProperties->write(writer, QLatin1String("a:pPr"));
    for (const auto &p: textRuns) p.write(writer);
    if (endParagraphDefaultCharacterProperties.has_value())
        endParagraphDefaultCharacterProperties->write(writer, QLatin1String("a:endParaRPr"));
    writer.writeEndElement(); //name
}

bool Paragraph::isValid() const
{
    return (paragraphProperties.has_value() || !textRuns.isEmpty()
            || endParagraphDefaultCharacterProperties.has_value());
}


bool Paragraph::operator ==(const Paragraph &other) const
{
    return (paragraphProperties == other.paragraphProperties &&
            textRuns == other.textRuns &&
            endParagraphDefaultCharacterProperties == other.endParagraphDefaultCharacterProperties);
}

bool Paragraph::operator !=(const Paragraph &other) const
{
    return (paragraphProperties != other.paragraphProperties ||
            textRuns != other.textRuns ||
            endParagraphDefaultCharacterProperties != other.endParagraphDefaultCharacterProperties);
}

void TextRun::read(QXmlStreamReader &reader)
{
    const auto &name = reader.name();
    if (type == Type::None) {
        if (name == QLatin1String("r")) type = Type::Regular;
        else if (name == QLatin1String("br")) type = Type::LineBreak;
        else if (name == QLatin1String("fld")) type = Type::TextField;
        else return;
    }

    const auto &a = reader.attributes();
    if (a.hasAttribute(QLatin1String("id"))) guid = a.value(QLatin1String("id")).toString();
    if (a.hasAttribute(QLatin1String("type"))) fieldType = a.value(QLatin1String("type")).toString();

    while (!reader.atEnd()) {
        auto token = reader.readNext();
        if (token == QXmlStreamReader::StartElement) {
            if (reader.name() == QLatin1String("pPr")) {
                ParagraphProperties p;
                p.read(reader);
                paragraphProperties = p;
            }
            else if (reader.name() == QLatin1String("rPr")) {
                CharacterProperties p;
                p.read(reader);
                characterProperties = p;
            }
            else if (reader.name() == QLatin1String("t")) {
                text = reader.readElementText();
            }
            else reader.skipCurrentElement();
        }
        else if (token == QXmlStreamReader::EndElement && reader.name() == name)
            break;
    }
}

void TextRun::write(QXmlStreamWriter &writer) const
{
    switch (type) {
        case Type::None: return;
        case Type::Regular: {
            //if (text.isEmpty()) return; //required field
            writer.writeStartElement(QLatin1String("a:r"));
            if (characterProperties.has_value()) characterProperties->write(writer, QLatin1String("a:rPr"));
            writer.writeTextElement(QLatin1String("a:t"), text);
            writer.writeEndElement();
            break;
        }
        case Type::LineBreak: {
            writer.writeStartElement(QLatin1String("a:br"));
            if (characterProperties.has_value()) characterProperties->write(writer, QLatin1String("a:rPr"));
            writer.writeEndElement();
            break;
        }
        case Type::TextField: {
            if (guid.isEmpty()) return; //required field
            writer.writeStartElement(QLatin1String("a:fld"));
            writer.writeAttribute(QLatin1String("id"), guid);
            if (!fieldType.isEmpty()) writer.writeAttribute(QLatin1String("type"), fieldType);
            if (characterProperties.has_value()) characterProperties->write(writer, QLatin1String("a:rPr"));
            if (paragraphProperties.has_value()) paragraphProperties->write(writer, QLatin1String("a:pPr"));
            if (!text.isEmpty()) writer.writeTextElement(QLatin1String("a:t"), text);
            writer.writeEndElement();
            break;
        }
    }
}

TextRun::TextRun()
{

}

TextRun::TextRun(const TextRun &other) :
    type{other.type}, text{other.text},
    characterProperties{other.characterProperties},
    paragraphProperties{other.paragraphProperties},
    guid{other.guid}, fieldType{other.fieldType}
{

}

TextRun::TextRun(TextRun::Type type, const QString &text, const CharacterProperties &properties)
    : type{type}, text{text}, characterProperties{properties}
{

}

TextRun::TextRun(TextRun::Type type)
    : type{type}
{

}

bool TextRun::operator ==(const TextRun &other) const
{
    if (type != other.type) return false;
    if (text != other.text) return false;
    if (characterProperties != other.characterProperties) return false;
    if (paragraphProperties != other.paragraphProperties) return false;
    if (guid != other.guid) return false;
    if (fieldType != other.fieldType) return false;
    return true;
}

bool TextRun::operator !=(const TextRun &other) const
{
    return !operator==(other);
}

bool CharacterProperties::operator ==(const CharacterProperties &other) const
{
    if (kumimoji != other.kumimoji) return false;
    if (language != other.language) return false;
    if (alternateLanguage != other.alternateLanguage) return false;
    if (fontSize != other.fontSize) return false;
    if (bold != other.bold) return false;
    if (italic != other.italic) return false;
    if (underline != other.underline) return false;
    if (strike != other.strike) return false;
    if (kerningFontSize != other.kerningFontSize) return false;
    if (capitalization != other.capitalization) return false;
    if (spacing != other.spacing) return false;
    if (normalizeHeights != other.normalizeHeights) return false;
    if (noProofing != other.noProofing) return false;
    if (proofingNeeded != other.proofingNeeded) return false;
    if (checkForSmartTagsNeeded != other.checkForSmartTagsNeeded) return false;
    if (spellingErrorFound != other.spellingErrorFound) return false;
    if (baseline != other.baseline) return false;
    if (smartTagId != other.smartTagId) return false;
    return true;
}

bool CharacterProperties::operator !=(const CharacterProperties &other) const
{
    return !operator==(other);
}

void CharacterProperties::read(QXmlStreamReader &reader)
{
//    <xsd:complexType name="CT_TextCharacterProperties">
//        <xsd:sequence>
//          <xsd:element name="ln" type="CT_LineProperties" minOccurs="0" maxOccurs="1"/>
//          <xsd:group ref="EG_FillProperties" minOccurs="0" maxOccurs="1"/>
//          <xsd:group ref="EG_EffectProperties" minOccurs="0" maxOccurs="1"/>
//          <xsd:element name="highlight" type="CT_Color" minOccurs="0" maxOccurs="1"/>
//          <xsd:group ref="EG_TextUnderlineLine" minOccurs="0" maxOccurs="1"/>
//          <xsd:group ref="EG_TextUnderlineFill" minOccurs="0" maxOccurs="1"/>
//          <xsd:element name="latin" type="CT_TextFont" minOccurs="0" maxOccurs="1"/>
//          <xsd:element name="ea" type="CT_TextFont" minOccurs="0" maxOccurs="1"/>
//          <xsd:element name="cs" type="CT_TextFont" minOccurs="0" maxOccurs="1"/>
//          <xsd:element name="sym" type="CT_TextFont" minOccurs="0" maxOccurs="1"/>
//          <xsd:element name="hlinkClick" type="CT_Hyperlink" minOccurs="0" maxOccurs="1"/>
//          <xsd:element name="hlinkMouseOver" type="CT_Hyperlink" minOccurs="0" maxOccurs="1"/>
//          <xsd:element name="rtl" type="CT_Boolean" minOccurs="0"/>
//          <xsd:element name="extLst" type="CT_OfficeArtExtensionList" minOccurs="0" maxOccurs="1"/>
//        </xsd:sequence>
//        <xsd:attribute name="kumimoji" type="xsd:boolean" use="optional"/>
//        <xsd:attribute name="lang" type="s:ST_Lang" use="optional"/>
//        <xsd:attribute name="altLang" type="s:ST_Lang" use="optional"/>
//        <xsd:attribute name="sz" type="ST_TextFontSize" use="optional"/>
//        <xsd:attribute name="b" type="xsd:boolean" use="optional"/>
//        <xsd:attribute name="i" type="xsd:boolean" use="optional"/>
//        <xsd:attribute name="u" type="ST_TextUnderlineType" use="optional"/>
//        <xsd:attribute name="strike" type="ST_TextStrikeType" use="optional"/>
//        <xsd:attribute name="kern" type="ST_TextNonNegativePoint" use="optional"/>
//        <xsd:attribute name="cap" type="ST_TextCapsType" use="optional" default="none"/>
//        <xsd:attribute name="spc" type="ST_TextPoint" use="optional"/>
//        <xsd:attribute name="normalizeH" type="xsd:boolean" use="optional"/>
//        <xsd:attribute name="baseline" type="ST_Percentage" use="optional"/>
//        <xsd:attribute name="noProof" type="xsd:boolean" use="optional"/>
//        <xsd:attribute name="dirty" type="xsd:boolean" use="optional" default="true"/>
//        <xsd:attribute name="err" type="xsd:boolean" use="optional" default="false"/>
//        <xsd:attribute name="smtClean" type="xsd:boolean" use="optional" default="true"/>
//        <xsd:attribute name="smtId" type="xsd:unsignedInt" use="optional" default="0"/>
//        <xsd:attribute name="bmk" type="xsd:string" use="optional"/>
//      </xsd:complexType>
    const auto &name = reader.name();

    const auto &a = reader.attributes();
    parseAttributeBool(a, QLatin1String("kumimoji"), kumimoji);
    parseAttributeString(a, QLatin1String("lang"), language);
    parseAttributeString(a, QLatin1String("altLang"), alternateLanguage);
    if (a.hasAttribute(QLatin1String("sz"))) fontSize = a.value(QLatin1String("sz")).toDouble() / 100;
    parseAttributeBool(a, QLatin1String("b"), bold);
    parseAttributeBool(a, QLatin1String("i"), italic);
    if (a.hasAttribute(QLatin1String("u"))) {
        UnderlineType t;
        fromString(a.value(QLatin1String("u")).toString(), t);
        underline = t;
    }
    if (a.hasAttribute(QLatin1String("strike"))) {
        StrikeType t;
        fromString(a.value(QLatin1String("strike")).toString(), t);
        strike = t;
    }
    if (a.hasAttribute(QLatin1String("kern"))) kerningFontSize = a.value(QLatin1String("kern")).toDouble() / 100;
    if (a.hasAttribute(QLatin1String("cap"))) {
        CapitalizationType t;
        fromString(a.value(QLatin1String("cap")).toString(), t);
        capitalization = t;
    }
    if (a.hasAttribute(QLatin1String("spc"))) spacing = TextPoint::create(a.value(QLatin1String("spc")).toString());
    if (a.hasAttribute(QLatin1String("bmk"))) bookmarkLinkTarget = a.value(QLatin1String("bmk")).toString();
    parseAttributeBool(a, QLatin1String("normalizeH"), normalizeHeights);
    parseAttributeBool(a, QLatin1String("noProof"), noProofing);
    parseAttributeBool(a, QLatin1String("dirty"), proofingNeeded);
    parseAttributeBool(a, QLatin1String("err"), spellingErrorFound);
    parseAttributeBool(a, QLatin1String("smtClean"), checkForSmartTagsNeeded);
    parseAttributePercent(a, QLatin1String("baseline"), baseline);
    parseAttributeInt(a, QLatin1String("smtId"), smartTagId);

    while (!reader.atEnd()) {
        auto token = reader.readNext();
        if (token == QXmlStreamReader::StartElement) {
            if (reader.name() == QLatin1String("ln")) {
                line.read(reader);
            }
            else if (reader.name().endsWith(QLatin1String("Fill"))) {
                fill.read(reader);
            }
            else if (reader.name() == QLatin1String("effectLst")) {
                effect.read(reader);
            }
            else if (reader.name() == QLatin1String("effectDag")) {

            }
            else if (reader.name() == QLatin1String("highlight")) {
                reader.readNextStartElement();
                highlightColor.read(reader);
            }
            else if (reader.name() == QLatin1String("uLnTx")) {
                textUnderlineFollowsText = true;
            }
            else if (reader.name() == QLatin1String("uLn")) {
                textUnderline.read(reader);
            }
            else if (reader.name() == QLatin1String("uFillTx")) {
                textUnderlineFillFollowsText = true;
            }
            else if (reader.name() == QLatin1String("uFill")) {
                reader.readNextStartElement();
                textUnderlineFill.read(reader);
            }
            else if (reader.name() == QLatin1String("latin")) {
                latinFont.read(reader);
            }
            else if (reader.name() == QLatin1String("ea")) {
                eastAsianFont.read(reader);
            }
            else if (reader.name() == QLatin1String("cs")) {
                complexScriptFont.read(reader);
            }
            else if (reader.name() == QLatin1String("sym")) {
                symbolFont.read(reader);
            }
            else if (reader.name() == QLatin1String("rtl")) {
                rightToLeft = fromST_Boolean(reader.attributes().value(QLatin1String("val")));
            }
            else reader.skipCurrentElement();
            //TODO: hyperlinks
        //          <xsd:element name="hlinkClick" type="CT_Hyperlink" minOccurs="0" maxOccurs="1"/>
        //          <xsd:element name="hlinkMouseOver" type="CT_Hyperlink" minOccurs="0" maxOccurs="1"/>
        }
        else if (token == QXmlStreamReader::EndElement && reader.name() == name)
            break;
    }
}

void CharacterProperties::write(QXmlStreamWriter &writer, const QString &name) const
{
    writer.writeStartElement(name);
    if (kumimoji.has_value()) writer.writeAttribute(QLatin1String("kumimoji"),
                                                    toST_Boolean(kumimoji.value()));
    if (!language.isEmpty()) writer.writeAttribute(QLatin1String("lang"), language);
    if (!alternateLanguage.isEmpty()) writer.writeAttribute(QLatin1String("altLang"),
                                                            alternateLanguage);
    if (fontSize.has_value()) writer.writeAttribute(QLatin1String("sz"),
                                                    QString::number(qRound(fontSize.value()*100)));
    writeAttribute(writer, QLatin1String("b"), bold);
    writeAttribute(writer, QLatin1String("i"), italic);
    if (underline.has_value()) {
        QString s; toString(underline.value(), s);
        writer.writeAttribute(QLatin1String("u"), s);
    }
    if (strike.has_value()) {
        QString s; toString(strike.value(), s);
        writer.writeAttribute(QLatin1String("strike"), s);
    }
    if (kerningFontSize.has_value()) writer.writeAttribute(QLatin1String("kern"),
                                                           QString::number(qRound(kerningFontSize.value()*100)));
    if (capitalization.has_value()) {
        QString s; toString(capitalization.value(), s);
        writer.writeAttribute(QLatin1String("cap"), s);
    }
    if (spacing.has_value()) writer.writeAttribute(QLatin1String("spc"), spacing.value().toString());
    writeAttribute(writer, QLatin1String("normalizeH"), normalizeHeights);
    writeAttribute(writer, QLatin1String("noProof"), noProofing);
    writeAttribute(writer, QLatin1String("dirty"), proofingNeeded);
    writeAttribute(writer, QLatin1String("err"), spellingErrorFound);
    writeAttribute(writer, QLatin1String("smtClean"), checkForSmartTagsNeeded);
    if (baseline.has_value()) writer.writeAttribute(QLatin1String("baseline"),
                                                    toST_Percent(baseline.value()));
    if (smartTagId.has_value()) writer.writeAttribute(QLatin1String("smtId"),
                                                    QString::number(smartTagId.value()));
    if (!bookmarkLinkTarget.isEmpty()) writer.writeAttribute(QLatin1String("bmk"),
                                                             bookmarkLinkTarget);
    if (line.isValid()) line.write(writer, QLatin1String("a:ln"));
    if (fill.isValid()) fill.write(writer);
    if (effect.isValid()) effect.write(writer);
    if (highlightColor.isValid()) highlightColor.write(writer, QLatin1String("a:highlight"));
    if (textUnderlineFollowsText.has_value() && textUnderlineFollowsText.value())
        writer.writeEmptyElement(QLatin1String("a:uLnTx"));
    else if (textUnderline.isValid()) textUnderline.write(writer, QLatin1String("a:uLn"));
    if (textUnderlineFillFollowsText.has_value() && textUnderlineFillFollowsText.value())
        writer.writeEmptyElement(QLatin1String("a:uFillTx"));
    else if (textUnderlineFill.isValid()) {
        writer.writeStartElement(QLatin1String("a:uFill"));
        textUnderlineFill.write(writer);
        writer.writeEndElement();
    }
    if (latinFont.isValid()) latinFont.write(writer, QLatin1String("a:latin"));
    if (eastAsianFont.isValid()) eastAsianFont.write(writer, QLatin1String("a:ea"));
    if (complexScriptFont.isValid()) complexScriptFont.write(writer, QLatin1String("a:cs"));
    if (symbolFont.isValid()) symbolFont.write(writer, QLatin1String("a:sym"));
    if (rightToLeft.has_value()) writeEmptyElement(writer, QLatin1String("a:rtl"), rightToLeft);

    //TODO: hyperlinks
//          <xsd:element name="hlinkClick" type="CT_Hyperlink" minOccurs="0" maxOccurs="1"/>
//          <xsd:element name="hlinkMouseOver" type="CT_Hyperlink" minOccurs="0" maxOccurs="1"/>
    writer.writeEndElement();
}

CharacterProperties CharacterProperties::from(const QTextCharFormat &format)
{
    CharacterProperties pr;
    if (format.isValid()) {
//        std::optional<bool> kumimoji;
//        QString language; // f.e. "en-US"
//        QString alternateLanguage; // f.e. "en-US"
        if (format.hasProperty(QTextFormat::FontPointSize))
            pr.fontSize = format.fontPointSize();
        if (format.hasProperty(QTextFormat::FontWeight))
            pr.bold = format.fontWeight() > QFont::Normal;
        if (format.hasProperty(QTextFormat::FontItalic))
            pr.italic = format.fontItalic();
        if (format.hasProperty(QTextFormat::TextUnderlineStyle))
            switch (format.underlineStyle()) {
                case QTextCharFormat::NoUnderline: pr.underline = CharacterProperties::UnderlineType::None;
                    break;
                case QTextCharFormat::SingleUnderline: pr.underline = CharacterProperties::UnderlineType::Single;
                    break;
                case QTextCharFormat::DashUnderline: pr.underline = CharacterProperties::UnderlineType::Dash;
                    break;
                case QTextCharFormat::DotLine: pr.underline = CharacterProperties::UnderlineType::Dotted;
                    break;
                case QTextCharFormat::DashDotLine: pr.underline = CharacterProperties::UnderlineType::DotDash;
                    break;
                case QTextCharFormat::DashDotDotLine: pr.underline = CharacterProperties::UnderlineType::DotDotDash;
                    break;
                case QTextCharFormat::WaveUnderline: pr.underline = CharacterProperties::UnderlineType::Wavy;
                    break;
                case QTextCharFormat::SpellCheckUnderline: pr.underline = CharacterProperties::UnderlineType::Wavy;
                    break;
            }
        if (format.hasProperty(QTextFormat::FontStrikeOut))
            pr.strike = format.fontStrikeOut() ? CharacterProperties::StrikeType::Single : CharacterProperties::StrikeType::None;
//        std::optional<double> kerningFontSize; //in pt
        if (format.hasProperty(QTextFormat::FontCapitalization))
            switch (format.fontCapitalization()) {
                case QFont::MixedCase: pr.capitalization = CapitalizationType::None;
                    break;
                case QFont::AllUppercase: pr.capitalization = CapitalizationType::AllCaps;
                    break;
                case QFont::AllLowercase:
                    break;
                case QFont::SmallCaps: pr.capitalization = CapitalizationType::SmallCaps;
                    break;
                case QFont::Capitalize:
                    break;
            }
        if (format.hasProperty(QTextFormat::FontLetterSpacing))
            pr.spacing = TextPoint(format.fontLetterSpacing());

//        std::optional<bool> normalizeHeights;
//        std::optional<bool> noProofing;
//        std::optional<bool> proofingNeeded;
//        std::optional<bool> checkForSmartTagsNeeded;
//        std::optional<bool> spellingErrorFound;

        if (format.hasProperty(QTextFormat::TextVerticalAlignment)) {
            switch (format.verticalAlignment()) {
                case QTextCharFormat::AlignNormal:
                    break;
                case QTextCharFormat::AlignSuperScript: pr.baseline = 30;
                    break;
                case QTextCharFormat::AlignSubScript: pr.baseline = -25;
                    break;
                case QTextCharFormat::AlignMiddle:
                    break;
                case QTextCharFormat::AlignTop:
                    break;
                case QTextCharFormat::AlignBottom:
                    break;
                case QTextCharFormat::AlignBaseline:
                    break;
            }
        }
//        std::optional<int> smartTagId;
        if (format.hasProperty(QTextFormat::OutlinePen))
            pr.line.setColor(format.property(QTextFormat::OutlinePen).value<QPen>().color());
        if (format.hasProperty(QTextFormat::ForegroundBrush))
            pr.fill.setColor(format.foreground().color());
        if (format.hasProperty(QTextFormat::FontFamily))
            pr.latinFont = Font(format.font());
        //        Font eastAsianFont;
        //        Font complexScriptFont;
        //        Font symbolFont;

//        QString bookmarkLinkTarget;
//        LineFormat line;
//        Effect effect;
//        Color highlightColor;

//        std::optional<bool> textUnderlineFollowsText;
//        LineFormat textUnderline;
        if (format.hasProperty(QTextFormat::TextUnderlineColor))
            pr.textUnderlineFill.setColor(format.property(QTextFormat::OutlinePen).value<QColor>());
//        std::optional<bool> textUnderlineFillFollowsText;
    }
    return pr;
}

TextProperties Text::textProperties() const
{
    if (d) return d->textProperties;
    return {};
}

TextProperties &Text::textProperties()
{
    if (!d) d = new TextPrivate;
    return d->textProperties;
}

void Text::setTextProperties(const TextProperties &textProperties)
{
    if (!d) d = new TextPrivate;
    d->textProperties = textProperties;
}

ParagraphProperties Text::defaultParagraphProperties() const
{
    //p -> pPr
    if (!d) return {};
    if (d->paragraphs.isEmpty()) return {};
    return d->paragraphs.first().paragraphProperties.value_or(ParagraphProperties());
}

ParagraphProperties &Text::defaultParagraphProperties()
{
    //p->pPr
    if (!d) d = new TextPrivate;
    if (d->paragraphs.isEmpty()) d->paragraphs << Paragraph();
    auto &p = d->paragraphs.first();
    if (!p.paragraphProperties.has_value()) p.paragraphProperties = ParagraphProperties();
    return p.paragraphProperties.value();
}

void Text::setDefaultParagraphProperties(const ParagraphProperties &defaultParagraphProperties)
{
    //p->pPr
    if (!d) d = new TextPrivate;
    if (d->paragraphs.isEmpty()) d->paragraphs << Paragraph();
    auto &p = d->paragraphs.first();
    p.paragraphProperties = defaultParagraphProperties;
}

CharacterProperties Text::defaultCharacterProperties() const
{
    //p->pPr->defRPr
    if (!d) return {};
    if (d->paragraphs.isEmpty()) return {};
    const auto p = d->paragraphs.first();
    return p.paragraphProperties.value_or(ParagraphProperties())
            .defaultTextCharacterProperties.value_or(CharacterProperties());
}

CharacterProperties &Text::defaultCharacterProperties()
{
    if (!d) d = new TextPrivate;
    if (d->paragraphs.isEmpty()) d->paragraphs << Paragraph();
    auto &p = d->paragraphs.first();
    if (!p.paragraphProperties.has_value()) p.paragraphProperties = ParagraphProperties();
    auto &pp = p.paragraphProperties.value();
    if (!pp.defaultTextCharacterProperties.has_value()) pp.defaultTextCharacterProperties = CharacterProperties();
    return pp.defaultTextCharacterProperties.value();
}

void Text::setDefaultCharacterProperties(const CharacterProperties &defaultCharacterProperties)
{
    if (!d) d = new TextPrivate;
    if (d->paragraphs.isEmpty()) d->paragraphs << Paragraph();
    auto &p = d->paragraphs.first();
    if (!p.paragraphProperties.has_value()) p.paragraphProperties = ParagraphProperties();
    auto &pp = p.paragraphProperties.value();
    pp.defaultTextCharacterProperties = defaultCharacterProperties;
}

QDebug operator<<(QDebug dbg, const Text &t)
{
    QDebugStateSaver saver(dbg);
    dbg.setAutoInsertSpaces(false);
    dbg << "QXlsx::Text(";
    if (t.isValid()) {
        dbg << "type: " << static_cast<int>(t.d->type);
        dbg << ", paragraphs: " << t.d->paragraphs;
        dbg << ", reference: " <<t.d->reference;
        dbg << ", plainText: " <<t.d->plainText;
        dbg << ", cashe: " <<t.d->cashe;
        dbg << ", textProperties: ";
        if (t.d->textProperties.isValid()) dbg << t.d->textProperties;
        else dbg << "not set";
        dbg << ", ";
        dbg << "defaultParagraphProperties: " << t.d->defaultParagraphProperties;
    }
    dbg << ")";
    return dbg;
}

QDebug operator<<(QDebug dbg, const TextProperties &t)
{
    QDebugStateSaver saver(dbg);
    dbg.setAutoInsertSpaces(false);
    dbg << "QXlsx::TextProperties(";
    if (t.isValid()) {
        if (t.spcFirstLastPara.has_value()) dbg << "spcFirstLastPara: " << t.spcFirstLastPara.value() << ", ";
        if (t.verticalOverflow.has_value()) dbg << "verticalOverflow: " << static_cast<int>(t.verticalOverflow.value()) << ", ";
        if (t.horizontalOverflow.has_value()) dbg << "horizontalOverflow: " << static_cast<int>(t.horizontalOverflow.value()) << ", ";
        if (t.verticalOrientation.has_value()) dbg << "verticalOrientation: " << static_cast<int>(t.verticalOrientation.value()) << ", ";
        if (t.rotation.has_value()) dbg << "rotation: " << t.rotation.value().toString() << ", ";
        if (t.wrap.has_value()) dbg << "wrap: " << t.wrap.value() << ", ";
        if (t.leftInset.isValid()) dbg << "leftInset: " << t.leftInset.toString() << ", ";
        if (t.rightInset.isValid()) dbg << "rightInset: " << t.rightInset.toString() << ", ";
        if (t.topInset.isValid()) dbg << "topInset: " << t.topInset.toString() << ", ";
        if (t.bottomInset.isValid()) dbg << "bottomInset: " << t.bottomInset.toString() << ", ";
        if (t.columnCount.has_value()) dbg << "columnCount: " << t.columnCount.value() << ", ";
        if (t.columnSpace.isValid()) dbg << "columnSpace: " << t.columnSpace.toString() << ", ";
        if (t.columnsRtl.has_value()) dbg << "columnsRtl: " << t.columnsRtl.value() << ", ";
        if (t.fromWordArt.has_value()) dbg << "fromWordArt: " << t.fromWordArt.value() << ", ";
        if (t.anchor.has_value()) dbg << "anchor: " << static_cast<int>(t.anchor.value()) << ", ";
        if (t.anchorCentering.has_value()) dbg << "anchorCentering: " << t.anchorCentering.value() << ", ";
        if (t.forceAntiAlias.has_value()) dbg << "forceAntiAlias: " << t.forceAntiAlias.value() << ", ";
        if (t.upright.has_value()) dbg << "upright: " << t.upright.value() << ", ";
        if (t.compatibleLineSpacing.has_value()) dbg << "compatibleLineSpacing: " << t.compatibleLineSpacing.value() << ", ";
        if (t.textShape.has_value()) dbg << "textShape: " << t.textShape.value() << ", ";
        if (t.textAutofit.has_value()) dbg << "textAutofit: " << static_cast<int>(t.textAutofit.value()) << ", ";
        if (t.fontScale.has_value()) dbg << "fontScale: " << t.fontScale.value() << ", ";
        if (t.lineSpaceReduction.has_value()) dbg << "lineSpaceReduction: " << t.lineSpaceReduction.value() << ", ";
        if (t.scene3D.has_value()) dbg << "scene3D: " << t.scene3D.value();
        if (t.text3D.has_value()) dbg << "text3D: " << t.text3D.value();
        if (t.z.isValid()) dbg << "z: " << t.z.toString();
    }
    dbg << ")";
    return dbg;
}

QDebug operator<<(QDebug dbg, const ListStyleProperties &t)
{
    QDebugStateSaver saver(dbg);
    dbg.setAutoInsertSpaces(false);
    dbg << "QXlsx::ListStyleProperties(";
//TODO:
    dbg << ")";
    return dbg;
}

QDebug operator<<(QDebug dbg, const Paragraph &t)
{
    QDebugStateSaver saver(dbg);
    dbg.setAutoInsertSpaces(false);
    dbg << "QXlsx::Paragraph(";

    if (t.paragraphProperties.has_value()) dbg << "paragraphProperties: " << t.paragraphProperties.value() << ", ";
    dbg << "textRuns: " << t.textRuns << ", ";
    if (t.endParagraphDefaultCharacterProperties.has_value())
        dbg << "endParagraphDefaultCharacterProperties: " << t.endParagraphDefaultCharacterProperties.value() << ", ";

    dbg << ")";
    return dbg;
}

QDebug operator<<(QDebug dbg, const ParagraphProperties &t)
{
    QDebugStateSaver saver(dbg);
    dbg.setAutoInsertSpaces(false);
    dbg << "QXlsx::ParagraphProperties(";

    if (t.leftMargin.isValid()) dbg << "leftMargin: " << t.leftMargin.toString() << ", ";
    if (t.rightMargin.isValid()) dbg << "rightMargin: " << t.rightMargin.toString() << ", ";
    if (t.indent.isValid()) dbg << "indent: " << t.indent.toString() << ", ";
    if (t.indentLevel.has_value()) dbg << "indentLevel: " << t.indentLevel.value() << ", "; //[0..8]
    if (t.align.has_value()) dbg << "align: " << static_cast<int>(t.align.value()) << ", ";
    if (t.defaultTabSize.isValid()) dbg << "defaultTabSize: " << t.defaultTabSize.toString() << ", ";
    if (t.rtl.has_value()) dbg << "rtl: " << t.rtl.value() << ", ";
    if (t.eastAsianLineBreak.has_value()) dbg << "eastAsianLineBreak: " << t.eastAsianLineBreak.value() << ", ";
    if (t.latinLineBreak.has_value()) dbg << "latinLineBreak: " << t.latinLineBreak.value() << ", ";
    if (t.hangingPunctuation.has_value()) dbg << "hangingPunctuation: " << t.hangingPunctuation.value() << ", ";
    if (t.fontAlign.has_value()) dbg << "fontAlign: " << static_cast<int>(t.fontAlign.value()) << ", ";
    if (t.lineSpacing.has_value()) dbg << "lineSpacing: " << t.lineSpacing->toString() << ", ";
    if (t.spacingBefore.has_value()) dbg << "spacingBefore: " << t.spacingBefore->toString() << ", ";
    if (t.spacingAfter.has_value()) dbg << "spacingAfter: " << t.spacingAfter->toString() << ", ";
    if (t.bulletColor.has_value()) dbg << "bulletColor: " << t.bulletColor.value() << ", ";
    if (t.bulletSize.has_value()) dbg << "bulletSize: " << t.bulletSize->toString() << ", ";
    if (t.bulletFont.has_value()) dbg << "bulletFont: " << t.bulletFont.value() << ", ";
    if (t.bulletType.has_value()) dbg << "bulletType: " << static_cast<int>(t.bulletType.value()) << ", ";
    dbg << "bulletAutonumberType: " << static_cast<int>(t.bulletAutonumberType) << ", ";
    if (t.bulletAutonumberStart.has_value()) dbg << "bulletAutonumberStart: " << t.bulletAutonumberStart.value() << ", ";
    if (!t.bulletChar.isEmpty()) dbg << "bulletChar: " << t.bulletChar << ", ";
    if (t.defaultTextCharacterProperties.has_value())
        dbg << "defaultTextCharacterProperties: " << t.defaultTextCharacterProperties.value() << ", ";
    dbg << "tabStops: " << t.tabStops;

    dbg << ")";
    return dbg;
}

QDebug operator<<(QDebug dbg, const CharacterProperties &t)
{
    QDebugStateSaver saver(dbg);
    dbg.setAutoInsertSpaces(false);
    dbg << "QXlsx::CharacterProperties(";

    if (t.kumimoji.has_value()) dbg << "kumimoji: " << t.kumimoji.value() << ", ";
    if (!t.language.isEmpty()) dbg << "language: " << t.language << ", "; // f.e. "en-US"
    if (!t.alternateLanguage.isEmpty()) dbg << "alternateLanguage: " << t.alternateLanguage << ", "; // f.e. "en-US"
    if (t.fontSize.has_value()) dbg << "fontSize: " << static_cast<int>(t.fontSize.value()) << ", "; //in pt
    if (t.bold.has_value()) dbg << "bold: " << t.bold.value() << ", ";
    if (t.italic.has_value()) dbg << "italic: " << t.italic.value() << ", ";
    if (t.underline.has_value()) dbg << "underline: " << static_cast<int>(t.underline.value()) << ", ";
    if (t.strike.has_value()) dbg << "strike: " << static_cast<int>(t.strike.value()) << ", ";
    if (t.kerningFontSize.has_value()) dbg << "kerningFontSize: " << static_cast<int>(t.kerningFontSize.value()) << ", "; //in pt
    if (t.capitalization.has_value()) dbg << "capitalization: " << static_cast<int>(t.capitalization.value()) << ", ";
    if (t.spacing.has_value()) dbg << "spacing: " << t.spacing.value().toString() << ", ";
    if (t.normalizeHeights.has_value()) dbg << "normalizeHeights: " << t.normalizeHeights.value() << ", ";
    if (t.baseline.has_value()) dbg << "baseline: " << static_cast<int>(t.baseline.value()) << ", ";
    if (t.noProofing.has_value()) dbg << "noProofing: " << t.noProofing.value() << ", ";
    if (t.proofingNeeded.has_value()) dbg << "proofingNeeded: " << t.proofingNeeded.value() << ", ";
    if (t.checkForSmartTagsNeeded.has_value()) dbg << "checkForSmartTagsNeeded: " << t.checkForSmartTagsNeeded.value() << ", ";
    if (t.spellingErrorFound.has_value()) dbg << "spellingErrorFound: " << t.spellingErrorFound.value() << ", ";
    if (t.smartTagId.has_value()) dbg << "smartTagId: " << t.smartTagId.value() << ", ";
    if (!t.bookmarkLinkTarget.isEmpty()) dbg << "bookmarkLinkTarget: " << t.bookmarkLinkTarget << ", ";
    if (t.line.isValid()) dbg << "line: " << t.line << ", ";
    if (t.fill.isValid()) dbg << "fill: " << t.fill << ", ";
    if (t.effect.isValid()) dbg << "effect: " << t.effect << ", ";
    if (t.highlightColor.isValid()) dbg << "highlightColor: " << t.highlightColor << ", ";
    if (t.textUnderlineFollowsText.has_value()) dbg << "textUnderlineFollowsText: " << t.textUnderlineFollowsText.value() << ", ";
    if (t.textUnderline.isValid()) dbg << "textUnderline: " << t.textUnderline << ", ";
    if (t.textUnderlineFillFollowsText.has_value()) dbg << "textUnderlineFillFollowsText: " << t.textUnderlineFillFollowsText.value() << ", ";
    if (t.textUnderlineFill.isValid()) dbg << "textUnderlineFill: " << t.textUnderlineFill << ", ";
    if (t.latinFont.isValid()) dbg << "latinFont: " << t.latinFont << ", ";
    if (t.eastAsianFont.isValid()) dbg << "eastAsianFont: " << t.eastAsianFont << ", ";
    if (t.complexScriptFont.isValid()) dbg << "complexScriptFont: " << t.complexScriptFont << ", ";
    if (t.symbolFont.isValid()) dbg << "symbolFont: " << t.symbolFont << ", ";
    if (t.rightToLeft.has_value()) dbg << "rightToLeft: " << t.rightToLeft.value() << ", ";

    dbg << ")";
    return dbg;
}

QDebug operator<<(QDebug dbg, const TextRun &t)
{
    QDebugStateSaver saver(dbg);
    dbg.setAutoInsertSpaces(false);
    dbg << "QXlsx::TextRun(";
    dbg << "type: " << static_cast<int>(t.type) << ", ";
    dbg << "text: " << t.text << ", ";
    if (t.characterProperties.has_value())
        dbg << "characterProperties: " << t.characterProperties.value() << ", ";
    if (t.paragraphProperties.has_value()) dbg << "paragraphProperties: " << t.paragraphProperties.value() << ", ";
    dbg << "guid: " << t.guid << ", "; //required
    dbg << "fieldType: " << t.fieldType << ", "; //optional
    dbg << ")";
    return dbg;
}

QDebug operator<<(QDebug dbg, const ParagraphProperties::TabAlign &t)
{
    QDebugStateSaver saver(dbg);
    dbg.nospace() << static_cast<int>(t);
    return dbg;
}

//std::optional<bool> (\w+);
//if (t.\1.has_value()) dbg << "\1: " << t.\1.value() << ", ";

QT_END_NAMESPACE_XLSX











