#include <QtGlobal>
#include <QDataStream>
#include <QDebug>

#include "xlsxtitle.h"
#include "xlsxtext.h"
#include "xlsxshapeformat.h"
#include "xlsxabstractsheet.h"

QT_BEGIN_NAMESPACE_XLSX

//      <xsd:element name="tx" type="CT_Tx" minOccurs="0" maxOccurs="1"/>
//      <xsd:element name="layout" type="CT_Layout" minOccurs="0" maxOccurs="1"/>
//      <xsd:element name="overlay" type="CT_Boolean" minOccurs="0" maxOccurs="1"/>
//      <xsd:element name="spPr" type="a:CT_ShapeProperties" minOccurs="0" maxOccurs="1"/>
//      <xsd:element name="txPr" type="a:CT_TextBody" minOccurs="0" maxOccurs="1"/>
//      <xsd:element name="extLst" type="CT_ExtensionList" minOccurs="0" maxOccurs="1"/>

class TitlePrivate : public QSharedData
{
public:
    Text text; //optional, c:tx
    Layout layout;
    std::optional<bool> overlay; //optional, c:overlay
    ShapeFormat shape;

    Text textProperties;

    bool operator==(const TitlePrivate &other) const;

    TitlePrivate();
    TitlePrivate(const TitlePrivate &other);
    ~TitlePrivate();
};

Title::Title()
{

}

Title::Title(const QString &text)
{
    d = new TitlePrivate;
    d->text.setHtml(text, false);
}

Title::Title(const Title &other) : d(other.d)
{

}

Title &Title::operator=(const Title &other)
{
    d = other.d;
    return *this;
}

Title::~Title()
{

}

QString Title::toPlainText() const
{
    if (d) return d->text.toPlainString();
    return {};
}

void Title::setPlainText(const QString &text)
{
    if (!d) d = new TitlePrivate;
    d->text.setHtml(text, false);
}

QString Title::toHtml() const
{
    if (d) return d->text.toHtml();
    return {};
}

void Title::setHtml(const QString &formattedText)
{
    if (!d) d = new TitlePrivate;
    d->text.setHtml(formattedText);
}

QString Title::stringReference() const
{
    if (d) return d->text.stringReference();
    return {};
}

void Title::setStringReference(const QString &reference)
{
    if (!d) d = new TitlePrivate;
    d->text.setStringReference(reference);
}

void Title::setStringReference(const CellRange &range, AbstractSheet *sheet)
{
    if (!d) d = new TitlePrivate;
    if (!range.isValid())
        return;
    if (sheet && sheet->sheetType() != AbstractSheet::ST_WorkSheet)
        return;
    QString sheetName = escapeSheetName(sheet->sheetName());
    QString reference = sheetName + QLatin1String("!") + range.toString(true, true);
    setStringReference(reference);
}

bool Title::isRichString() const
{
    if (!d) return false;
    return (d->text.isRichString() || d->textProperties.isValid());
}

bool Title::isPlainString() const
{
    if (!d) return false;
    return ((d->text.isPlainString() || d->text.isStringReference()) && !d->textProperties.isValid());
}

bool Title::isStringReference() const
{
    return (d ? d->text.isStringReference() : false);
}

Text &Title::text()
{
    if (!d) d = new TitlePrivate;
    return d->text;
}

Text Title::text() const
{
    if (d) return d->text;
    return {};
}

void Title::setText(const Text &text)
{
    if (!d) d = new TitlePrivate;
    d->text = text;
}

TextProperties Title::textProperties() const
{
    if (!d) return {};

    if (!d->text.isValid() || d->text.isPlainString() || d->text.isStringReference())
        return d->textProperties.textProperties();
    if (d->text.isRichString()) return d->text.textProperties();
    return {};
}

TextProperties &Title::textProperties()
{
    if (!d) d = new TitlePrivate;
    if (!d->text.isValid() || d->text.isPlainString() || d->text.isStringReference())
        return d->textProperties.textProperties();
    return d->text.textProperties();
}

void Title::setTextProperties(const TextProperties &textProperties)
{
    if (!d) d = new TitlePrivate;
    if (!d->text.isValid() || d->text.isPlainString() || d->text.isStringReference())
        d->textProperties.setTextProperties(textProperties);
    else d->text.setTextProperties(textProperties);
}

ParagraphProperties Title::defaultParagraphProperties() const
{
    if (!d) return {};

    if (!d->text.isValid() || d->text.isPlainString() || d->text.isStringReference())
        return d->textProperties.defaultParagraphProperties();
    if (d->text.isRichString()) return d->text.defaultParagraphProperties();
    return {};
}

ParagraphProperties &Title::defaultParagraphProperties()
{
    if (!d) d = new TitlePrivate;
    if (!d->text.isValid() || d->text.isPlainString() || d->text.isStringReference())
        return d->textProperties.defaultParagraphProperties();
    return d->text.defaultParagraphProperties();
}

void Title::setDefaultParagraphProperties(const ParagraphProperties &defaultParagraphProperties)
{
    if (!d) d = new TitlePrivate;
    if (!d->text.isValid() || d->text.isPlainString() || d->text.isStringReference())
        d->textProperties.setDefaultParagraphProperties(defaultParagraphProperties);
    else d->text.setDefaultParagraphProperties(defaultParagraphProperties);
}

CharacterProperties Title::defaultCharacterProperties() const
{
    if (!d) return {};

    if (!d->text.isValid() || d->text.isPlainString() || d->text.isStringReference())
        return d->textProperties.defaultCharacterProperties();
    if (d->text.isRichString()) return d->text.defaultCharacterProperties();
    return {};
}

CharacterProperties &Title::defaultCharacterProperties()
{
    if (!d) d = new TitlePrivate;
    if (!d->text.isValid())
        return d->textProperties.defaultCharacterProperties();
    if (d->text.type() != Text::Type::RichText)
        return d->textProperties.defaultCharacterProperties();

    return d->text.defaultCharacterProperties();
}

void Title::setDefaultCharacterProperties(const CharacterProperties &defaultCharacterProperties)
{
    if (!d) d = new TitlePrivate;
    if (!d->text.isValid() || d->text.isPlainString() || d->text.isStringReference())
        d->textProperties.setDefaultCharacterProperties(defaultCharacterProperties);
    else d->text.setDefaultCharacterProperties(defaultCharacterProperties);
}

Layout Title::layout() const
{
    if (d) return d->layout;
    return {};
}

Layout &Title::layout()
{
    if (!d) d = new TitlePrivate;
    return d->layout;
}

void Title::setLayout(const Layout &layout)
{
    if (!d) d = new TitlePrivate;
    d->layout = layout;
}

void Title::moveTo(const QPointF &point)
{
    if (!d) d = new TitlePrivate;
    d->layout.setPosition(point);
}

std::optional<bool> Title::overlay() const
{
    if (d) return d->overlay;
    return {};
}

void Title::setOverlay(bool overlay)
{
    if (!d) d = new TitlePrivate;
    d->overlay = overlay;
}

ShapeFormat Title::shape() const
{
    if (d) return d->shape;
    return {};
}

ShapeFormat &Title::shape()
{
    if (!d) d = new TitlePrivate;
    return d->shape;
}

void Title::setShape(const ShapeFormat &shape)
{
    if (!d) d = new TitlePrivate;
    d->shape = shape;
}

bool Title::isValid() const
{
    if (d) return true;
    return false;
}

void Title::write(QXmlStreamWriter &writer) const
{
    writer.writeStartElement(QLatin1String("c:title"));
    if (d->text.isValid()) d->text.write(writer, "c:tx");
    if (d->layout.isValid()) d->layout.write(writer, "c:layout");
    writeEmptyElement(writer, QLatin1String("c:overlay"), d->overlay);
    if (d->shape.isValid()) d->shape.write(writer, "a:spPr");

    if (d->textProperties.isValid())
        d->textProperties.write(writer, QLatin1String("c:txPr"), false);

    writer.writeEndElement(); // c:title
}

void Title::read(QXmlStreamReader &reader)
{
    if (!d) d = new TitlePrivate;
    const auto &name = reader.name();
    while (!reader.atEnd()) {
        auto token = reader.readNext();
        if (token == QXmlStreamReader::StartElement) {
            const auto &a = reader.attributes();
            if (reader.name() == QLatin1String("tx")) d->text.read(reader);
            if (reader.name() == QLatin1String("layout")) d->layout.read(reader);
            if (reader.name() == QLatin1String("overlay"))
                parseAttributeBool(a, QLatin1String("val"), d->overlay);
            if (reader.name() == QLatin1String("spPr")) d->shape.read(reader);
            if (reader.name() == QLatin1String("txPr")) {
                d->textProperties.read(reader, false);
            }
        }
        else if (token == QXmlStreamReader::EndElement && reader.name() == name) {
            break;
        }
    }
}

bool Title::operator ==(const Title &other) const
{
    if (d != other.d) return false;
    if (!d || !other.d) return false;
    return *d == *other.d;
}

bool Title::operator !=(const Title &other) const
{
    return !operator==(other);
}

QXlsx::Title::operator QVariant() const
{
    const auto& cref
#if QT_VERSION >= 0x060000 // Qt 6.0 or over
        = QMetaType::fromType<Title>();
#else
        = qMetaTypeId<Title>() ;
#endif
    return QVariant(cref, this);
}

bool TitlePrivate::operator==(const TitlePrivate &other) const
{
    return (text == other.text && layout == other.layout
            && overlay == other.overlay && shape == other.shape
            && textProperties == other.textProperties);
}

TitlePrivate::TitlePrivate()
{

}

TitlePrivate::TitlePrivate(const TitlePrivate &other) : QSharedData(other),
    text{other.text}, layout{other.layout}, overlay{other.overlay},
    shape{other.shape}, textProperties{other.textProperties}
{

}

TitlePrivate::~TitlePrivate()
{

}

QDebug operator<<(QDebug dbg, const Title &f)
{
    QDebugStateSaver saver(dbg);
//    dbg.setAutoInsertSpaces(false);

    dbg.nospace() << "QXlsx::Title(";
    if (f.isValid()) {
        dbg.nospace() << "text: ";
        if (f.d->text.isValid()) dbg.nospace() << f.d->text;
//        else dbg.nospace() << "not set";
        dbg.nospace() << ", ";
        dbg.nospace() << "layout: ";
        if (f.d->layout.isValid()) dbg.nospace() << f.d->layout;
        else dbg.nospace() << "not set";
        dbg.nospace() << ", ";
        dbg.nospace() << "overlay: ";
        if (f.d->overlay.has_value()) dbg.nospace() << f.d->overlay.value();
        else dbg.nospace() << "not set";
        dbg.nospace() << ", ";
        dbg.nospace() << "shape: ";
        if (f.d->shape.isValid()) dbg.nospace() << f.d->shape;
        else dbg.nospace() << "not set";
        dbg.nospace() << ", ";
        dbg.nospace() << "textProperties: ";
        if (f.d->textProperties.isValid()) dbg.nospace() << f.d->textProperties;
        else dbg.nospace() << "not set";
    }
    dbg.nospace() << ")";

    return dbg;
}

QT_END_NAMESPACE_XLSX
