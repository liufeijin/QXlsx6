#include "xlsxheaderfooter.h"
#include "xlsxutility_p.h"

namespace QXlsx {

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

void HeaderFooter::write(QXmlStreamWriter &writer, const QString &name) const
{
    if (isValid()) {
        writer.writeStartElement(name);
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
}
