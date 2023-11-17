// xlsxdocpropscore.cpp

#include <QtGlobal>
#include <QXmlStreamWriter>
#include <QXmlStreamReader>
#include <QDir>
#include <QFile>
#include <QDateTime>
#include <QDebug>
#include <QBuffer>

#include "xlsxdocpropscore_p.h"
#include "xlsxutility_p.h"

namespace QXlsx {

DocPropsCore::DocPropsCore(CreateFlag flag)
    :AbstractOOXmlFile(flag)
{
}

void DocPropsCore::setProperty(Document::Metadata name, const QVariant &value)
{
    if (!value.isValid())
        m_properties.remove(name);
    else
        m_properties[name] = value;
}

//QStringList DocPropsCore::propertyNames() const
//{
//    return m_properties.keys();
//}

void DocPropsCore::saveToXmlFile(QIODevice *device) const
{
    QXmlStreamWriter writer(device);
    const QString cp = QStringLiteral("http://schemas.openxmlformats.org/package/2006/metadata/core-properties");
    const QString dc = QStringLiteral("http://purl.org/dc/elements/1.1/");
    const QString dcterms = QStringLiteral("http://purl.org/dc/terms/");
    const QString dcmitype = QStringLiteral("http://purl.org/dc/dcmitype/");
    const QString xsi = QStringLiteral("http://www.w3.org/2001/XMLSchema-instance");
    writer.writeStartDocument(QStringLiteral("1.0"), true);
    writer.writeStartElement(QStringLiteral("cp:coreProperties"));
    writer.writeNamespace(cp, QStringLiteral("cp"));
    writer.writeNamespace(dc, QStringLiteral("dc"));
    writer.writeNamespace(dcterms, QStringLiteral("dcterms"));
    writer.writeNamespace(dcmitype, QStringLiteral("dcmitype"));
    writer.writeNamespace(xsi, QStringLiteral("xsi"));

    writeTextElement(writer, QLatin1String("dc:title"), m_properties.value(Document::Metadata::Title).toString());
    writeTextElement(writer, QLatin1String("dc:subject"), m_properties.value(Document::Metadata::Subject).toString());

    QString creator = m_properties.value(Document::Metadata::Creator).toString();
    writeTextElement(writer, QLatin1String("dc:creator"), creator.isEmpty() ? QStringLiteral("QXlsx Library") : creator);

    writeTextElement(writer, QLatin1String("cp:keywords"), m_properties.value(Document::Metadata::Keywords).toString());
    writeTextElement(writer, QLatin1String("dc:description"), m_properties.value(Document::Metadata::Description).toString());

    QString lastModifiedBy = m_properties.value(Document::Metadata::LastModifiedBy).toString();
    writeTextElement(writer, QLatin1String("cp:lastModifiedBy"), lastModifiedBy.isEmpty() ? QStringLiteral("QXlsx Library") : lastModifiedBy);

    auto created = m_properties.value(Document::Metadata::Created);
    if (!created.isValid()) created = QDateTime::currentDateTime();
    writer.writeStartElement(dcterms, QStringLiteral("created"));
    writer.writeAttribute(xsi, QStringLiteral("type"), QStringLiteral("dcterms:W3CDTF"));
    writer.writeCharacters(created.toDateTime().toString(Qt::ISODate));
    writer.writeEndElement();//dcterms:created

    writer.writeStartElement(dcterms, QStringLiteral("modified"));
    writer.writeAttribute(xsi, QStringLiteral("type"), QStringLiteral("dcterms:W3CDTF"));
    writer.writeCharacters(QDateTime::currentDateTime().toString(Qt::ISODate));
    writer.writeEndElement();//dcterms:created

    writeTextElement(writer, QLatin1String("cp:category"), m_properties.value(Document::Metadata::Category).toString());
    writeTextElement(writer, QLatin1String("cp:contentStatus"), m_properties.value(Document::Metadata::ContentStatus).toString());
    writeTextElement(writer, QLatin1String("dc:identifier"), m_properties.value(Document::Metadata::Identifier).toString());
    writeTextElement(writer, QLatin1String("dc:language"), m_properties.value(Document::Metadata::Language).toString());

    auto lastPrinted = m_properties.value(Document::Metadata::LastPrinted);
    if (lastPrinted.isValid()) {
        writer.writeStartElement(dcterms, QStringLiteral("lastPrinted"));
        writer.writeAttribute(xsi, QStringLiteral("type"), QStringLiteral("dcterms:W3CDTF"));
        writer.writeCharacters(lastPrinted.toDateTime().toString(Qt::ISODate));
        writer.writeEndElement();
    }

    writeTextElement(writer, QLatin1String("cp:revision"), m_properties.value(Document::Metadata::Revision).toString());
    writeTextElement(writer, QLatin1String("cp:version"), m_properties.value(Document::Metadata::Version).toString());

    writer.writeEndElement(); //cp:coreProperties
    writer.writeEndDocument();
}

bool DocPropsCore::loadFromXmlFile(QIODevice *device)
{
    QXmlStreamReader reader(device);

    const QString cp = QStringLiteral("http://schemas.openxmlformats.org/package/2006/metadata/core-properties");
    const QString dc = QStringLiteral("http://purl.org/dc/elements/1.1/");
    const QString dcterms = QStringLiteral("http://purl.org/dc/terms/");

    while (!reader.atEnd()) {
         auto token = reader.readNext();

         if (token == QXmlStreamReader::StartElement) {

             const auto& nsUri = reader.namespaceUri();
             const auto& name = reader.name();

             if (name == QStringLiteral("subject") && nsUri == dc)
                 setProperty(Document::Metadata::Subject, reader.readElementText());
             else if (name == QStringLiteral("title") && nsUri == dc)
                 setProperty(Document::Metadata::Title, reader.readElementText());
             else if (name == QStringLiteral("creator") && nsUri == dc)
                 setProperty(Document::Metadata::Creator, reader.readElementText());
             else if (name == QStringLiteral("description") && nsUri == dc)
                 setProperty(Document::Metadata::Description, reader.readElementText());
             else if (name == QStringLiteral("keywords") && nsUri == cp)
                 setProperty(Document::Metadata::Keywords, reader.readElementText());
             else if (name == QStringLiteral("created") && nsUri == dcterms)
                 setProperty(Document::Metadata::Created, QDateTime::fromString(reader.readElementText(), Qt::ISODate));
             else if (name == QStringLiteral("category") && nsUri == cp)
                 setProperty(Document::Metadata::Category, reader.readElementText());
             else if (name == QStringLiteral("contentStatus") && nsUri == cp)
                 setProperty(Document::Metadata::ContentStatus, reader.readElementText());
             else if (name == QStringLiteral("identifier") && nsUri == dc)
                 setProperty(Document::Metadata::Identifier, reader.readElementText());
             else if (name == QStringLiteral("language") && nsUri == dc)
                 setProperty(Document::Metadata::Language, reader.readElementText());
             else if (name == QStringLiteral("lastModifiedBy") && nsUri == cp)
                 setProperty(Document::Metadata::LastModifiedBy, reader.readElementText());
             else if (name == QStringLiteral("lastPrinted") && nsUri == cp)
                 setProperty(Document::Metadata::LastPrinted, QDateTime::fromString(reader.readElementText(), Qt::ISODate));
             else if (name == QStringLiteral("revision") && nsUri == cp)
                 setProperty(Document::Metadata::Revision, reader.readElementText());
             else if (name == QStringLiteral("version") && nsUri == cp)
                 setProperty(Document::Metadata::Version, reader.readElementText());
         }

         if (reader.hasError())
             qDebug() << "Error when read doc props core file." << reader.errorString();
    }
    return true;
}

}
