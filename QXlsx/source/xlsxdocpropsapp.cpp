// xlsxdocpropsapp.cpp

#include "xlsxdocpropsapp_p.h"
#include "xlsxutility_p.h"

#include <QXmlStreamWriter>
#include <QXmlStreamReader>
#include <QDir>
#include <QFile>
#include <QDateTime>
#include <QVariant>
#include <QBuffer>

namespace QXlsx {

DocPropsApp::DocPropsApp(CreateFlag flag)
    :AbstractOOXmlFile(flag)
{
}

void DocPropsApp::addPartTitle(const QString &title)
{
    m_titlesOfPartsList.append(title);
}

void DocPropsApp::addHeadingPair(const QString &name, int value)
{
    m_headingPairsList.append({ name, value });
}

void DocPropsApp::setProperty(Document::Metadata name, const QVariant &value)
{
    if (!value.isValid())
        m_properties.remove(name);
    else
        m_properties[name] = value;
}

void DocPropsApp::saveToXmlFile(QIODevice *device) const
{
    QXmlStreamWriter writer(device);
    QString vt = QStringLiteral("http://schemas.openxmlformats.org/officeDocument/2006/docPropsVTypes");

    writer.writeStartDocument(QStringLiteral("1.0"), true);
    writer.writeStartElement(QStringLiteral("Properties"));
    writer.writeDefaultNamespace(QStringLiteral("http://schemas.openxmlformats.org/officeDocument/2006/extended-properties"));
    writer.writeNamespace(vt, QStringLiteral("vt"));
    writer.writeTextElement(QStringLiteral("Application"),
                            m_properties.value(Document::Metadata::Application, QStringLiteral("Microsoft Excel")).toString());
    writer.writeTextElement(QStringLiteral("DocSecurity"), QStringLiteral("0"));
    writer.writeTextElement(QStringLiteral("ScaleCrop"), QStringLiteral("false"));

    writer.writeStartElement(QStringLiteral("HeadingPairs"));
    writer.writeStartElement(vt, QStringLiteral("vector"));
    writer.writeAttribute(QStringLiteral("size"), QString::number(m_headingPairsList.size()*2));
    writer.writeAttribute(QStringLiteral("baseType"), QStringLiteral("variant"));

    for (const auto &pair : qAsConst(m_headingPairsList)) {
        writer.writeStartElement(vt, QStringLiteral("variant"));
        writer.writeTextElement(vt, QStringLiteral("lpstr"), pair.first);
        writer.writeEndElement(); //vt:variant
        writer.writeStartElement(vt, QStringLiteral("variant"));
        writer.writeTextElement(vt, QStringLiteral("i4"), QString::number(pair.second));
        writer.writeEndElement(); //vt:variant
    }
    writer.writeEndElement();//vt:vector
    writer.writeEndElement();//HeadingPairs

    writer.writeStartElement(QStringLiteral("TitlesOfParts"));
    writer.writeStartElement(vt, QStringLiteral("vector"));
    writer.writeAttribute(QStringLiteral("size"), QString::number(m_titlesOfPartsList.size()));
    writer.writeAttribute(QStringLiteral("baseType"), QStringLiteral("lpstr"));
    for (const QString &title : qAsConst(m_titlesOfPartsList))
        writer.writeTextElement(vt, QStringLiteral("lpstr"), title);
    writer.writeEndElement();//vt:vector
    writer.writeEndElement();//TitlesOfParts

    writeTextElement(writer, QLatin1String("Manager"), m_properties.value(Document::Metadata::Manager).toString());
    writer.writeTextElement(QStringLiteral("Company"), m_properties.value(Document::Metadata::Company).toString());
    writer.writeTextElement(QStringLiteral("LinksUpToDate"), QStringLiteral("false"));
    writer.writeTextElement(QStringLiteral("SharedDoc"), QStringLiteral("false"));
    writer.writeTextElement(QStringLiteral("HyperlinksChanged"), QStringLiteral("false"));

    writer.writeTextElement(QStringLiteral("AppVersion"),
                            m_properties.value(Document::Metadata::AppVersion, QStringLiteral("12.0000")).toString());

    writer.writeEndElement(); //Properties
    writer.writeEndDocument();
}

bool DocPropsApp::loadFromXmlFile(QIODevice *device)
{
    QXmlStreamReader reader(device);
    while (!reader.atEnd()) {
         QXmlStreamReader::TokenType token = reader.readNext();
         if (token == QXmlStreamReader::StartElement) {
             if (reader.name() == QStringLiteral("Manager"))
                 setProperty(Document::Metadata::Manager, reader.readElementText());
             else if (reader.name() == QStringLiteral("Company"))
                 setProperty(Document::Metadata::Company, reader.readElementText());
             else if (reader.name() == QStringLiteral("Application"))
                 setProperty(Document::Metadata::Application, reader.readElementText());
             else if (reader.name() == QStringLiteral("AppVersion"))
                 setProperty(Document::Metadata::AppVersion, reader.readElementText());
         }

         if (reader.hasError()) {
             qDebug("Error when read doc props app file.");
         }
    }
    return true;
}

}
