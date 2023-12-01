// xlsxsharedstrings.cpp

#include <QtGlobal>
#include <QXmlStreamWriter>
#include <QXmlStreamReader>
#include <QDir>
#include <QFile>
#include <QDebug>
#include <QBuffer>

#include "xlsxrichstring.h"
#include "xlsxsharedstrings_p.h"
#include "xlsxutility_p.h"
#include "xlsxformat_p.h"
#include "xlsxcolor.h"

namespace QXlsx {

/*
 * Note that, when we open an existing .xlsx file (broken file?),
 * duplicated string items may exist in the shared string table.
 *
 * In such case, the size of stringList will larger than stringTable.
 * Duplicated items can be removed once we loaded all the worksheets.
 */

SharedStrings::SharedStrings(CreateFlag flag)
    :AbstractOOXmlFile(flag)
{
    m_stringCount = 0;
}

int SharedStrings::count() const
{
    return m_stringCount;
}

bool SharedStrings::isEmpty() const
{
    return m_stringList.isEmpty();
}

int SharedStrings::addSharedString(const QString &string)
{
    return addSharedString(RichString(string));
}

int SharedStrings::addSharedString(const RichString &string)
{
    m_stringCount += 1;

    auto it = m_stringTable.find(string);
    if (it != m_stringTable.end()) {
        it->count += 1;
        return it->index;
    }

    int index = m_stringList.size();
    m_stringTable[string] = XlsxSharedStringInfo(index);
    m_stringList.append(string);
    return index;
}

void SharedStrings::incRefByStringIndex(int idx)
{
    if (idx <0 || idx >= m_stringList.size()) {
        qDebug("SharedStrings: invlid index");
        return;
    }

    addSharedString(m_stringList[idx]);
}

/*
 * Broken, don't use.
 */
void SharedStrings::removeSharedString(const QString &string)
{
    removeSharedString(RichString(string));
}

/*
 * Broken, don't use.
 */
void SharedStrings::removeSharedString(const RichString &string)
{
    auto it = m_stringTable.find(string);
    if (it == m_stringTable.end())
        return;

    m_stringCount -= 1;

    it->count -= 1;

    if (it->count <= 0) {
        for (int i=it->index+1; i<m_stringList.size(); ++i)
            m_stringTable[m_stringList[i]].index -= 1;

        m_stringList.removeAt(it->index);
        m_stringTable.remove(string);
    }
}

std::optional<int> SharedStrings::getSharedStringIndex(const RichString &string) const
{
    auto it = m_stringTable.constFind(string);
    if (it != m_stringTable.constEnd())
        return it->index;
    return {};
}

RichString SharedStrings::getSharedString(int index) const
{
    if (index < m_stringList.count() && index >= 0)
        return m_stringList[index];
    return RichString();
}

QList<RichString> SharedStrings::getSharedStrings() const
{
    return m_stringList;
}

void SharedStrings::saveToXmlFile(QIODevice *device) const
{
    QXmlStreamWriter writer(device);

    if (m_stringList.size() != m_stringTable.size()) {
        //Duplicated string items exist in m_stringList
        //Clean up can not be done here, as the indices
        //have been used when we save the worksheets part.
    }

    writer.writeStartDocument(QStringLiteral("1.0"), true);
    writer.writeStartElement(QStringLiteral("sst"));
    writer.writeAttribute(QStringLiteral("xmlns"), QStringLiteral("http://schemas.openxmlformats.org/spreadsheetml/2006/main"));
    writer.writeAttribute(QStringLiteral("count"), QString::number(m_stringCount));
    writer.writeAttribute(QStringLiteral("uniqueCount"), QString::number(m_stringList.size()));

    for (const RichString &string : qAsConst(m_stringList)) {
        string.write(writer, QStringLiteral("si"));
    }

    writer.writeEndElement(); //sst
    writer.writeEndDocument();
}

void SharedStrings::readString(QXmlStreamReader &reader)
{
    Q_ASSERT(reader.name() == QLatin1String("si"));

    RichString richString;
    richString.read(reader, QLatin1String("si"));

    int idx = m_stringList.size();
    m_stringTable[richString] = XlsxSharedStringInfo(idx, 0);
    m_stringList.append(richString);
}

bool SharedStrings::loadFromXmlFile(QIODevice *device)
{
    QXmlStreamReader reader(device);
    int count = 0;
    bool hasUniqueCountAttr=true;
    while (!reader.atEnd()) {
         QXmlStreamReader::TokenType token = reader.readNext();
         if (token == QXmlStreamReader::StartElement) {
             if (reader.name() == QLatin1String("sst")) {
                 QXmlStreamAttributes attributes = reader.attributes();
                 if ((hasUniqueCountAttr = attributes.hasAttribute(QLatin1String("uniqueCount"))))
                     count = attributes.value(QLatin1String("uniqueCount")).toInt();
             } else if (reader.name() == QLatin1String("si")) {
                 readString(reader);
             }
         }
    }

    if (hasUniqueCountAttr && m_stringList.size() != count) {
        qDebug("Error: Shared string count");
        return false;
    }

    if (m_stringList.size() != m_stringTable.size()) {
        //qDebug("Warning: Duplicated items exist in shared string table.");
        //Nothing we can do here, as indices of the strings will be used when loading sheets.
    }

    return true;
}

}
