// xlsxdocpropsapp_p.h

#ifndef XLSXDOCPROPSAPP_H
#define XLSXDOCPROPSAPP_H

#include <QList>
#include <QStringList>
#include <QMap>

#include "xlsxglobal.h"
#include "xlsxabstractooxmlfile.h"
#include "xlsxdocument.h"

class QIODevice;

namespace QXlsx {

class  DocPropsApp : public AbstractOOXmlFile
{
public:
    DocPropsApp(CreateFlag flag);
    
    void addPartTitle(const QString &title);
    void addHeadingPair(const QString &name, int value);

    void setProperty(Document::Metadata name, const QVariant &value);
    QMap<Document::Metadata, QVariant> properties() const
    {
        return m_properties;
    }

    void saveToXmlFile(QIODevice *device) const override;
    bool loadFromXmlFile(QIODevice *device) override;

private:
    QStringList m_titlesOfPartsList;
    QList<std::pair<QString, int> > m_headingPairsList;
    QMap<Document::Metadata, QVariant> m_properties;
};

}

#endif // XLSXDOCPROPSAPP_H
