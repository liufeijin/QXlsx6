// xlsxdocpropscore_p.h

#ifndef XLSXDOCPROPSCORE_H
#define XLSXDOCPROPSCORE_H

#include "xlsxglobal.h"
#include "xlsxabstractooxmlfile.h"
#include "xlsxdocument.h"

#include <QMap>
#include <QStringList>

class QIODevice;

namespace QXlsx {

class DocPropsCore : public AbstractOOXmlFile
{
public:
    explicit DocPropsCore(CreateFlag flag);

    void setProperty(Document::Metadata name, const QVariant &value);
    inline QMap<Document::Metadata, QVariant> properties() const {return m_properties;}
        
    void saveToXmlFile(QIODevice *device) const override;
    bool loadFromXmlFile(QIODevice *device) override;

private:
    QMap<Document::Metadata, QVariant> m_properties;
};

}

#endif // XLSXDOCPROPSCORE_H
