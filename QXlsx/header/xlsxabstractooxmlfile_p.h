// xlsxabstractooxmlfile_p.h

#ifndef XLSXOOXMLFILE_P_H
#define XLSXOOXMLFILE_P_H

#include "xlsxglobal.h"

#include "xlsxabstractooxmlfile.h"
#include "xlsxrelationships_p.h"

namespace QXlsx {

class AbstractOOXmlFilePrivate
{
    Q_DECLARE_PUBLIC(AbstractOOXmlFile)

public:
    AbstractOOXmlFilePrivate(AbstractOOXmlFile* q, AbstractOOXmlFile::CreateFlag flag);
    virtual ~AbstractOOXmlFilePrivate();

public:
    QString filePathInPackage; //such as "xl/worksheets/sheet1.xml"

    Relationships *relationships = nullptr;
    AbstractOOXmlFile::CreateFlag flag;
    AbstractOOXmlFile *q_ptr = nullptr;
};

}

#endif // XLSXOOXMLFILE_P_H
