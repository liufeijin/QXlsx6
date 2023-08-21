// xlsxdrwaing_p.h

#ifndef QXLSX_DRAWING_H
#define QXLSX_DRAWING_H

#include <QtGlobal>
#include <QList>
#include <QString>

#include "xlsxrelationships_p.h"
#include "xlsxabstractooxmlfile.h"

class QIODevice;
class QXmlStreamWriter;

namespace QXlsx {

class DrawingAnchor;
class Workbook;
class AbstractSheet;
class MediaFile;

class Drawing : public AbstractOOXmlFile
{
public:
    Drawing(AbstractSheet *sheet, CreateFlag flag);
    ~Drawing();
    void saveToXmlFile(QIODevice *device) const override;
    bool loadFromXmlFile(QIODevice *device) override;

    AbstractSheet *sheet;
    Workbook *workbook;
    QList<DrawingAnchor *> anchors;
};

}

#endif // QXLSX_DRAWING_H
