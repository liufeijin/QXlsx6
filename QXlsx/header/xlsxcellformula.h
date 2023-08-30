// xlsxcellformula.h

#ifndef QXLSX_XLSXCELLFORMULA_H
#define QXLSX_XLSXCELLFORMULA_H

#include "xlsxglobal.h"

#include <QExplicitlySharedDataPointer>

class QXmlStreamWriter;
class QXmlStreamReader;

namespace QXlsx {

class CellFormulaPrivate;
class CellRange;
class Worksheet;
class WorksheetPrivate;

class QXLSX_EXPORT CellFormula
{
public:
    enum class Type {
        Normal,
        Array,
        DataTable,
        Shared
    };

public:
    CellFormula();
    CellFormula(const char *formula, Type type=Type::Normal);
    CellFormula(const QString &formula, Type type=Type::Normal);
    CellFormula(const QString &formula, const CellRange &ref, Type type);
    CellFormula(const CellFormula &other);
    ~CellFormula();

public:
    CellFormula &operator =(const CellFormula &other);
    bool isValid() const;

    Type formulaType() const;
    QString formulaText() const;
    CellRange reference() const;
    int sharedIndex() const;

    bool operator == (const CellFormula &formula) const;
    bool operator != (const CellFormula &formula) const;

    bool saveToXml(QXmlStreamWriter &writer) const;
    bool loadFromXml(QXmlStreamReader &reader);
private:
    friend class Worksheet;
    friend class WorksheetPrivate;
    QExplicitlySharedDataPointer<CellFormulaPrivate> d;
};

}

#endif // QXLSX_XLSXCELLFORMULA_H
