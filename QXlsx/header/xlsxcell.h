// xlsxcell.h

#ifndef QXLSX_XLSXCELL_H
#define QXLSX_XLSXCELL_H

#include <cstdio>

#include <QtGlobal>
#include <QObject>
#include <QString>
#include <QVariant>
#include <QDate>
#include <QDateTime>
#include <QTime>

#include "xlsxglobal.h"
#include "xlsxformat.h"
#include "xlsxutility_p.h"
#include "xlsxrichstring.h"
#include "xlsxcellformula.h"

namespace QXlsx {

class Worksheet;
class Format;
class CellFormula;
class CellPrivate;
class WorksheetPrivate;

/**
 * @brief The Cell class represents one cell on a worksheet.
 */
class QXLSX_EXPORT Cell
{
    Q_DECLARE_PRIVATE(Cell)

private:
    friend class Worksheet;
    friend class WorksheetPrivate;

public:
    /**
     * @brief The CellType enum specifies the cell's data type.
     */
    enum class Type
    {
        Custom, /**< Custom cell type, not defined in the ECMA-376 */
        Boolean, /**< Cell contains a boolean */
        Date, /**< Cell contains a date in the ISO 8601 format. */
        Error, /**< Cell contains an error. */
        InlineString, /**< Cell contains an (inline) rich string, i.e., one not in
the shared string table. The string is written in the cell itself, with possible
duplication of data. */
        Number, /**< Cell contains a number */
        SharedString, /**< Cell contains a shared string. The actual string is stored not in the
cell but in the table of shared strings, thus reducing the document size. */
        String, /**< Cell contains a formula string */
    };

public:
    Cell(const QVariant &data = QVariant(),
         Type type = Type::Number,
         const Format &format = Format(),
         Worksheet *parent = nullptr,
         qint32 styleIndex = -1,
         const RichString &richString = {});
    Cell(const Cell * const cell, Worksheet *parent = nullptr);
    ~Cell();

public:
    CellPrivate * const d_ptr; // See D-pointer and Q-pointer of Qt, for more information.

public:
    Type type() const;
    QVariant value() const;
    QVariant readValue() const;
    void setValue(const QVariant &value);

    Format format() const;
    void setFormat(const Format &format);
    
    bool hasFormula() const;
    CellFormula formula() const;
    void setFormula(const CellFormula &formula);

    bool isDateTime() const;
    QVariant dateTime() const; // QDateTime, QDate, QTime

    bool isRichString() const;
    RichString richString() const;
    void setRichString(const RichString &richString);

    qint32 styleNumber() const;

    static bool isDateType(Type cellType, const Format &format);
private:
    SERIALIZE_ENUM(Type, {
        {Type::Custom, "c"},
        {Type::Date, "d"},
        {Type::Error, "e"},
        {Type::Number, "n"},
        {Type::String, "str"},
        {Type::Boolean, "b"},
        {Type::InlineString, "inlineStr"},
        {Type::SharedString, "s"},
    });
};

}

#endif // QXLSX_XLSXCELL_H
