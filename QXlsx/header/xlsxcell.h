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
        InlineString, /**< Cell containing an (inline) rich string, i.e., one not in
                            the shared string table.  If this cell type is used, then
                            the cell value is in the _is_ element rather than the _v_
                            element in the cell (c element) */
        Number, /**< Cell contains a number */
        SharedString, /**< Cell contains a shared string */
        String, /**< Cell contains a formula string */
    };

public:
    Cell(const QVariant &data = QVariant(),
            Type type = Type::Number,
            const Format &format = Format(),
            Worksheet *parent = nullptr,
            qint32 styleIndex = (-1) );
    Cell(const Cell * const cell);
    ~Cell();

public:
    CellPrivate * const d_ptr; // See D-pointer and Q-pointer of Qt, for more information.

public:
    Type type() const;
    QVariant value() const;
    QVariant readValue() const;
    Format format() const;
    
    bool hasFormula() const;
    CellFormula formula() const;

    bool isDateTime() const;
    QVariant dateTime() const; // QDateTime, QDate, QTime

    bool isRichString() const;

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
