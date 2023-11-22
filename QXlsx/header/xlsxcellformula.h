// xlsxcellformula.h

#ifndef QXLSX_XLSXCELLFORMULA_H
#define QXLSX_XLSXCELLFORMULA_H

#include "xlsxglobal.h"

#include <QSharedDataPointer>

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
    /**
     * @brief The Type enum specifies the formula type.
     */
    enum class Type {
        Normal, /**< A single-cell formula */
        Array, /**< An array formula */
        DataTable, /**< Table formula */
        Shared /**< Shared formula */
    };

public:
    /**
     * @brief Creates an invalid CellFormula.
     */
    CellFormula();
    /**
     * @brief Creates a CellFormula object of specified @a type with specified @a formula.
     * @param formula A string containing the formula text. The trailing '=' will be removed.
     * @param type The formula type.
     */
    CellFormula(const char *formula, Type type = Type::Normal);
    /**
     * @overload
     * @brief Creates a CellFormula object of specified @a type with specified @a formula.
     * @param formula A string containing the formula text. The trailing '=' will be removed.
     * @param type The formula type.
     */
    CellFormula(const QString &formula, Type type = Type::Normal);
    /**
     * @brief Creates a CellFormula object for an array, table or shared formula
     * with specified @a formula text, @a range to apply the formula to and @a type.
     * @param formula A string containing the formula text. The trailing = will be removed.
     * @param range The range the formula is applied to.
     * @param type The formula type. Can be Type::Array, Type::DataTable or Type::Shared.
     */
    CellFormula(const QString &formula, const CellRange &range, Type type);
    /**
     * @brief Creates a copy of @a other CellFormula.
     */
    CellFormula(const CellFormula &other);
    ~CellFormula();

public:
    CellFormula &operator =(const CellFormula &other);
    bool isValid() const;

    /**
     * @brief returns the formula type.
     * @return CellFormula::Type enum value if the type is set, `nullopt` otherwise.
     *
     * If formula type is not set, CellFormula::Type::Normal is assumed.
     */
    std::optional<Type> type() const;
    /**
     * @brief returns the formula text (without starting =).
     * @return QString object. Text may be empty.
     */
    QString text() const;
    /**
     * @brief Returns the reference range of the shared, array or table formulas.
     * For normal formula this will return an invalid CellRange object.
     * @return CellRange object.
     */
    CellRange reference() const;
    /**
     * @brief Sets the reference range of the shared, array or table formulas.
     *
     * If #type() is Type::Normal, does nothing.
     *
     * @param range A valid CellRange.
     */
    void setReference(const CellRange &range);
    /**
     * @brief Returns the shared index for shared formula.
     *
     * This index is used to locate the group to which this particular cell's formula belongs.
     *
     * @return integer shared index if formula type is Type::Shared, `nullopt` otherwise.
     * The index may still be `nullopt` even if the formula is shared.
     */
    std::optional<int> sharedIndex() const;
    /**
     * @brief Sets the shared index for shared formula.
     *
     * This index is used to locate the group to which this particular cell's formula belongs.
     *
     * @param sharedIndex integer shared index that references the 'master' formula.
     */
    void setSharedIndex(int sharedIndex);
    /**
     * @brief Returns whether a flag was set that the formula needs to be
     * recalculated the next time calculation is performed.
     *
     * The default value is `false`.
     *
     * @return bool value if the flag was set, `nullopt` otherwise.
     * @note By default if the parameter is not set by the moment the document
     * is being saved (#needsRecalculation() returns `nullopt`), this library
     * invokes `#setNeedsRecalculation(true)`. If you don't need recalculation,
     * manually set the flag to `false`.
     */
    std::optional<bool> needsRecalculation() const;
    /**
     * @brief sets a flag whether the formula needs to be recalculated the next
     * time calculation is performed.
     * @param recalculate If `true`, the formula will be recalculated automatically.
     *
     * If the parameter is not set, `false` is assumed (the formula doesn't need
     * recalculation).
     * @note By default if the parameter is not set by the moment the document
     * is being saved (#needsRecalculation() returns `nullopt`), this library
     * invokes `#setNeedsRecalculation(true)`. This is an inherited behavior that
     * can be changed in the future releases. If you don't need recalculation,
     * manually set the flag to `false`.
     */
    void setNeedsRecalculation(bool recalculate);

    bool operator == (const CellFormula &formula) const;
    bool operator != (const CellFormula &formula) const;
    operator QVariant() const;

    bool saveToXml(QXmlStreamWriter &writer) const;
    bool loadFromXml(QXmlStreamReader &reader);
private:
    friend class Worksheet;
    friend class WorksheetPrivate;
    QSharedDataPointer<CellFormulaPrivate> d;
};

}

Q_DECLARE_METATYPE(QXlsx::CellFormula)
Q_DECLARE_TYPEINFO(QXlsx::CellFormula, Q_MOVABLE_TYPE);

#endif // QXLSX_XLSXCELLFORMULA_H
