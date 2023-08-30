// xlsxabstractsheet.h

#ifndef XLSXABSTRACTSHEET_H
#define XLSXABSTRACTSHEET_H

#include "xlsxglobal.h"
#include "xlsxabstractooxmlfile.h"

namespace QXlsx {

class Workbook;
class Drawing;
class AbstractSheetPrivate;

/**
 * @brief The AbstractSheet class is a base class for different sheet types.
 */
class QXLSX_EXPORT AbstractSheet : public AbstractOOXmlFile
{
    Q_DECLARE_PRIVATE(AbstractSheet)

public:
    /**
     * @brief returns pointer to this sheet's workbook.
     * @return
     */
    Workbook *workbook() const;

public:
    /**
     * @brief The Type enum specifies the sheet type.
     */
    enum class Type {
        Worksheet, /**< Worksheet that contains cells with data etc.*/
        Chartsheet, /**< Chartsheet that contains a chart.*/
        Dialogsheet, /**< @note This sheet type is currently not supported. */
        Macrosheet /**< @note This sheet type is currently not supported.*/
    };
    /**
     * @brief The Visibility enum specifies the sheet visibility.
     */
    enum class Visibility {
        Visible, /**< The sheet is visible. */
        Hidden, /**< The sheet is hidden but can be shown via the UI.*/
        VeryHidden /**< The sheet is hidden and cannot be shown via the UI.
To set the sheet VeryHidden use #setVisibility() method.*/
    };

public:
    /**
     * @brief returns the sheet's name.
     * @return
     */
    QString name() const;
    /**
     * @brief returns the sheet's type.
     * @return
     */
    Type type() const;
    /**
     * @brief returns the sheet's visibility.
     * @return either Visibility::Visible, Visibility::Hidden or Visibility::VeryHidden.
     * Sheets are Visible by default.
     * @sa #isVisible(), #isHidden()
     */
    Visibility visibility() const;
    /**
     * @brief sets the sheet's visibility.
     * @param visibility
     * @note Use this method to set the sheet Visibility::VeryHidden.
     * @sa #setHidden(), #setVisible().
     */
    void setVisibility(Visibility visibility);
    /**
     * @brief returns whether the sheet is hidden or very hidden.
     * @return true if the sheet's #visibility() is Visibility::Hidden or Visibility::VeryHidden,
     * false otherwise.
     * This is a convenience method, equivalent to @code visibility() == Visibility::Hidden || visibility() == Visibility::VeryHidden @endcode
     * By default sheets are visible.
     * @sa #visibility(), #isVisible().
     */
    bool isHidden() const;
    /**
     * @brief returns whether the sheet is visible.
     * @return true if the sheet's #visibility() is Visibility::Visible, false otherwise.
     * This is a convenience method, equivalent to @code visibility() == Visibility::Visible @endcode
     * By default sheets are visible.
     * @sa #isHidden(), #visibility().
     */
    bool isVisible() const;
    /**
     * @brief sets the sheet's visibility.
     * @param hidden if true, the sheet's visibility will be set to Visibility::Hidden,
     * if false, the sheet will be Visibility::Visible.
     * @note If @a hidden is true, the very hidden sheet stays very hidden.
     * You can check the exact visibility with #visibility().
     * @sa #setVisibility(), #setVisible().
     */
    void setHidden(bool hidden);
    /**
     * @brief sets the sheet's visibility.
     * @param visible if true, the sheet's visibility will be set to Visibility::Visible,
     * if false, the sheet will be Visibility::Hidden.
     * @note If @a visible is false, the very hidden sheet stays very hidden.
     * You can check the exact visibility with #visibility().
     * @sa #setVisibility(), #setHidden().
     */
    void setVisible(bool visible);

protected:
    friend class Workbook;
    AbstractSheet(const QString &sheetName, int sheetId, Workbook *book, AbstractSheetPrivate *d);
    /**
     * @brief Copies the current sheet to a sheet called @a distName with @a distId.
     * Returns the new sheet.
     */
    virtual AbstractSheet *copy(const QString &distName, int distId) const = 0;
    void setName(const QString &sheetName);
    void setType(Type type);
    int id() const;

    Drawing *drawing() const;
};

}
#endif // XLSXABSTRACTSHEET_H
