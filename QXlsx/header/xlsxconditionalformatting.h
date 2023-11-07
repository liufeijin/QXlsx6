// xlsxconditionalformatting.h

#ifndef QXLSX_XLSXCONDITIONALFORMATTING_H
#define QXLSX_XLSXCONDITIONALFORMATTING_H

#include <QtGlobal>
#include <QString>
#include <QList>
#include <QColor>
#include <QXmlStreamReader>
#include <QXmlStreamWriter>
#include <QSharedDataPointer>

#include "xlsxglobal.h"
#include "xlsxcellrange.h"
#include "xlsxcellreference.h"

class ConditionalFormattingTest;

namespace QXlsx {

class Format;
class Worksheet;
class Styles;
class ConditionalFormattingPrivate;

//class QXLSX_EXPORT ConditionalFormattingRule
//{
//public:

//};

/**
 * @brief Conditional formatting for single cell or ranges
 *
 * A Conditional Format is a format, such as cell shading or font color, that a
 * spreadsheet application can automatically apply to cells if a specified
 * condition is true.
 */
class QXLSX_EXPORT ConditionalFormatting
{
public:
    /**
     * @brief The Type enum specifies the type of condition for the cells highlighting rules.
     */
    enum class Type {
        LessThan,
        LessThanOrEqual,
        Equal,
        NotEqual,
        GreaterThanOrEqual,
        GreaterThan,
        Between,
        NotBetween,

        ContainsText,
        NotContainsText,
        BeginsWith,
        EndsWith,

        TimePeriod,

        Duplicate,
        Unique,
        Blanks,
        NoBlanks,
        Errors,
        NoErrors,

        Top,
        TopPercent,
        Bottom,
        BottomPercent,

        AboveAverage,
        AboveOrEqualAverage,
        AboveStdDev1,
        AboveStdDev2,
        AboveStdDev3,
        BelowAverage,
        BelowOrEqualAverage,
        BelowStdDev1,
        BelowStdDev2,
        BelowStdDev3,

        Expression
    };
    /**
     * @brief The ValueObjectType enum specifies the type of condition for the data bar
     * and color scale rules.
     */
    enum class ValueObjectType
    {
        Formula,
        Max,
        Min,
        Num,
        Percent,
        Percentile
    };

public:
    ConditionalFormatting();
    ConditionalFormatting(const ConditionalFormatting &other);
    ~ConditionalFormatting();

public:
    /**
     * @brief adds the rule of highlighting cells.
     *
     * Depending of the @a type conditions @a formula1 and @a formula2 are
     * evaluated as follows:
     *
     * type | formula1 | formula2
     * ----|----|----
     * LessThan, LessThanOrEqual, Equal, NotEqual, GreaterThanOrEqual, GreaterThan | numeric value | ignored
     * Between, NotBetween | numeric value 1 | numeric value 2
     * ContainsText, NotContainsText, BeginsWith, EndsWith | text value | ignored
     * TimePeriod | not implemented | not implemented
     * Duplicate, Unique, Blanks, NoBlanks, Errors, NoErrors | ignored | ignored
     * Top, TopPercent, Bottom, BottomPercent | The value of "n" in a "top/bottom n" rule | ignored
     * AboveAverage, AboveOrEqualAverage, AboveStdDev1, AboveStdDev2, AboveStdDev3 | ignored | ignored
     * BelowAverage, BelowOrEqualAverage, BelowStdDev1, BelowStdDev2, BelowStdDev3 | ignored | ignored
     * Expression | used for expression | used for expression
     *
     * @param type The rule type.
     * @param formula1 The first (main) condition.
     * @param formula2 The second condition.
     * @param format valid Format to apply if the rule evaluates to true.
     * @param stopIfTrue If true, then no rules with lower priority shall be
     * applied over this rule, when this rule evaluates to true.
     * @return true if @a format is valid and adding the rule was successful.
     */
    bool addHighlightCellsRule(Type type,
                               const QString &formula1,
                               const QString &formula2,
                               const Format &format,
                               bool stopIfTrue = false);
    /**
     * @overload
     * @brief adds the rule of highlighting cells. Equivalent to `addHighlightCellsRule(type, "", "", format, stopIfTrue)`.
     * @param type The rule type.
     * @param format valid Format to apply if the rule evaluates to true.
     * @param stopIfTrue If true, then no rules with lower priority shall be
     * applied over this rule, when this rule evaluates to true.
     * @return true if @a format is valid and adding the rule was successful.
     */
    bool addHighlightCellsRule(Type type,
                               const Format &format,
                               bool stopIfTrue = false);
    /**
     * @overload
     * @brief adds the rule of highlighting cells. Equivalent to `addHighlightCellsRule(type, formula, "", format, stopIfTrue)`.
     * @param type The rule type.
     * @param formula The rule condition.
     * @param format valid Format to apply if the rule evaluates to true.
     * @param stopIfTrue If true, then no rules with lower priority shall be
     * applied over this rule, when this rule evaluates to true.
     * @return true if @a format is valid and adding the rule was successful.
     */
    bool addHighlightCellsRule(Type type,
                               const QString &formula,
                               const Format &format,
                               bool stopIfTrue = false);

    /**
     * @brief adds the rule of formatting cells with a color bar.
     * @param color
     * @param type1
     * @param val1
     * @param type2
     * @param val2
     * @param showData If true, then cell value will be shown over the color bar.
     * If false, then only the color bar will be visible.
     * @param stopIfTrue If true, then no rules with lower priority shall be
     * applied over this rule, when this rule evaluates to true.
     * @return true
     */
    bool addDataBarRule(const QColor &color,
                        ValueObjectType type1, const QString &val1,
                        ValueObjectType type2, const QString &val2,
                        bool showData = true,
                        bool stopIfTrue = false);
    /**
     * @brief adds the rule of formatting cells with a color bar.
     * Equivalent to `addDataBarRule(color,
     * @param color
     * @param showData
     * @param stopIfTrue
     * @return
     */
    bool addDataBarRule(const QColor &color, bool showData=true, bool stopIfTrue = false);
    bool add2ColorScaleRule(const QColor &minColor, const QColor &maxColor, bool stopIfTrue=false);
    bool add3ColorScaleRule(const QColor &minColor, const QColor &midColor, const QColor &maxColor,  bool stopIfTrue=false);

    /**
     * @brief returns the count of rules defined in this conditional formatting.
     */
    int rulesCount() const;
    /**
     * @brief removes all rules defined in this conditional formatting.
     */
    void clearRules();

    /**
     * @brief removes rule with @a ruleIndex.
     * @param ruleIndex zero-based index of the rule from 0 to #rulesCount()-1.
     * @return true if @a ruleIndex is valid and removing was successful.
     */
    bool removeRule(int ruleIndex);

    /**
     * @brief sets @a priority to all rules defined in this conditional formatting.
     * @param priority integer value >= 1. Lower numeric value are higher priority
     * than higher numeric value, where 1 is the highest priority. The default value is 1.
     * @return true if @a priority is valid and rules are not empty.
     */
    bool setRulesPriority(int priority);
    /**
     * @brief sets @a priority to the rule with @a ruleIndex defined in this conditional formatting.
     * @param ruleIndex zero-based index of the rule from 0 to #rulesCount()-1.
     * @param priority integer value >= 1. Lower numeric value are higher priority
     * than higher numeric value, where 1 is the highest priority. The default value is 1.
     * @return true if @a priority is valid and @a ruleIndex is valid.
     */
    bool setRulePriority(int ruleIndex, int priority);
    /**
     * @brief Sets the priority behavior when adding rules.
     * @param autodecrement If true, then each new added rule has lower priority
     * than the previously added one.
     *
     * The default behaviour is to set the highest priority (1) to all rules. You
     * can change the priority of the specific rule with #setRulePriority() or
     * set fixed priority to all rules with #setRulesPriority().
     */
    void setAutoDecrementPriority(bool autodecrement);
    /**
     * @brief updates priorities of the rules defined in this conditional formatting.
     * @param firstRuleHasMaximumPriority If true, then the first rule will have
     * the highest priority, the last rule will have the lowest priority. If
     * false, then the first rule will have the lowest priority, the last rule
     * will have the highest priority.
     */
    void updateRulesPriorities(bool firstRuleHasMaximumPriority = true);
    /**
     * @brief returns the priority of the rule with @a ruleIndex.
     * @param ruleIndex zero-based index of the rule from 0 to #rulesCount()-1.
     * @return valid int if @a ruleIndex is valid, nullopt otherwise.
     */
    std::optional<int> rulePriority(int ruleIndex) const;

    /**
     * @brief return the list of ranges conditional formatting applies to.
     * @return a list of CellRange.
     */
    QList<CellRange> ranges() const;
    /**
     * @brief adds the cell which the conditional formatting is applied to.
     * @param cell valid CellReference object.
     */
    void addCell(const CellReference &cell);
    /**
     * @overload
     * @brief adds the cell which the conditional formatting is applied to.
     * @param row the (1-based) cell row number.
     * @param col the (1-based) cell column number.
     */
    void addCell(int row, int col);
    /**
     * @brief adds cell range which the conditional formatting is applied to.
     * @param range valid CellRange object.
     */
    void addRange(const CellRange &range);
    /**
     * @overload
     * @brief adds cell range which the conditional formatting is applied to.
     * @param firstRow the (1-based) first row of the range.
     * @param firstCol the (1-based) first column of the range.
     * @param lastRow the (1-based) last row of the range.
     * @param lastCol the (1-based) last column of the range.
     */
    void addRange(int firstRow, int firstCol, int lastRow, int lastCol);

    /**
     * @brief removes any previously added ranges and adds @a range.
     * @param range valid CellRange object.
     */
    void setRange(const CellRange &range);
    /**
     * @brief removes any previously added ranges.
     *
     * @warning If no ranges are defined ConditionalFormatting is invalid.
     */
    void clearRanges();

    //needed by QSharedDataPointer!!
    ConditionalFormatting &operator=(const ConditionalFormatting &other);

private:
    friend class Worksheet;
    friend class ::ConditionalFormattingTest;

private:
    bool saveToXml(QXmlStreamWriter &writer) const;
    bool loadFromXml(QXmlStreamReader &reader, Styles* styles = NULL);

    QSharedDataPointer<ConditionalFormattingPrivate> d;
};

}

Q_DECLARE_TYPEINFO(QXlsx::ConditionalFormatting, Q_MOVABLE_TYPE);

#endif // QXLSX_XLSXCONDITIONALFORMATTING_H
