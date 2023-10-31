#ifndef XLSXSHEETPROTECTION_H
#define XLSXSHEETPROTECTION_H

#include "xlsxglobal.h"

#include <QXmlStreamWriter>
#include <QSharedData>

namespace QXlsx {

class SheetProtectionPrivate;

/**
 * @brief The SheetProtection class specifies the sheet protection parameters for worksheets
 * and chartsheets.
 *
 * Not all parameters are applicable to chartsheets. See the description of parameters.
 * If a non-applicable parameter is set for a chartsheet protection, it will be ignored
 * when saving the document.
 */
class QXLSX_EXPORT SheetProtection
{
public:
    SheetProtection();
    SheetProtection(const SheetProtection &other);
    ~SheetProtection();
    SheetProtection &operator=(const SheetProtection &other);

    bool operator == (const SheetProtection &other) const;
    bool operator != (const SheetProtection &other) const;

    operator QVariant() const;

    /**
     * @brief returns the specific cryptographic hashing algorithm which was
     * used along with the salt attribute and input password in order to
     * compute the hash value.
     * @return QString object. May be empty.
     */
    QString algorithmName() const;
    /**
     * @brief sets the specific cryptographic hashing algorithm which shall
     * be used along with the salt attribute and input password in order to
     * compute the hash value.
     *
     * This is an optional parameter. If algorithmName is empty, it will not be stored in the doc.
     *
     * @param name The algorithm name. The following names are reserved:
     *
     * Algorithm | Description
     * ----|----
     * MD2   |Specifies that the MD2 algorithm, as defined by RFC 1319, shall be used.
     * MD4 |  Specifies that the MD4 algorithm, as defined by RFC 1320, shall be used.
     * MD5 | Specifies that the MD5 algorithm, as defined by RFC 1321, shall be used.
     * RIPEMD-128 | Specifies that the RIPEMD-128 algorithm, as defined by ISO/IEC 10118-3:2004 shall be used.
     * RIPEMD-160 | Specifies that the RIPEMD-160 algorithm, as defined by ISO/IEC 10118-3:2004 shall be used.
     * SHA-1 | Specifies that the SHA-1 algorithm, as defined by ISO/IEC 10118-3:2004 shall be used.
     * SHA-256 | Specifies that the SHA-256 algorithm, as defined by ISO/IEC 10118-3:2004 shall be used.
     * SHA-384 | Specifies that the SHA-384 algorithm, as defined by ISO/IEC 10118-3:2004  shall be used.
     * SHA-512 | Specifies that the SHA-512 algorithm, as defined by ISO/IEC 10118-3:2004 shall be used.
     * WHIRLPOOL | Specifies that the WHIRLPOOL algorithm, as defined by ISO/IEC 10118-3:2004 shall be used.
     */
    void setAlgorithmName(const QString &name);
    /**
     * @brief Returns the hash value for the password required to edit this worksheet.
     *
     * This value shall be compared with the resulting hash value after hashing
     * the user-supplied password using the #algorithmName, and if the two values
     * match, the protection shall no longer be enforced.
     *
     * This is an optional parameter. If hashValue is empty, it will not be stored in the doc.
     * @return a base64 string.
     */
    QString hashValue() const;
    /**
     * @brief sets the hash value for the password required to edit this worksheet.
     *
     * This value shall be compared with the resulting hash value after hashing
     * the user-supplied password using the #algorithmName(), and if the two values
     * match, the protection shall no longer be enforced.
     *
     * This is an optional parameter. If @a hashValue is empty, it will not be stored in the doc.
     * @param hashValue a base64 string.
     */
    void setHashValue(const QString hashValue);
    /**
     * @brief Returns the salt which was prepended to the user-supplied
     * password before it was hashed using #algorithmName() to generate #hashValue(),
     * and which shall also be prepended to the user-supplied password before
     * attempting to generate a hash value for comparison.
     *
     * A salt is a random string which is added to a user-supplied password
     * before it is hashed in order to prevent a malicious party from
     * pre-calculating all possible password/hash combinations and simply using
     * those pre-calculated values (often referred to as a "dictionary attack").
     *
     * @return a base64 string.
     *
     * This is an optional parameter. If saltValue is empty, it will not be stored in the doc.
     */
    QString saltValue() const;
    /**
     * @brief sets the salt which was prepended to the user-supplied
     * password before it was hashed using #algorithmName() to generate #hashValue(),
     * and which shall also be prepended to the user-supplied password before
     * attempting to generate a hash value for comparison.
     *
     * A salt is a random string which is added to a user-supplied password
     * before it is hashed in order to prevent a malicious party from
     * pre-calculating all possible password/hash combinations and simply using
     * those pre-calculated values (often referred to as a "dictionary attack").
     * @param saltValue a base64 string.
     *
     * This is an optional parameter. If saltValue is empty, it will not be stored in the doc.
     */
    void setSaltValue(const QString &saltValue);
    /**
     * @brief Returns the number of times the hashing function shall be
     * iteratively run (runs using each iteration's result plus a 4 byte value
     * (0-based, little endian) containing the number of the iteration as the
     * input for the next iteration) when attempting to compare a user-supplied
     * password with the value stored in #hashValue().
     * @return valid int value if the parameter was set, nullopt otherwise.
     */
    std::optional<int> spinCount() const;
    /**
     * @brief sets the number of times the hashing function shall be
     * iteratively run (runs using each iteration's result plus a 4 byte value
     * (0-based, little endian) containing the number of the iteration as the
     * input for the next iteration) when attempting to compare a user-supplied
     * password with the value stored in #hashValue().
     * @param spinCount a positive integer value.
     */
    void setSpinCount(int spinCount);
    /**
     * @brief returns whether editing of sheet content should not be allowed when
     * the chartsheet is protected.
     *
     * The default value is false.
     * @note This parameter is only applicable to chartsheets.
     * @return Valid bool value if the parameter is set, nullopt otherwise.
     */
    std::optional<bool> protectContent() const;
    /**
     * @brief sets whether editing of sheet content should not be allowed when
     * the chartsheet is protected.
     *
     * If not set, false is assumed.
     * @note This parameter is only applicable to chartsheets.
     * @param protectContent If true, then editing of sheet content is blocked.
     */
    void setProtectContent(bool protectContent);
    /**
     * @brief returns whether editing of objects should not be allowed when the sheet is protected.
     *
     * The default value is false.
     * @return Valid bool value if the parameter is set, nullopt otherwise.
     */
    std::optional<bool> protectObjects() const;
    /**
     * @brief sets whether editing of objects should not be allowed when the sheet is protected.
     * @param protectObjects If true, editing of objects is blocked.
     *
     * If not set, false is assumed.
     */
    void setProtectObjects(bool protectObjects);
    /**
     * @brief returns whether the other attributes of SheetProtection should be applied.
     *
     * If true then the other parameters of SheetProtection should be applied.
     * If false then the other parameters of SheetProtection should not be applied.
     *
     * If not set, the default value is false.
     * @note This parameter is not applicable to chartsheets.
     * @return Valid bool value if the parameter is set, nullopt otherwise.
     */
    std::optional<bool> protectSheet() const;
    /**
     * @brief sets whether the other attributes of SheetProtection should be applied.
     * @param protectSheet If true then the other parameters of SheetProtection should be applied.
     * If false then the other parameters of SheetProtection should not be applied.
     *
     * If not set, the default value is false.
     * @note This parameter is not applicable to chartsheets.
     */
    void setProtectSheet(bool protectSheet);
    /**
     * @brief returns whether editing of scenarios should not be allowed when
     * the sheet is protected.
     *
     * If not set, the default value is false.
     * @note This parameter is not applicable to chartsheets.
     */
    std::optional<bool> protectScenarios() const;
    /**
     * @brief sets whether editing of scenarios should not be allowed when
     * the sheet is protected.
     * @param protectScenarios If true, editing of scenarios is blocked.
     * @note This parameter is not applicable to chartsheets.
     */
    void setProtectScenarios(bool protectScenarios);

    /**
     * @brief Returns whether editing of cells formatting should not be
     * allowed when the sheet is protected.
     *
     * If not set, the default value is true.
     * @note This parameter is not applicable to chartsheets.
     */
    std::optional<bool> protectFormatCells() const;
    /**
     * @brief sets whether editing of cells formatting should not be
     * allowed when the sheet is protected.
     *
     * If not set, the default value is true.
     * @note This parameter is not applicable to chartsheets.
     * @param protectFormatCells If true, editing of cells formatting is blocked.
     */
    void setProtectFormatCells(bool protectFormatCells);
    /**
     * @brief returns whether editing of columns formatting should not be
     * allowed when the sheet is protected.
     *
     * If not set, the default value is true.
     * @note This parameter is not applicable to chartsheets.
     */
    std::optional<bool> protectFormatColumns() const;
    /**
     * @brief sets whether editing of columns formatting should not be
     * allowed when the sheet is protected.
     *
     * If not set, the default value is true.
     * @note This parameter is not applicable to chartsheets.
     * @param protectFormatColumns If true, then editing of columns formatting is blocked.
     */
    void setProtectFormatColumns(bool protectFormatColumns);
    /**
     * @brief returns whether editing of rows formatting should not be allowed
     * when the sheet is protected.
     *
     * If not set, the default value is true.
     * @note This parameter is not applicable to chartsheets.
     */
    std::optional<bool> protectFormatRows() const;
    /**
     * @brief sets whether editing of rows formatting should not be allowed
     * when the sheet is protected.
     *
     * If not set, the default value is true.
     * @note This parameter is not applicable to chartsheets.
     * @param protectFormatRows If true, then editing of rows formatting is blocked.
     */
    void setProtectFormatRows(bool protectFormatRows);
    /**
     * @brief returns whether inserting of columns should not be allowed when
     * the sheet is protected.
     *
     * If not set, the default value is true.
     * @note This parameter is not applicable to chartsheets.
     */
    std::optional<bool> protectInsertColumns() const;
    /**
     * @brief sets whether inserting of columns should not be allowed when
     * the sheet is protected.
     *
     * If not set, the default value is true.
     * @note This parameter is not applicable to chartsheets.
     * @param protectInsertColumns If true, then inserting of columns is blocked.
     */
    void setProtectInsertColumns(bool protectInsertColumns);
    /**
     * @brief returns whether inserting of rows should not be allowed when the
     * sheet is protected.
     *
     * If not set, the default value is true.
     * @note This parameter is not applicable to chartsheets.
     */
    std::optional<bool> protectInsertRows() const;
    /**
     * @brief sets whether inserting of rows should not be allowed when the
     * sheet is protected.
     *
     * If not set, the default value is true.
     * @note This parameter is not applicable to chartsheets.
     * @param protectInsertRows If true, then inserting of rows is not allowed.
     */
    void setProtectInsertRows(bool protectInsertRows);
    /**
     * @brief returns whether inserting of hyperlinks should not be allowed
     * when the sheet is protected.
     *
     * If not set, the default value is true.
     * @note This parameter is not applicable to chartsheets.
     */
    std::optional<bool> protectInsertHyperlinks() const;
    /**
     * @brief sets whether inserting of hyperlinks should not be allowed
     * when the sheet is protected.
     *
     * If not set, the default value is true.
     * @note This parameter is not applicable to chartsheets.
     * @param protectInsertHyperlinks If true, then inserting of hyperlinks is not allowed.
     */
    void setProtectInsertHyperlinks(bool protectInsertHyperlinks);
    /**
     * @brief returns whether deleting of columns should not be allowed when
     * the sheet is protected.
     *
     * If not set, the default value is true.
     * @note This parameter is not applicable to chartsheets.
     */
    std::optional<bool> protectDeleteColumns() const;
    /**
     * @brief sets whether deleting of columns should not be allowed when
     * the sheet is protected.
     *
     * If not set, the default value is true.
     * @note This parameter is not applicable to chartsheets.
     * @param protectDeleteColumns If true, then deleting of columns is not allowed.
     */
    void setProtectDeleteColumns(bool protectDeleteColumns);
    /**
     * @brief returns whether deleting of rows should not be allowed when the
     * sheet is protected.
     *
     * If not set, the default value is true.
     * @note This parameter is not applicable to chartsheets.
     */
    std::optional<bool> protectDeleteRows() const;
    /**
     * @brief sets whether deleting of rows should not be allowed when the
     * sheet is protected.
     *
     * If not set, the default value is true.
     * @note This parameter is not applicable to chartsheets.
     * @param protectDeleteRows If true, then deleting of rows is not allowed.
     */
    void setProtectDeleteRows(bool protectDeleteRows);
    /**
     * @brief returns whether selection of locked cells should not be allowed
     * when the sheet is protected.
     *
     * If not set, the default value is false.
     * @note This parameter is not applicable to chartsheets.
     */
    std::optional<bool> protectSelectLockedCells() const;
    /**
     * @brief sets whether selection of locked cells should not be allowed
     * when the sheet is protected.
     *
     * If not set, the default value is false.
     * @note This parameter is not applicable to chartsheets.
     * @param protectSelectLockedCells If true, then selection of locked cells is not allowed.
     */
    void setProtectSelectLockedCells(bool protectSelectLockedCells);
    /**
     * @brief returns whether sorting should not be allowed when the sheet is protected.
     *
     * If not set, the default value is true.
     * @note This parameter is not applicable to chartsheets.
     */
    std::optional<bool> protectSort() const;
    /**
     * @brief sets whether sorting should not be allowed when the sheet is protected.
     *
     * If not set, the default value is true.
     * @note This parameter is not applicable to chartsheets.
     * @param protectSort If true, then sorting is not allowed.
     */
    void setProtectSort(bool protectSort);
    /**
     * @brief returns whether applying autofilters should not be allowed when
     * the sheet is protected.
     *
     * If not set, the default value is true.
     * @note This parameter is not applicable to chartsheets.
     */
    std::optional<bool> protectAutoFilter() const;
    /**
     * @brief sets whether applying autofilters should not be allowed when
     * the sheet is protected.
     *
     * If not set, the default value is true.
     * @note This parameter is not applicable to chartsheets.
     * @param protectAutoFilter If true, then applying autofilter is blocked.
     */
    void setProtectAutoFilter(bool protectAutoFilter);
    /**
     * @brief returns whether operating pivot tables should not be allowed
     * when the sheet is protected.
     *
     * If not set, the default value is true.
     * @note This parameter is not applicable to chartsheets.
     */
    std::optional<bool> protectPivotTables() const;
    /**
     * @brief sets whether operating pivot tables should not be allowed
     * when the sheet is protected.
     *
     * If not set, the default value is true.
     * @note This parameter is not applicable to chartsheets.
     * @param protectPivotTables If true, then operating pivot tables is not allowed.
     */
    void setProtectPivotTables(bool protectPivotTables);
    /**
     * @brief returns whether selection of unlocked cells should not be
     * allowed when the sheet is protected.
     *
     * If not set, the default value is false.
     * @note This parameter is not applicable to chartsheets.
     */
    std::optional<bool> protectSelectUnlockedCells() const;
    /**
     * @brief sets whether selection of unlocked cells should not be
     * allowed when the sheet is protected.
     *
     * If not set, the default value is false.
     * @note This parameter is not applicable to chartsheets.
     * @param protectSelectUnlockedCells If true, then selection of unlocked is not allowed.
     */
    void setProtectSelectUnlockedCells(bool protectSelectUnlockedCells);

    /**
     * @brief returns whether any of protection parameters are set.
     * @return
     */
    bool isValid() const;
    void write(QXmlStreamWriter &writer, bool chartsheet = false) const;
    void read(QXmlStreamReader &reader);
#ifndef QT_NO_DEBUG_STREAM
    friend QDebug operator<<(QDebug, const SheetProtection &f);
#endif
    QSharedDataPointer<SheetProtectionPrivate> d;
};

#ifndef QT_NO_DEBUG_STREAM
QDebug operator<<(QDebug dbg, const SheetProtection &f);
#endif

}

Q_DECLARE_METATYPE(QXlsx::SheetProtection)

#endif // XLSXSHEETPROTECTION_H
