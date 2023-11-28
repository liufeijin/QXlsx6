#ifndef XLSXSHEETPROTECTION_H
#define XLSXSHEETPROTECTION_H

#include "xlsxglobal.h"

#include <QXmlStreamWriter>
#include <QSharedData>

namespace QXlsx {

/**
 * @brief The PasswordProtection struct represents the protection parameters
 * common for different classes: SheetProtection, WorkbookProtection etc. that
 * specify the password protection options (hash, salt, algorithm and spin
 * count).
 *
 * A static methods #randomizedSalt() and #setRandomizedSalt() allow to specify
 * the auto-generation of salt value. A 16-bit length random QByteArray will be
 * generated and used as a salt value if the salt in #hash() method is empty and
 * `setRandomizedSalt(true)` was invoked.
 */
struct QXLSX_EXPORT PasswordProtection
{
    /**
     * @brief creates an invalid PasswordProtection object.
     */
    PasswordProtection() {}
    /**
     * @brief creates a PasswordProtection object with specified password
     * parameters.
     */
    PasswordProtection(const QString &password, const QString &algorithm,
               const QString &salt, std::optional<int> spinCount);
    /**
     * @brief specifies the cryptographic hashing algorithm which shall
     * be used along with the salt attribute and input password in order to
     * compute the hash value.
     *
     * This is an optional parameter.
     *
     * The following names are reserved:
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
     *
     * The #hash() method supports only MD4, MD5, SHA-1, SHA-256, SHA-384,
     * SHA-512 algorithms.
     */
    QString algorithmName;
    /**
     * @brief specifies the hash value as a base64 string for the password
     * required to edit the document.
     *
     * This value shall be compared with the resulting hash value after hashing
     * the user-supplied password using the #algorithmName, and if the two
     * values match, the protection shall no longer be enforced.
     *
     * This is an optional parameter.
     */
    QString hashValue;
    /**
     * @brief specifies the salt as a base64 string which was prepended to the
     * user-supplied password before it was hashed using #algorithmName to
     * generate #hashValue, and which shall also be prepended to the
     * user-supplied password before attempting to generate a hash value for
     * comparison.
     *
     * A salt is a random string which is added to a user-supplied password
     * before it is hashed in order to prevent a malicious party from
     * pre-calculating all possible password/hash combinations and simply using
     * those pre-calculated values (often referred to as a "dictionary attack").
     *
     * This is an optional parameter.
     */
    QString saltValue;
    /**
     * @brief specifies the number of times the hashing function shall be
     * iteratively run (runs using each iteration's result plus a 4 byte value
     * (0-based, little endian) containing the number of the iteration as the
     * input for the next iteration) when attempting to compare a user-supplied
     * password with the value stored in #hashValue.
     *
     * This is an optional parameter.
     */
    std::optional<int> spinCount;
    bool isValid() const;
    bool operator==(const PasswordProtection &other) const;
    bool operator!=(const PasswordProtection &other) const;

    /**
     * @brief checks the user-specified @a password against the stored hashed
     * password and returns the result.
     * @param password A string that contains password.
     * @return `true` if @a passport generates the same hash as #hashValue,
     * `false` if the password protection is not set or @a password is invalid.
     */
    bool checkPassword(const QString &password) const;

    /**
     * @brief returns the base64-encoded hash value generated for @a password
     * with @a salt and @a spinCount using @a algorithm.
     *
     * The procedure is as follows:
     *
     * 1. Concatenate salt and password as a UTF-16 string.
     * 2. Compute hash of the result.
     * 3. If @a spinCount <=1, return the base64-encoded result.
     * 3. Otherwise append a 4-byte little-endian number of iterations so far
     *    (starting from 0).
     * 4. Compute hash of the result.
     * 5. Go to 3 until @a spinCount is exhausted.
     * 6. Return the result.
     *
     * @param algorithm The algorithm name. See #algorithmName for a list of
     * reserved names. See QCryptographicHash::Algorithm docs on the list of
     * supported algorithms.
     * @param password A string containing the password.
     * @param salt A string containing the salt. It can be any string of any
     * length.
     * @param spinCount A positive integer that specifies how many times to
     * compute the hash. A value of 100,000 or 1000,000 can be used.
     * @return The base64-encoded QByteArray value that can be directly stored
     * in the #hashValue.
     */
    static QByteArray hash(const QString &algorithm, const QString &password,
                           const QByteArray &salt, int spinCount);
    /**
     * @brief returns whether the salt will be auto-generated for the #hash()
     * method.
     * @return `true` if random salt value will be used instead of the empty
     * salt attribute of a #hash() method.
     *
     * The default value is `false`.
     */
    static bool randomizedSalt();
    /**
     * @brief sets whether the salt will be auto-generated for the #hash()
     * method.
     * @param use If `true` then random salt value will be used instead of the
     * empty salt attribute of the #hash() method. If `false` then the specified
     * salt value will be used in the #hash() method even if it's empty.
     *
     * The default value is `false`.
     */
    static void setRandomizedSalt(bool use);
    /**
     * @brief returns one of the QCryptographicHash::Algorithm enum values
     * @param algorithmName The algorithm name.
     * @return An QCryptographicHash::Algorithm enum value or -1 if
     * @a algorithmName is unsupported.
     */
    static int algorithmForName(const QString &algorithmName);
private:
    static bool randomized;
};

class SheetProtectionPrivate;

/**
 * @brief The SheetProtection class specifies the sheet protection parameters
 * for worksheets and chartsheets.
 *
 * Not all parameters are applicable to chartsheets. See the description of
 * parameters. If a non-applicable parameter is set for a chartsheet protection,
 * it will be ignored when saving the document.
 *
 * The class is _implicitly shared_: the deep copy occurs only in the non-const
 * methods.
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
     * @brief returns the sheet password protection parameters.
     * @return Protection object. May be invalid.
     */
    PasswordProtection protection() const;
    /**
     * @brief returns the sheet password protection parameters.
     * @return a reference to the Protection object.
     */
    PasswordProtection &protection();
    /**
     * @brief sets the sheet password protection parameters.
     * @param protection Protection object. If invalid, password protection is
     * unset.
     */
    void setProtection(const PasswordProtection &protection);


    /**
     * @brief returns whether editing of chartsheet content should not be
     * allowed when the chartsheet is protected.
     *
     * The default value is `false`.
     * @note This parameter is only applicable to chartsheets.
     * @return Valid bool value if the parameter is set, `nullopt` otherwise.
     */
    std::optional<bool> protectContent() const;
    /**
     * @brief sets whether editing of chartsheet content should not be allowed
     * when the chartsheet is protected.
     *
     * If not set, `false` is assumed.
     * @note This parameter is only applicable to chartsheets.
     * @param protectContent If `true`, then editing of sheet content is
     * blocked.
     */
    void setProtectContent(bool protectContent);
    /**
     * @brief returns whether editing of objects should not be allowed when the
     * sheet is protected.
     *
     * The default value is `false`.
     * @return Valid bool value if the parameter is set, `nullopt` otherwise.
     */
    std::optional<bool> protectObjects() const;
    /**
     * @brief sets whether editing of objects should not be allowed when the
     * sheet is protected.
     * @param protectObjects if `true`, editing of objects is blocked.
     *
     * If not set, `false` is assumed.
     */
    void setProtectObjects(bool protectObjects);
    /**
     * @brief returns whether the other attributes of SheetProtection should be
     * applied.
     *
     * If not set, the default value is `false`.
     * @note This parameter is not applicable to chartsheets.
     * @return Valid bool value if the parameter is set, `nullopt` otherwise.
     */
    std::optional<bool> protectSheet() const;
    /**
     * @brief sets whether the other attributes of SheetProtection should be
     * applied.
     * @param protectSheet If `true` then the other parameters of
     * SheetProtection should be applied.
     * If `false` then the other parameters of SheetProtection should not be
     * applied.
     *
     * If not set, the default value is `false`.
     * @note This parameter is not applicable to chartsheets.
     */
    void setProtectSheet(bool protectSheet);
    /**
     * @brief returns whether editing of scenarios should not be allowed when
     * the sheet is protected.
     *
     * If not set, the default value is `false`.
     * @note This parameter is not applicable to chartsheets.
     */
    std::optional<bool> protectScenarios() const;
    /**
     * @brief sets whether editing of scenarios should not be allowed when
     * the sheet is protected.
     * @param protectScenarios If `true`, editing of scenarios is blocked.
     * @note This parameter is not applicable to chartsheets.
     */
    void setProtectScenarios(bool protectScenarios);

    /**
     * @brief Returns whether editing of cells formatting should not be
     * allowed when the sheet is protected.
     *
     * If not set, the default value is `true`.
     * @note This parameter is not applicable to chartsheets.
     */
    std::optional<bool> protectFormatCells() const;
    /**
     * @brief sets whether editing of cells formatting should not be
     * allowed when the sheet is protected.
     *
     * If not set, the default value is `true`.
     * @note This parameter is not applicable to chartsheets.
     * @param protectFormatCells If `true`, editing of cells formatting is
     * blocked.
     */
    void setProtectFormatCells(bool protectFormatCells);
    /**
     * @brief returns whether editing of columns formatting should not be
     * allowed when the sheet is protected.
     *
     * If not set, the default value is `true`.
     * @note This parameter is not applicable to chartsheets.
     */
    std::optional<bool> protectFormatColumns() const;
    /**
     * @brief sets whether editing of columns formatting should not be
     * allowed when the sheet is protected.
     *
     * If not set, the default value is `true`.
     * @note This parameter is not applicable to chartsheets.
     * @param protectFormatColumns If `true`, then editing of columns formatting
     * is blocked.
     */
    void setProtectFormatColumns(bool protectFormatColumns);
    /**
     * @brief returns whether editing of rows formatting should not be allowed
     * when the sheet is protected.
     *
     * If not set, the default value is `true`.
     * @note This parameter is not applicable to chartsheets.
     */
    std::optional<bool> protectFormatRows() const;
    /**
     * @brief sets whether editing of rows formatting should not be allowed
     * when the sheet is protected.
     *
     * If not set, the default value is `true`.
     * @note This parameter is not applicable to chartsheets.
     * @param protectFormatRows If `true`, then editing of rows formatting is
     * blocked.
     */
    void setProtectFormatRows(bool protectFormatRows);
    /**
     * @brief returns whether inserting of columns should not be allowed when
     * the sheet is protected.
     *
     * If not set, the default value is `true`.
     * @note This parameter is not applicable to chartsheets.
     */
    std::optional<bool> protectInsertColumns() const;
    /**
     * @brief sets whether inserting of columns should not be allowed when
     * the sheet is protected.
     *
     * If not set, the default value is `true`.
     * @note This parameter is not applicable to chartsheets.
     * @param protectInsertColumns If `true`, then inserting of columns is
     * blocked.
     */
    void setProtectInsertColumns(bool protectInsertColumns);
    /**
     * @brief returns whether inserting of rows should not be allowed when the
     * sheet is protected.
     *
     * If not set, the default value is `true`.
     * @note This parameter is not applicable to chartsheets.
     */
    std::optional<bool> protectInsertRows() const;
    /**
     * @brief sets whether inserting of rows should not be allowed when the
     * sheet is protected.
     *
     * If not set, the default value is `true`.
     * @note This parameter is not applicable to chartsheets.
     * @param protectInsertRows If `true`, then inserting of rows is not
     * allowed.
     */
    void setProtectInsertRows(bool protectInsertRows);
    /**
     * @brief returns whether inserting of hyperlinks should not be allowed
     * when the sheet is protected.
     *
     * If not set, the default value is `true`.
     * @note This parameter is not applicable to chartsheets.
     */
    std::optional<bool> protectInsertHyperlinks() const;
    /**
     * @brief sets whether inserting of hyperlinks should not be allowed
     * when the sheet is protected.
     *
     * If not set, the default value is `true`.
     * @note This parameter is not applicable to chartsheets.
     * @param protectInsertHyperlinks If `true`, then inserting of hyperlinks is
     * not allowed.
     */
    void setProtectInsertHyperlinks(bool protectInsertHyperlinks);
    /**
     * @brief returns whether deleting of columns should not be allowed when
     * the sheet is protected.
     *
     * If not set, the default value is `true`.
     * @note This parameter is not applicable to chartsheets.
     */
    std::optional<bool> protectDeleteColumns() const;
    /**
     * @brief sets whether deleting of columns should not be allowed when
     * the sheet is protected.
     *
     * If not set, the default value is `true`.
     * @note This parameter is not applicable to chartsheets.
     * @param protectDeleteColumns If `true`, then deleting of columns is not
     * allowed.
     */
    void setProtectDeleteColumns(bool protectDeleteColumns);
    /**
     * @brief returns whether deleting of rows should not be allowed when the
     * sheet is protected.
     *
     * If not set, the default value is `true`.
     * @note This parameter is not applicable to chartsheets.
     */
    std::optional<bool> protectDeleteRows() const;
    /**
     * @brief sets whether deleting of rows should not be allowed when the
     * sheet is protected.
     *
     * If not set, the default value is `true`.
     * @note This parameter is not applicable to chartsheets.
     * @param protectDeleteRows If `true`, then deleting of rows is not allowed.
     */
    void setProtectDeleteRows(bool protectDeleteRows);
    /**
     * @brief returns whether selection of locked cells should not be allowed
     * when the sheet is protected.
     *
     * If not set, the default value is `false`.
     * @note This parameter is not applicable to chartsheets.
     */
    std::optional<bool> protectSelectLockedCells() const;
    /**
     * @brief sets whether selection of locked cells should not be allowed
     * when the sheet is protected.
     *
     * If not set, the default value is `false`.
     * @note This parameter is not applicable to chartsheets.
     * @param protectSelectLockedCells If `true`, then selection of locked cells
     * is not allowed.
     */
    void setProtectSelectLockedCells(bool protectSelectLockedCells);
    /**
     * @brief returns whether sorting should not be allowed when the sheet is
     * protected.
     *
     * If not set, the default value is `true`.
     * @note This parameter is not applicable to chartsheets.
     */
    std::optional<bool> protectSort() const;
    /**
     * @brief sets whether sorting should not be allowed when the sheet is
     * protected.
     *
     * If not set, the default value is `true`.
     * @note This parameter is not applicable to chartsheets.
     * @param protectSort If `true`, then sorting is not allowed.
     */
    void setProtectSort(bool protectSort);
    /**
     * @brief returns whether applying autofilters should not be allowed when
     * the sheet is protected.
     *
     * If not set, the default value is `true`.
     * @note This parameter is not applicable to chartsheets.
     */
    std::optional<bool> protectAutoFilter() const;
    /**
     * @brief sets whether applying autofilters should not be allowed when
     * the sheet is protected.
     *
     * If not set, the default value is `true`.
     * @note This parameter is not applicable to chartsheets.
     * @param protectAutoFilter If `true`, then applying autofilter is blocked.
     */
    void setProtectAutoFilter(bool protectAutoFilter);
    /**
     * @brief returns whether operating pivot tables should not be allowed
     * when the sheet is protected.
     *
     * If not set, the default value is `true`.
     * @note This parameter is not applicable to chartsheets.
     */
    std::optional<bool> protectPivotTables() const;
    /**
     * @brief sets whether operating pivot tables should not be allowed
     * when the sheet is protected.
     *
     * If not set, the default value is `true`.
     * @note This parameter is not applicable to chartsheets.
     * @param protectPivotTables If `true`, then operating pivot tables is not
     * allowed.
     */
    void setProtectPivotTables(bool protectPivotTables);
    /**
     * @brief returns whether selection of unlocked cells should not be
     * allowed when the sheet is protected.
     *
     * If not set, the default value is `false`.
     * @note This parameter is not applicable to chartsheets.
     */
    std::optional<bool> protectSelectUnlockedCells() const;
    /**
     * @brief sets whether selection of unlocked cells should not be
     * allowed when the sheet is protected.
     *
     * If not set, the default value is `false`.
     * @note This parameter is not applicable to chartsheets.
     * @param protectSelectUnlockedCells If `true`, then selection of unlocked
     * is not allowed.
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

class WorkbookProtectionPrivate;

/**
 * @brief The WorkbookProtection class represents the workbook protection
 * parameters.
 */
class QXLSX_EXPORT WorkbookProtection
{
public:
    WorkbookProtection();
    WorkbookProtection(const WorkbookProtection &other);
    ~WorkbookProtection();
    WorkbookProtection &operator=(const WorkbookProtection &other);

    bool operator == (const WorkbookProtection &other) const;
    bool operator != (const WorkbookProtection &other) const;

    operator QVariant() const;

    /**
     * @brief returns the workbook password protection parameters.
     * @return Protection object. May be invalid.
     */
    PasswordProtection protection() const;
    /**
     * @brief returns the workbook password protection parameters.
     * @return a reference to the Protection object.
     */
    PasswordProtection &protection();
    /**
     * @brief sets the workbook password protection parameters.
     * @param protection Protection object. If invalid, password protection is
     * unset.
     */
    void setProtection(const PasswordProtection &protection);

    /**
     * @brief returns the revisions password protection parameters.
     * @return Protection object. May be invalid.
     */
    PasswordProtection revisionsProtection() const;
    /**
     * @brief returns the revisions password protection parameters.
     * @return a reference to the Protection object.
     */
    PasswordProtection &revisionsProtection();
    /**
     * @brief sets the revisions password protection parameters.
     * @param protection Protection object. If invalid, password protection is
     * unset.
     */
    void setRevisionsProtection(const PasswordProtection &protection);

    /**
     * @brief returns whether the workbook structure is locked.
     *
     * If the workbook structure is locked, then sheets in the workbook can't be
     * moved, deleted, hidden, unhidden, or renamed, and new sheets can't be
     * inserted.
     *
     * The default value is `false`.
     */
    std::optional<bool> structureLocked() const;
    /**
     * @brief sets whether the workbook structure is locked.
     * @param locked If `true` then sheets in the workbook can't be
     * moved, deleted, hidden, unhidden, or renamed, and new sheets can't be
     * inserted.
     *
     * If not set, the default value is `false` (not locked).
     */
    void setStructureLocked(bool locked);
    /**
     * @brief returns whether the workbook windows are locked.
     *
     * If windows are locked, then they'll have the same size and position each
     * time the document is opened.
     *
     * The default value is `false` (not locked).
     */
    std::optional<bool> windowsLocked() const;
    /**
     * @brief sets whether the workbook windows are locked.
     * @param locked if `true`, then windows will have the same size and
     * position each time the document is opened.
     *
     * If not set, `false` (not locked) is assumed.
     */
    void setWindowsLocked(bool locked);
    /**
     * @brief returns whether the workbook is locked for revisions.
     */
    std::optional<bool> revisionsLocked() const;
    /**
     * @brief sets whether the workbook is locked for revisions.
     * @param locked if `true` then the workbook cannot get revisions.
     */
    void setRevisionsLocked(bool locked);
    /**
     * @brief returns whether any of protection parameters are set.
     */
    bool isValid() const;
    void write(QXmlStreamWriter &writer) const;
    void read(QXmlStreamReader &reader);

    QSharedDataPointer<WorkbookProtectionPrivate> d;
};

#ifndef QT_NO_DEBUG_STREAM
QDebug operator<<(QDebug dbg, const SheetProtection &f);
#endif

}

Q_DECLARE_METATYPE(QXlsx::SheetProtection)
Q_DECLARE_TYPEINFO(QXlsx::SheetProtection, Q_MOVABLE_TYPE);
Q_DECLARE_METATYPE(QXlsx::WorkbookProtection)
Q_DECLARE_TYPEINFO(QXlsx::WorkbookProtection, Q_MOVABLE_TYPE);

#endif // XLSXSHEETPROTECTION_H
