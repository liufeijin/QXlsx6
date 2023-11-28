#include "xlsxsheetprotection.h"
#include "xlsxutility_p.h"

#include <QCryptographicHash>
#include <QDebug>
#include <QBuffer>
#include <QDataStream>
#include <QFile>

#if QT_VERSION >= QT_VERSION_CHECK(5,10,0)
#include <QRandomGenerator>
#else
#include <random>
#endif

QByteArray getRandomSalt()
{
    QByteArray result;
    result.resize(16);

#if QT_VERSION >= QT_VERSION_CHECK(5,10,0)
    for (auto &b: result) {
        quint32 val = QRandomGenerator::global()->bounded(1, 255);
        b = (char)val;
    }
#else
    std::random_device r;
    std::default_random_engine engine(r());
    std::uniform_int_distribution<int> uniform_dist(1, 255);
    for (auto &b: result) {
        int val = uniform_dist(engine);
        b = (char)val;
    }
#endif
    return result;
}

namespace QXlsx {

bool PasswordProtection::randomized = false;

class SheetProtectionPrivate : public QSharedData
{
public:
    PasswordProtection protection;
    std::optional<bool> protectContent;
    std::optional<bool> protectObjects;
    std::optional<bool> protectSheet;
    std::optional<bool> protectScenarios;
    std::optional<bool> protectFormatCells;
    std::optional<bool> protectFormatColumns;
    std::optional<bool> protectFormatRows;
    std::optional<bool> protectInsertColumns;
    std::optional<bool> protectInsertRows;
    std::optional<bool> protectInsertHyperlinks;
    std::optional<bool> protectDeleteColumns;
    std::optional<bool> protectDeleteRows;
    std::optional<bool> protectSelectLockedCells;
    std::optional<bool> protectSort;
    std::optional<bool> protectAutoFilter;
    std::optional<bool> protectPivotTables;
    std::optional<bool> protectSelectUnlockedCells;

    SheetProtectionPrivate();
    SheetProtectionPrivate(const SheetProtectionPrivate &other);
    ~SheetProtectionPrivate();

    bool operator == (const SheetProtectionPrivate &other) const;
};

class WorkbookProtectionPrivate : public QSharedData
{
public:
    PasswordProtection protection;
    PasswordProtection revisionsProtection;
    std::optional<bool> lockStructure;
    std::optional<bool> lockWindows;
    std::optional<bool> lockRevision;

    WorkbookProtectionPrivate();
    WorkbookProtectionPrivate(const WorkbookProtectionPrivate &other);
    ~WorkbookProtectionPrivate();

    bool operator == (const WorkbookProtectionPrivate &other) const;
};

SheetProtection::SheetProtection()
{

}

SheetProtection::SheetProtection(const SheetProtection &other) : d{other.d}
{

}

SheetProtection::~SheetProtection()
{

}

SheetProtection &SheetProtection::operator=(const SheetProtection &other)
{
    if (*this != other) d = other.d;
    return *this;
}

bool SheetProtection::operator ==(const SheetProtection &other) const
{
    if (d == other.d) return true;
    if (!d || !other.d) return false;
    return *this->d.constData() == *other.d.constData();
}

bool SheetProtection::operator !=(const SheetProtection &other) const
{
    return !operator==(other);
}

SheetProtection::operator QVariant() const
{
    const auto& cref
#if QT_VERSION >= 0x060000 // Qt 6.0 or over
        = QMetaType::fromType<SheetProtection>();
#else
        = qMetaTypeId<SheetProtection>() ;
#endif
    return QVariant(cref, this);
}

bool SheetProtection::checkPassword(const QString &password) const
{
    if (!d) return false;
    if (!d->protection.isValid()) return false;
    auto salt = QByteArray::fromBase64(d->protection.saltValue.toLocal8Bit());
    auto hashed = PasswordProtection::hash(d->protection.algorithmName,
                                   password,
                                   salt,
                                   d->protection.spinCount.value_or(0));
    return hashed == d->protection.hashValue;
}

WorkbookProtection &WorkbookProtection::operator=(const WorkbookProtection &other)
{
    if (*this != other) d = other.d;
    return *this;
}

bool WorkbookProtection::operator == (const WorkbookProtection &other) const
{
    if (d == other.d) return true;
    if (!d || !other.d) return false;
    return *this->d.constData() == *other.d.constData();
}
bool WorkbookProtection::operator != (const WorkbookProtection &other) const
{
    return !operator==(other);
}

WorkbookProtection::operator QVariant() const
{
    const auto& cref
#if QT_VERSION >= 0x060000 // Qt 6.0 or over
        = QMetaType::fromType<WorkbookProtection>();
#else
        = qMetaTypeId<WorkbookProtection>() ;
#endif
    return QVariant(cref, this);
}

PasswordProtection SheetProtection::protection() const
{
    if (d) return d->protection;
    return {};
}

PasswordProtection &SheetProtection::protection()
{
    if (!d) d = new SheetProtectionPrivate;
    return d->protection;
}

void SheetProtection::setProtection(const PasswordProtection &protection)
{
    if (!d) d = new SheetProtectionPrivate;
    d->protection = protection;
}

std::optional<bool> SheetProtection::protectSelectUnlockedCells() const
{
    if (!d) return {};
    return d->protectSelectUnlockedCells;
}

void SheetProtection::setProtectSelectUnlockedCells(bool protectSelectUnlockedCells)
{
    if (!d) d = new SheetProtectionPrivate;
    d->protectSelectUnlockedCells = protectSelectUnlockedCells;
}

std::optional<bool> SheetProtection::protectPivotTables() const
{
    if (!d) return {};
    return d->protectPivotTables;
}

void SheetProtection::setProtectPivotTables(bool protectPivotTables)
{
    if (!d) d = new SheetProtectionPrivate;
    d->protectPivotTables = protectPivotTables;
}

std::optional<bool> SheetProtection::protectAutoFilter() const
{
    if (!d) return {};
    return d->protectAutoFilter;
}

void SheetProtection::setProtectAutoFilter(bool protectAutoFilter)
{
    if (!d) d = new SheetProtectionPrivate;
    d->protectAutoFilter = protectAutoFilter;
}

std::optional<bool> SheetProtection::protectSort() const
{
    if (!d) return {};
    return d->protectSort;
}

void SheetProtection::setProtectSort(bool protectSort)
{
    if (!d) d = new SheetProtectionPrivate;
    d->protectSort = protectSort;
}

std::optional<bool> SheetProtection::protectSelectLockedCells() const
{
    if (!d) return {};
    return d->protectSelectLockedCells;
}

void SheetProtection::setProtectSelectLockedCells(bool protectSelectLockedCells)
{
    if (!d) d = new SheetProtectionPrivate;
    d->protectSelectLockedCells = protectSelectLockedCells;
}

std::optional<bool> SheetProtection::protectDeleteRows() const
{
    if (!d) return {};
    return d->protectDeleteRows;
}

void SheetProtection::setProtectDeleteRows(bool protectDeleteRows)
{
    if (!d) d = new SheetProtectionPrivate;
    d->protectDeleteRows = protectDeleteRows;
}

std::optional<bool> SheetProtection::protectDeleteColumns() const
{
    if (!d) return {};
    return d->protectDeleteColumns;
}

void SheetProtection::setProtectDeleteColumns(bool protectDeleteColumns)
{
    if (!d) d = new SheetProtectionPrivate;
    d->protectDeleteColumns = protectDeleteColumns;
}

std::optional<bool> SheetProtection::protectInsertHyperlinks() const
{
    if (!d) return {};
    return d->protectInsertHyperlinks;
}

void SheetProtection::setProtectInsertHyperlinks(bool newProtectInsertHyperlinks)
{
    if (!d) d = new SheetProtectionPrivate;
    d->protectInsertHyperlinks = newProtectInsertHyperlinks;
}

std::optional<bool> SheetProtection::protectInsertRows() const
{
    if (!d) return {};
    return d->protectInsertRows;
}

void SheetProtection::setProtectInsertRows(bool newProtectInsertRows)
{
    if (!d) d = new SheetProtectionPrivate;
    d->protectInsertRows = newProtectInsertRows;
}

std::optional<bool> SheetProtection::protectInsertColumns() const
{
    if (!d) return {};
    return d->protectInsertColumns;
}

void SheetProtection::setProtectInsertColumns(bool newProtectInsertColumns)
{
    if (!d) d = new SheetProtectionPrivate;
    d->protectInsertColumns = newProtectInsertColumns;
}

std::optional<bool> SheetProtection::protectFormatRows() const
{
    if (!d) return {};
    return d->protectFormatRows;
}

void SheetProtection::setProtectFormatRows(bool newProtectFormatRows)
{
    if (!d) d = new SheetProtectionPrivate;
    d->protectFormatRows = newProtectFormatRows;
}

std::optional<bool> SheetProtection::protectFormatColumns() const
{
    if (!d) return {};
    return d->protectFormatColumns;
}

void SheetProtection::setProtectFormatColumns(bool newProtectFormatColumns)
{
    if (!d) d = new SheetProtectionPrivate;
    d->protectFormatColumns = newProtectFormatColumns;
}

std::optional<bool> SheetProtection::protectFormatCells() const
{
    if (!d) return {};
    return d->protectFormatCells;
}

void SheetProtection::setProtectFormatCells(bool newProtectFormatCells)
{
    if (!d) d = new SheetProtectionPrivate;
    d->protectFormatCells = newProtectFormatCells;
}

std::optional<bool> SheetProtection::protectScenarios() const
{
    if (!d) return {};
    return d->protectScenarios;
}

void SheetProtection::setProtectScenarios(bool newProtectScenarios)
{
    if (!d) d = new SheetProtectionPrivate;
    d->protectScenarios = newProtectScenarios;
}

std::optional<bool> SheetProtection::protectSheet() const
{
    if (!d) return {};
    return d->protectSheet;
}

void SheetProtection::setProtectSheet(bool newProtectSheet)
{
    if (!d) d = new SheetProtectionPrivate;
    d->protectSheet = newProtectSheet;
}

std::optional<bool> SheetProtection::protectObjects() const
{
    if (!d) return {};
    return d->protectObjects;
}

void SheetProtection::setProtectObjects(bool newProtectObjects)
{
    if (!d) d = new SheetProtectionPrivate;
    d->protectObjects = newProtectObjects;
}

std::optional<bool> SheetProtection::protectContent() const
{
    if (!d) return {};
    return d->protectContent;
}

void SheetProtection::setProtectContent(bool protectContent)
{
    if (!d) d = new SheetProtectionPrivate;
    d->protectContent = protectContent;
}

bool SheetProtection::isValid() const
{
    if (d)
        return true;
    return false;
}

void SheetProtection::write(QXmlStreamWriter &writer, bool chartsheet) const
{
    if (!d) return;
    writer.writeEmptyElement(QLatin1String("sheetProtection"));
    writeAttribute(writer, QLatin1String("algorithmName"), d->protection.algorithmName);
    writeAttribute(writer, QLatin1String("hashValue"), d->protection.hashValue);
    writeAttribute(writer, QLatin1String("saltValue"), d->protection.saltValue);
    writeAttribute(writer, QLatin1String("spinCount"), d->protection.spinCount);
    if (chartsheet) writeAttribute(writer, QLatin1String("content"), d->protectContent);
    if (!chartsheet) writeAttribute(writer, QLatin1String("sheet"), d->protectSheet);
    writeAttribute(writer, QLatin1String("objects"), d->protectObjects);
    if (!chartsheet) writeAttribute(writer, QLatin1String("scenarios"), d->protectScenarios);
    if (!chartsheet) writeAttribute(writer, QLatin1String("formatCells"), d->protectFormatCells);
    if (!chartsheet) writeAttribute(writer, QLatin1String("formatColumns"), d->protectFormatColumns);
    if (!chartsheet) writeAttribute(writer, QLatin1String("formatRows"), d->protectFormatRows);
    if (!chartsheet) writeAttribute(writer, QLatin1String("insertColumns"), d->protectInsertColumns);
    if (!chartsheet) writeAttribute(writer, QLatin1String("insertRows"), d->protectInsertRows);
    if (!chartsheet) writeAttribute(writer, QLatin1String("insertHyperlinks"), d->protectInsertHyperlinks);
    if (!chartsheet) writeAttribute(writer, QLatin1String("deleteColumns"), d->protectDeleteColumns);
    if (!chartsheet) writeAttribute(writer, QLatin1String("deleteRows"), d->protectDeleteRows);
    if (!chartsheet) writeAttribute(writer, QLatin1String("selectLockedCells"), d->protectSelectLockedCells);
    if (!chartsheet) writeAttribute(writer, QLatin1String("sort"), d->protectSort);
    if (!chartsheet) writeAttribute(writer, QLatin1String("autoFilter"), d->protectAutoFilter);
    if (!chartsheet) writeAttribute(writer, QLatin1String("pivotTables"), d->protectPivotTables);
    if (!chartsheet) writeAttribute(writer, QLatin1String("selectUnlockedCells"), d->protectSelectUnlockedCells);
}

void SheetProtection::read(QXmlStreamReader &reader)
{
    d = new SheetProtectionPrivate;
    const auto &a = reader.attributes();
    parseAttributeString(a, QLatin1String("algorithmName"), d->protection.algorithmName);
    parseAttributeString(a, QLatin1String("hashValue"), d->protection.hashValue);
    parseAttributeString(a, QLatin1String("saltValue"), d->protection.saltValue);
    parseAttributeInt(a, QLatin1String("spinCount"), d->protection.spinCount);
    parseAttributeBool(a, QLatin1String("content"), d->protectContent);
    parseAttributeBool(a, QLatin1String("sheet"), d->protectSheet);
    parseAttributeBool(a, QLatin1String("objects"), d->protectObjects);
    parseAttributeBool(a, QLatin1String("scenarios"), d->protectScenarios);
    parseAttributeBool(a, QLatin1String("formatCells"), d->protectFormatCells);
    parseAttributeBool(a, QLatin1String("formatColumns"), d->protectFormatColumns);
    parseAttributeBool(a, QLatin1String("formatRows"), d->protectFormatRows);
    parseAttributeBool(a, QLatin1String("insertColumns"), d->protectInsertColumns);
    parseAttributeBool(a, QLatin1String("insertRows"), d->protectInsertRows);
    parseAttributeBool(a, QLatin1String("insertHyperlinks"), d->protectInsertHyperlinks);
    parseAttributeBool(a, QLatin1String("deleteColumns"), d->protectDeleteColumns);
    parseAttributeBool(a, QLatin1String("deleteRows"), d->protectDeleteRows);
    parseAttributeBool(a, QLatin1String("selectLockedCells"), d->protectSelectLockedCells);
    parseAttributeBool(a, QLatin1String("sort"), d->protectSort);
    parseAttributeBool(a, QLatin1String("autoFilter"), d->protectAutoFilter);
    parseAttributeBool(a, QLatin1String("pivotTables"), d->protectPivotTables);
    parseAttributeBool(a, QLatin1String("selectUnlockedCells"), d->protectSelectUnlockedCells);
}

SheetProtectionPrivate::SheetProtectionPrivate()
{

}

SheetProtectionPrivate::SheetProtectionPrivate(const SheetProtectionPrivate &other)
    : QSharedData(other),
    protection{other.protection},
    protectContent{other.protectContent},
    protectObjects{other.protectObjects},
    protectSheet{other.protectSheet},
    protectScenarios{other.protectScenarios},
    protectFormatCells{other.protectFormatCells},
    protectFormatColumns{other.protectFormatColumns},
    protectFormatRows{other.protectFormatRows},
    protectInsertColumns{other.protectInsertColumns},
    protectInsertRows{other.protectInsertRows},
    protectInsertHyperlinks{other.protectInsertHyperlinks},
    protectDeleteColumns{other.protectDeleteColumns},
    protectDeleteRows{other.protectDeleteRows},
    protectSelectLockedCells{other.protectSelectLockedCells},
    protectSort{other.protectSort},
    protectAutoFilter{other.protectAutoFilter},
    protectPivotTables{other.protectPivotTables},
    protectSelectUnlockedCells{other.protectSelectUnlockedCells}
{

}

SheetProtectionPrivate::~SheetProtectionPrivate()
{

}

bool SheetProtectionPrivate::operator ==(const SheetProtectionPrivate &other) const
{
    if (protection != other.protection) return false;
    if (protectContent != other.protectContent) return false;
    if (protectObjects != other.protectObjects) return false;
    if (protectSheet != other.protectSheet) return false;
    if (protectScenarios != other.protectScenarios) return false;
    if (protectFormatCells != other.protectFormatCells) return false;
    if (protectFormatColumns != other.protectFormatColumns) return false;
    if (protectFormatRows != other.protectFormatRows) return false;
    if (protectInsertColumns != other.protectInsertColumns) return false;
    if (protectInsertRows != other.protectInsertRows) return false;
    if (protectInsertHyperlinks != other.protectInsertHyperlinks) return false;
    if (protectDeleteColumns != other.protectDeleteColumns) return false;
    if (protectDeleteRows != other.protectDeleteRows) return false;
    if (protectSelectLockedCells != other.protectSelectLockedCells) return false;
    if (protectSort != other.protectSort) return false;
    if (protectAutoFilter != other.protectAutoFilter) return false;
    if (protectPivotTables != other.protectPivotTables) return false;
    if (protectSelectUnlockedCells != other.protectSelectUnlockedCells) return false;
    return true;
}

PasswordProtection::PasswordProtection(const QString &password, const QString &algorithm, const QString &salt, std::optional<int> spinCount)
    : algorithmName(algorithm), spinCount(spinCount)
{
    if (!password.isEmpty()) {
        auto s = salt.isEmpty() && randomized ? getRandomSalt() : salt.toLocal8Bit();
        saltValue = s.toBase64();
        hashValue = hash(algorithm, password, s, spinCount.value_or(0));
    }
}

bool PasswordProtection::isValid() const
{
    return !algorithmName.isEmpty() || !hashValue.isEmpty() || !saltValue.isEmpty()
           || spinCount.has_value();
}

bool PasswordProtection::operator==(const PasswordProtection &other) const
{
    return algorithmName==other.algorithmName && hashValue == other.hashValue &&
           saltValue == other.saltValue && spinCount == other.spinCount;
}

bool PasswordProtection::operator!=(const PasswordProtection &other) const
{
    return !operator==(other);
}

QByteArray PasswordProtection::hash(const QString &algorithm, const QString &password, const QByteArray &salt, int spinCount)
{
    auto algo = algorithmForName(algorithm);
    if (algo == -1) {
        return {};
    }

    QByteArray salty;
    QBuffer b(&salty);
    b.open(QIODevice::WriteOnly);
    QDataStream s(&b);
    for (auto c: qAsConst(salt)) {
        s.writeRawData(&c, 1);
    }

    auto passwordUtf16 = password.toStdU16String();
    s.writeRawData(reinterpret_cast<char*>(const_cast<char16_t*>(passwordUtf16.c_str())), passwordUtf16.size()*2);
    b.close();

    QCryptographicHash hash(static_cast<QCryptographicHash::Algorithm>(algo));

    hash.addData(salty);
    QByteArray hashed = hash.result();
    for (int i=0; i<spinCount; ++i) {
        QByteArray bi;
        QBuffer b(&bi);
        b.open(QIODevice::WriteOnly);
        QDataStream s(&b);
        s.setByteOrder(QDataStream::LittleEndian);
        s << (uint32_t)i;
        b.close();
        hashed.append(bi);

        hash.reset();
        hash.addData(hashed);
        hashed = hash.result();
    }
    return hashed.toBase64();
}

bool PasswordProtection::randomizedSalt()
{
    return randomized;
}

void PasswordProtection::setRandomizedSalt(bool use)
{
    randomized = use;
}

int PasswordProtection::algorithmForName(const QString &algorithmName)
{
    if (algorithmName == "MD4") return QCryptographicHash::Md4;
    if (algorithmName == "MD5") return QCryptographicHash::Md5;
    if (algorithmName == "SHA-1") return QCryptographicHash::Sha1;
    if (algorithmName == "SHA-224") return QCryptographicHash::Sha224;
    if (algorithmName == "SHA-256") return QCryptographicHash::Sha256;
    if (algorithmName == "SHA-384") return QCryptographicHash::Sha384;
    if (algorithmName == "SHA-512") return QCryptographicHash::Sha512;

    return -1;
}

WorkbookProtectionPrivate::WorkbookProtectionPrivate()
{

}

WorkbookProtectionPrivate::WorkbookProtectionPrivate(const WorkbookProtectionPrivate &other)
    : QSharedData(other),
    protection{other.protection},
    revisionsProtection{other.revisionsProtection},
    lockStructure{other.lockStructure},
    lockWindows{other.lockWindows},
    lockRevision{other.lockRevision}
{

}

WorkbookProtectionPrivate::~WorkbookProtectionPrivate()
{

}

bool WorkbookProtectionPrivate::operator ==(const WorkbookProtectionPrivate &other) const
{
    if (protection != other.protection) return false;
    if (revisionsProtection != other.revisionsProtection) return false;
    if (lockStructure != other.lockStructure) return false;
    if (lockWindows != other.lockWindows) return false;
    if (lockRevision != other.lockRevision) return false;
    return true;
}

WorkbookProtection::WorkbookProtection()
{

}

WorkbookProtection::WorkbookProtection(const WorkbookProtection &other) : d{other.d}
{

}

WorkbookProtection::~WorkbookProtection()
{

}

PasswordProtection WorkbookProtection::protection() const
{
    if (d) return d->protection;
    return {};
}

PasswordProtection &WorkbookProtection::protection()
{
    if (!d) d = new WorkbookProtectionPrivate;
    return d->protection;
}

void WorkbookProtection::setProtection(const PasswordProtection &protection)
{
    if (!d) d = new WorkbookProtectionPrivate;
    d->protection = protection;
}

PasswordProtection WorkbookProtection::revisionsProtection() const
{
    if (d) return d->revisionsProtection;
    return {};
}

PasswordProtection &WorkbookProtection::revisionsProtection()
{
    if (!d) d = new WorkbookProtectionPrivate;
    return d->revisionsProtection;
}

void WorkbookProtection::setRevisionsProtection(const PasswordProtection &protection)
{
    if (!d) d = new WorkbookProtectionPrivate;
    d->revisionsProtection = protection;
}

std::optional<bool> WorkbookProtection::structureLocked() const
{
    if (d) return d->lockStructure;
    return {};
}

void WorkbookProtection::setStructureLocked(bool locked)
{
    if (!d) d = new WorkbookProtectionPrivate;
    d->lockStructure = locked;
}

std::optional<bool> WorkbookProtection::windowsLocked() const
{
    if (d) return d->lockWindows;
    return {};
}

void WorkbookProtection::setWindowsLocked(bool locked)
{
    if (!d) d = new WorkbookProtectionPrivate;
    d->lockWindows = locked;
}

std::optional<bool> WorkbookProtection::revisionsLocked() const
{
    if (d) return d->lockRevision;
    return {};
}

void WorkbookProtection::setRevisionsLocked(bool locked)
{
    if (!d) d = new WorkbookProtectionPrivate;
    d->lockRevision = locked;
}

bool WorkbookProtection::isValid() const
{
    if (d) return true;
    return false;
}

void WorkbookProtection::write(QXmlStreamWriter &writer) const
{
    if (!d) return;
    writer.writeEmptyElement(QLatin1String("workbookProtection"));
    writeAttribute(writer, QLatin1String("lockStructure"), d->lockStructure);
    writeAttribute(writer, QLatin1String("lockWindows"), d->lockWindows);
    writeAttribute(writer, QLatin1String("lockRevision"), d->lockRevision);
    writeAttribute(writer, QLatin1String("revisionsAlgorithmName"), d->revisionsProtection.algorithmName);
    writeAttribute(writer, QLatin1String("revisionsHashValue"), d->revisionsProtection.hashValue);
    writeAttribute(writer, QLatin1String("revisionsSaltValue"), d->revisionsProtection.saltValue);
    writeAttribute(writer, QLatin1String("revisionsSpinCount"), d->revisionsProtection.spinCount);
    writeAttribute(writer, QLatin1String("workbookAlgorithmName"), d->protection.algorithmName);
    writeAttribute(writer, QLatin1String("workbookHashValue"), d->protection.hashValue);
    writeAttribute(writer, QLatin1String("workbookSaltValue"), d->protection.saltValue);
    writeAttribute(writer, QLatin1String("workbookSpinCount"), d->protection.spinCount);
}

void WorkbookProtection::read(QXmlStreamReader &reader)
{
    d = new WorkbookProtectionPrivate;
    const auto &a = reader.attributes();
    parseAttributeString(a, QLatin1String("workbookAlgorithmName"), d->protection.algorithmName);
    parseAttributeString(a, QLatin1String("workbookHashValue"), d->protection.hashValue);
    parseAttributeString(a, QLatin1String("workbookSaltValue"), d->protection.saltValue);
    parseAttributeInt(a, QLatin1String("workbookSpinCount"), d->protection.spinCount);
    parseAttributeString(a, QLatin1String("revisionsAlgorithmName"), d->revisionsProtection.algorithmName);
    parseAttributeString(a, QLatin1String("revisionsHashValue"), d->revisionsProtection.hashValue);
    parseAttributeString(a, QLatin1String("revisionsSaltValue"), d->revisionsProtection.saltValue);
    parseAttributeInt(a, QLatin1String("revisionsSpinCount"), d->revisionsProtection.spinCount);
    parseAttributeBool(a, QLatin1String("lockRevision"), d->lockRevision);
    parseAttributeBool(a, QLatin1String("lockWindows"), d->lockWindows);
    parseAttributeBool(a, QLatin1String("lockStructure"), d->lockStructure);
}

}
