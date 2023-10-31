#include "xlsxsheetprotection.h"
#include "xlsxutility_p.h"

namespace QXlsx {

class SheetProtectionPrivate : public QSharedData
{
public:
    QString algorithmName;
    QString hashValue;
    QString saltValue;
    std::optional<int> spinCount;
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

QString SheetProtection::algorithmName() const
{
    if (!d) return {};
    return d->algorithmName;
}

void SheetProtection::setAlgorithmName(const QString &name)
{
    if (!d) d = new SheetProtectionPrivate;
    d->algorithmName = name;
}

QString SheetProtection::hashValue() const
{
    if (!d) return {};
    return d->hashValue;
}

void SheetProtection::setHashValue(const QString hashValue)
{
    if (!d) d = new SheetProtectionPrivate;
    d->hashValue = hashValue;
}

QString SheetProtection::saltValue() const
{
    if (!d) return {};
    return d->saltValue;
}

void SheetProtection::setSaltValue(const QString &saltValue)
{
    if (!d) d = new SheetProtectionPrivate;
    d->saltValue = saltValue;
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

std::optional<int> SheetProtection::spinCount() const
{
    if (!d) return {};
    return d->spinCount;
}

void SheetProtection::setSpinCount(int spinCount)
{
    if (!d) d = new SheetProtectionPrivate;
    d->spinCount = spinCount;
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
    writeAttribute(writer, QLatin1String("algorithmName"), d->algorithmName);
    writeAttribute(writer, QLatin1String("hashValue"), d->hashValue);
    writeAttribute(writer, QLatin1String("saltValue"), d->saltValue);
    writeAttribute(writer, QLatin1String("spinCount"), d->spinCount);
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
    parseAttributeString(a, QLatin1String("algorithmName"), d->algorithmName);
    parseAttributeString(a, QLatin1String("hashValue"), d->hashValue);
    parseAttributeString(a, QLatin1String("saltValue"), d->saltValue);
    parseAttributeInt(a, QLatin1String("spinCount"), d->spinCount);
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
    algorithmName{other.algorithmName},
    hashValue{other.hashValue},
    saltValue{other.saltValue},
    spinCount{other.spinCount},
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
    if (algorithmName != other.algorithmName) return false;
    if (hashValue != other.hashValue) return false;
    if (saltValue != other.saltValue) return false;
    if (spinCount != other.spinCount) return false;
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

}
