// xlsxworkbook.cpp

#include <QBuffer>
#include <QDir>
#include <QFile>
#include <QXmlStreamReader>
#include <QXmlStreamWriter>
#include <QtDebug>
#include <QtGlobal>

#include "xlsxchart.h"
#include "xlsxchartsheet.h"
#include "xlsxmediafile_p.h"
#include "xlsxsharedstrings_p.h"
#include "xlsxstyles_p.h"
#include "xlsxutility_p.h"
#include "xlsxworkbook.h"
#include "xlsxworkbook_p.h"
#include "xlsxworksheet.h"

namespace QXlsx {

WorkbookPrivate::WorkbookPrivate(Workbook *q, Workbook::CreateFlag flag)
    : AbstractOOXmlFilePrivate(q, flag)
{
    sharedStrings = QSharedPointer<SharedStrings>(new SharedStrings(flag));
    styles = QSharedPointer<Styles>(new Styles(flag));
    theme = QSharedPointer<Theme>(new Theme(flag));

    defaultDateFormat = QStringLiteral("yyyy-mm-dd");
    table_count = 0;
}

void WorkbookView::write(QXmlStreamWriter &writer) const
{
    if (extLst.isValid()) writer.writeStartElement(QLatin1String("workbookView"));
    else writer.writeEmptyElement(QLatin1String("workbookView"));
    if (visibility.has_value())
        writeAttribute(writer, QLatin1String("visibility"), AbstractSheet::toString(visibility.value()));
    writeAttribute(writer, QLatin1String("minimized"), minimized);
    writeAttribute(writer, QLatin1String("showHorizontalScroll"), showHorizontalScroll);
    writeAttribute(writer, QLatin1String("showVerticalScroll"), showVerticalScroll);
    writeAttribute(writer, QLatin1String("showSheetTabs"), showSheetTabs);
    writeAttribute(writer, QLatin1String("xWindow"), xWindow);
    writeAttribute(writer, QLatin1String("yWindow"), yWindow);
    writeAttribute(writer, QLatin1String("windowWidth"), windowWidth);
    writeAttribute(writer, QLatin1String("windowHeight"), windowHeight);
    writeAttribute(writer, QLatin1String("tabRatio"), tabRatio);
    writeAttribute(writer, QLatin1String("firstSheet"), firstSheet);
    writeAttribute(writer, QLatin1String("activeTab"), activeTab);
    writeAttribute(writer, QLatin1String("autoFilterDateGrouping"), autoFilterDateGrouping);
    extLst.write(writer, QLatin1String("extLst"));
    if (extLst.isValid()) writer.writeEndElement();
}

void WorkbookView::read(QXmlStreamReader &reader)
{
    const auto &a = reader.attributes();
    if (a.hasAttribute(QLatin1String("visibility"))) {
        AbstractSheet::Visibility v;
        AbstractSheet::fromString(a.value(QLatin1String("visibility")).toString(), v);
        visibility = v;
    }
    parseAttributeBool(a, QLatin1String("minimized"), minimized);
    parseAttributeBool(a, QLatin1String("showHorizontalScroll"), showHorizontalScroll);
    parseAttributeBool(a, QLatin1String("showVerticalScroll"), showVerticalScroll);
    parseAttributeBool(a, QLatin1String("showSheetTabs"), showSheetTabs);
    parseAttributeInt(a, QLatin1String("xWindow"), xWindow);
    parseAttributeInt(a, QLatin1String("yWindow"), yWindow);
    parseAttributeInt(a, QLatin1String("windowWidth"), windowWidth);
    parseAttributeInt(a, QLatin1String("windowHeight"), windowHeight);
    parseAttributeInt(a, QLatin1String("tabRatio"), tabRatio);
    parseAttributeInt(a, QLatin1String("firstSheet"), firstSheet);
    parseAttributeInt(a, QLatin1String("activeTab"), activeTab);
    parseAttributeBool(a, QLatin1String("autoFilterDateGrouping"), autoFilterDateGrouping);
    reader.readNextStartElement();
    if (reader.name() == QLatin1String("extLst")) extLst.read(reader);
}

Workbook::Workbook(CreateFlag flag)
    : AbstractOOXmlFile(new WorkbookPrivate(this, flag))
{}

Workbook::~Workbook() {}

std::optional<bool> Workbook::date1904() const
{
    Q_D(const Workbook);
    return d->date1904;
}

void Workbook::setDate1904(bool date1904)
{
    Q_D(Workbook);
    d->date1904 = date1904;
}

void Workbook::setStringsToNumbersEnabled(bool enable)
{
    Q_D(Workbook);
    d->strings_to_numbers_enabled = enable;
}

bool Workbook::isStringsToNumbersEnabled() const
{
    Q_D(const Workbook);
    return d->strings_to_numbers_enabled;
}

void Workbook::setStringsToHyperlinksEnabled(bool enable)
{
    Q_D(Workbook);
    d->strings_to_hyperlinks_enabled = enable;
}

bool Workbook::isStringsToHyperlinksEnabled() const
{
    Q_D(const Workbook);
    return d->strings_to_hyperlinks_enabled;
}

void Workbook::setHtmlToRichStringEnabled(bool enable)
{
    Q_D(Workbook);
    d->html_to_richstring_enabled = enable;
}

bool Workbook::isHtmlToRichStringEnabled() const
{
    Q_D(const Workbook);
    return d->html_to_richstring_enabled;
}

QString Workbook::defaultDateFormat() const
{
    Q_D(const Workbook);
    return d->defaultDateFormat;
}

void Workbook::setDefaultDateFormat(const QString &format)
{
    Q_D(Workbook);
    d->defaultDateFormat = format;
}

std::optional<Workbook::ReferenceMode> Workbook::referenceMode() const
{
    Q_D(const Workbook);
    return d->refMode;
}

void Workbook::setReferenceMode(ReferenceMode mode)
{
    Q_D(Workbook);
    d->refMode = mode;
}

std::optional<Workbook::CalculationMode> Workbook::calculationMode() const
{
    Q_D(const Workbook);
    return d->calcMode;
}

void Workbook::setCalculationMode(CalculationMode mode)
{
    Q_D(Workbook);
    d->calcMode = mode;
}

void Workbook::setRecalculationOnLoad(bool recalculate)
{
    Q_D(Workbook);
    d->fullCalcOnLoad = recalculate;
}

std::optional<bool> Workbook::fullPrecisionOnCalculation() const
{
    Q_D(const Workbook);
    return d->fullPrecision;
}

void Workbook::setFullPrecisionOnCalculation(bool fullPrecision)
{
    Q_D(Workbook);
    d->fullPrecision = fullPrecision;
}

void Workbook::setCalculationParametersDefaults()
{
    Q_D(Workbook);
    d->calcMode.reset(); //default="auto"
    d->fullCalcOnLoad.reset(); // default="false"/>
    d->refMode.reset();// default="A1"/>
    d->iterate.reset();// default="false"/>
    d->iterateCount.reset();// default="100"/>
    d->iterateDelta.reset();// default="0.001"/>
    d->fullPrecision.reset();//default="true"/>
    d->calcCompleted.reset();// default="true"/>
    d->calcOnSave.reset();// default="true"/>
    d->concurrentCalc.reset();// default="true"/>
    d->concurrentManualCount.reset();
    d->forceFullCalc.reset();
}

#if 0
bool Workbook::readChartCashe() const
{
    Q_D(const Workbook);
    return d->readChartCashe;
}

void Workbook::setReadChartCashe(bool read)
{
    Q_D(Workbook);
    d->readChartCashe = read;
}

bool Workbook::writeChartCashe() const
{
    Q_D(const Workbook);
    return d->writeChartCashe;
}

void Workbook::setWriteChartCashe(bool write)
{
    Q_D(Workbook);
    d->writeChartCashe = write;
}
#endif

WorkbookProtection Workbook::workbookProtection() const
{
     Q_D(const Workbook);
     return d->protection;
}

WorkbookProtection &Workbook::workbookProtection()
{
    Q_D(Workbook);
    return d->protection;
}

void Workbook::setWorkbookProtection(const WorkbookProtection &protection)
{
    Q_D(Workbook);
    d->protection = protection;
}

bool Workbook::isWorkbookProtected() const
{
    Q_D(const Workbook);
    return d->protection.isValid();
}

bool Workbook::isPasswordProtectionSet() const
{
    Q_D(const Workbook);
    if (d->protection.isValid()) {
        return !d->protection.protection().algorithmName.isEmpty() and
               !d->protection.protection().hashValue.isEmpty();
    }
    return false;
}

bool Workbook::setPasswordProtection(const QString &algorithm, const QString &password, const QString &salt, int spinCount)
{
    Q_D(Workbook);
    if (password.isEmpty() || PasswordProtection::algorithmForName(algorithm) == -1)
        return false;
    d->protection.protection() = PasswordProtection(password, algorithm, salt, spinCount);
    return true;
}

void Workbook::removeWorkbookProtection()
{
    Q_D(Workbook);
    d->protection = {};
}

bool Workbook::addDefinedName(const QString &name,
                                      const QString &formula,
                                      const QString &scope)
{
    Q_D(Workbook);

    if (name.isEmpty() || formula.isEmpty())
        return false;
    if (std::find_if(d->definedNamesList.constBegin(),
                     d->definedNamesList.constEnd(),
                     [name](const DefinedName &dn) { return dn.name == name; })
        != d->definedNamesList.constEnd())
        return false;

    //Remove the = sign from the formula if it exists.
    QString formulaString = formula;
    if (formulaString.startsWith(QLatin1Char('=')))
        formulaString = formula.mid(1);

    int id = -1;
    if (!scope.isEmpty()) {
        for (int i = 0; i < d->sheets.size(); ++i) {
            if (d->sheets[i]->name() == scope) {
                id = d->sheets[i]->id();
                break;
            }
        }
    }

    d->definedNamesList.append(DefinedName(name, formulaString, id));
    return true;
}

bool Workbook::removeDefinedName(const QString &name)
{
    Q_D(Workbook);
#if QT_VERSION >= QT_VERSION_CHECK(6, 1, 0)
    return d->definedNamesList.removeIf([name](const DefinedName &d) { return d.name == name; })
           > 0;
#else
    for (int i = 0; i < d->definedNamesList.size(); ++i) {
        if (d->definedNamesList.at(i).name == name) {
            d->definedNamesList.removeAt(i);
            return true;
        }
    }
    return false;
#endif
}

bool Workbook::removeDefinedName(DefinedName *name)
{
    return removeDefinedName(name ? name->name : QString());
}

bool Workbook::hasDefinedName(const QString &name) const
{
    Q_D(const Workbook);
    return std::find_if(d->definedNamesList.constBegin(),
                        d->definedNamesList.constEnd(),
                        [name](const DefinedName &dn) { return dn.name == name; })
           != d->definedNamesList.constEnd();
}

QStringList Workbook::definedNames() const
{
    Q_D(const Workbook);
    QStringList result;
    std::transform(d->definedNamesList.constBegin(),
                   d->definedNamesList.constEnd(),
                   std::back_inserter(result),
                   [](const DefinedName &dn) { return dn.name; });
    return result;
}

DefinedName *Workbook::definedName(const QString &name)
{
    Q_D(Workbook);
    if (auto dn = std::find_if(d->definedNamesList.begin(),
                               d->definedNamesList.end(),
                               [name](DefinedName &dn) { return dn.name == name; });
        dn != d->definedNamesList.end())
        return &(*dn);
    return nullptr;
}

AbstractSheet *Workbook::addSheet(const QString &name, AbstractSheet::Type type)
{
    Q_D(Workbook);
    return insertSheet(d->sheets.size(), name, type);
}

Chartsheet *Workbook::addChartsheet(const QString &name)
{
    return dynamic_cast<Chartsheet*>(addSheet(name, AbstractSheet::Type::Chartsheet));
}
Worksheet *Workbook::addWorksheet(const QString &name)
{
    return dynamic_cast<Worksheet*>(addSheet(name));
}

/*!
 * \internal
 */
QStringList Workbook::sheetNames() const
{
    Q_D(const Workbook);
    QStringList result;
    for (const auto &s: qAsConst(d->sheets))
        result << s->name();
    return result;
}

QList<AbstractSheet *> Workbook::sheets(AbstractSheet::Type type) const
{
    Q_D(const Workbook);
    QList<AbstractSheet *> result;
    for (const auto &sheet: d->sheets)
        if (sheet->type() == type) result << sheet.get();
    return result;
}

QList<AbstractSheet *> Workbook::sheets() const
{
    Q_D(const Workbook);
    QList<AbstractSheet *> result;
    for (const auto &sheet: d->sheets)
        result << sheet.get();
    return result;
}

QList<Worksheet *> Workbook::worksheets() const
{
    Q_D(const Workbook);
    QList<Worksheet *> result;
    for (const auto &sheet: d->sheets)
        if (sheet->type() == AbstractSheet::Type::Worksheet)
            result << static_cast<Worksheet *>(sheet.get());
    return result;
}

QList<Chartsheet *> Workbook::chartsheets() const
{
    Q_D(const Workbook);
    QList<Chartsheet *> result;
    for (const auto &sheet: d->sheets)
        if (sheet->type() == AbstractSheet::Type::Chartsheet)
            result << static_cast<Chartsheet *>(sheet.get());
    return result;
}

/*!
 * \internal
 * Used only when load the xlsx file!!
 */
AbstractSheet *Workbook::addSheet(const QString &name, int sheetId, AbstractSheet::Type type)
{
    Q_D(Workbook);
    if (sheetId > d->lastSheetId)
        d->lastSheetId = sheetId;

    AbstractSheet *sheet = nullptr;
    switch (type) {
        case AbstractSheet::Type::Worksheet:
            sheet = new Worksheet(name, sheetId, this, F_LoadFromExists);
            break;
        case AbstractSheet::Type::Chartsheet:
            sheet = new Chartsheet(name, sheetId, this, F_LoadFromExists);
            break;
        default: {
            qWarning("unsupported sheet type.");
            break;
        }
    }

    if (sheet) {
        d->sheets.append(QSharedPointer<AbstractSheet>(sheet));
    }
    return sheet;
}

AbstractSheet *Workbook::insertSheet(int index, const QString &name, AbstractSheet::Type type)
{
    Q_D(Workbook);
    if (index > d->lastSheetId) {
        //User tries to insert, where no sheet has gone before.
        return nullptr;
    }
    QString sheetName = createSafeSheetName(name);

    if (!sheetName.isEmpty()) {
        //If user given an already in-used name, we should not continue any more!
        if (sheet(sheetName))
            return nullptr;
    }
    else {
        if (type == AbstractSheet::Type::Worksheet) {
            do {
                ++d->lastWorksheetIndex;
                sheetName = QStringLiteral("Sheet%1").arg(d->lastWorksheetIndex);
            } while (sheet(sheetName));
        }
        else if (type == AbstractSheet::Type::Chartsheet) {
            do {
                ++d->lastChartsheetIndex;
                sheetName = QStringLiteral("Chart%1").arg(d->lastChartsheetIndex);
            } while (sheet(sheetName));
        }
        else {
            qWarning("unsupported sheet type.");
            return 0;
        }
    }

    ++d->lastSheetId;

    AbstractSheet *sheet = NULL;
    if (type == AbstractSheet::Type::Worksheet) {
        sheet = new Worksheet(sheetName, d->lastSheetId, this, F_NewFromScratch);
    }
    else if (type == AbstractSheet::Type::Chartsheet) {
        sheet = new Chartsheet(sheetName, d->lastSheetId, this, F_NewFromScratch);
    }
    else {
        qWarning("unsupported sheet type.");
        Q_ASSERT(false);
    }

    d->sheets.insert(index, QSharedPointer<AbstractSheet>(sheet));
    if (d->views.isEmpty()) d->views << WorkbookView{};
    d->views.last().activeTab = index;

    return sheet;
}

AbstractSheet *Workbook::activeSheet() const
{
    Q_D(const Workbook);
    if (d->sheets.isEmpty())
        const_cast<Workbook *>(this)->addSheet();
    if (d->views.isEmpty()) d->views << WorkbookView{};
    return d->sheets[d->views.last().activeTab.value_or(0)].data();
}
Worksheet *Workbook::activeWorksheet() const
{
    AbstractSheet *st = activeSheet();
    if (st && st->type() == AbstractSheet::Type::Worksheet)
        return static_cast<Worksheet *>(st);
    return nullptr;
}

Chartsheet *Workbook::activeChartsheet() const
{
    AbstractSheet *st = activeSheet();
    if (st && st->type() == AbstractSheet::Type::Chartsheet)
        return static_cast<Chartsheet *>(st);
    return nullptr;
}

bool Workbook::setActiveSheet(int index)
{
    Q_D(Workbook);
    if (index < 0 || index >= d->sheets.size()) {
        //warning
        return false;
    }
    if (d->views.isEmpty()) d->views << WorkbookView{};
    d->views.last().activeTab = index;
    return true;
}

bool Workbook::renameSheet(int index, const QString &newName)
{
    Q_D(Workbook);
    if (index < 0 || index >= d->sheets.size())
        return false;

    return sheet(index)->rename(newName);
}

bool Workbook::renameSheet(const QString &oldName, const QString &newName)
{
    if (oldName == newName)
        return false;
    if (auto s = sheet(oldName)) return s->rename(newName);
    return false;
}

bool Workbook::deleteSheet(int index)
{
    Q_D(Workbook);
    if (d->sheets.size() <= 1)
        return false;
    if (index < 0 || index >= d->sheets.size())
        return false;
    d->sheets.removeAt(index);
    return true;
}

bool Workbook::moveSheet(int srcIndex, int dstIndex)
{
    Q_D(Workbook);
    if (srcIndex == dstIndex)
        return false;

    if (srcIndex < 0 || srcIndex >= d->sheets.size())
        return false;

    QSharedPointer<AbstractSheet> sheet = d->sheets.takeAt(srcIndex);
    if (dstIndex >= 0 || dstIndex <= d->sheets.size()) {
        d->sheets.insert(dstIndex, sheet);
    }
    else {
        d->sheets.append(sheet);
    }
    return true;
}

bool Workbook::moveSheet(const QString &sheetName, int dstIndex)
{
    int srcIndex = sheetNames().indexOf(sheetName);
    if (srcIndex >= 0) return moveSheet(srcIndex, dstIndex);
    return false;
}

bool Workbook::copySheet(int index, const QString &newName)
{
    Q_D(Workbook);
    if (index < 0 || index >= d->sheets.size())
        return false;

    QString worksheetName = createSafeSheetName(newName);
    if (!worksheetName.isEmpty()) {
        //If user given an already in-used name, we should not continue any more!
        if (sheet(worksheetName))
            return false;
    }
    else {
        int copy_index = 1;
        do {
            ++copy_index;
            worksheetName = QStringLiteral("%1(%2)").arg(d->sheets[index]->name()).arg(copy_index);
        } while (sheet(worksheetName));
    }

    ++d->lastSheetId;
    AbstractSheet *sheet = d->sheets[index]->copy(worksheetName, d->lastSheetId);
    d->sheets.append(QSharedPointer<AbstractSheet>(sheet));

    return true; // #162
}

int Workbook::sheetsCount() const
{
    Q_D(const Workbook);
    return d->sheets.count();
}

int Workbook::sheetsCount(AbstractSheet::Type type) const
{
    Q_D(const Workbook);
    return std::count_if(d->sheets.constBegin(), d->sheets.constEnd(), [type](QSharedPointer<AbstractSheet> sheet){
        return (sheet.get() && sheet->type() == type);
    });
}

AbstractSheet *Workbook::sheet(int index) const
{
    Q_D(const Workbook);
    if (index < 0 || index >= d->sheets.size())
        return 0;
    return d->sheets.at(index).data();
}

AbstractSheet *Workbook::sheet(const QString &name) const
{
    Q_D(const Workbook);
    for (const auto &s: d->sheets) {
        if (s->name() == name) return s.data();
    }
    return nullptr;
}

SharedStrings *Workbook::sharedStrings() const
{
    Q_D(const Workbook);
    return d->sharedStrings.data();
}

Styles *Workbook::styles()
{
    Q_D(Workbook);
    return d->styles.data();
}

Theme *Workbook::theme()
{
    Q_D(Workbook);
    return d->theme.data();
}

/*!
 * \internal
 *
 * Unlike media files, drawing file is a property of the sheet.
 */
QList<Drawing *> Workbook::drawings()
{
    Q_D(Workbook);
    QList<Drawing *> ds;
    for (int i = 0; i < d->sheets.size(); ++i) {
        QSharedPointer<AbstractSheet> sheet = d->sheets[i];
        if (sheet->drawing())
            ds.append(sheet->drawing());
    }

    return ds;
}

/*!
 * \internal
 */
QList<QSharedPointer<AbstractSheet> > Workbook::getSheetsByType(AbstractSheet::Type type) const
{
    Q_D(const Workbook);
    QList<QSharedPointer<AbstractSheet> > list;
    for (int i = 0; i < d->sheets.size(); ++i) {
        if (d->sheets[i]->type() == type)
            list.append(d->sheets[i]);
    }
    return list;
}

void Workbook::saveToXmlFile(QIODevice *device) const
{
    Q_D(const Workbook);
    d->relationships->clear();
    if (d->sheets.isEmpty())
        const_cast<Workbook *>(this)->addSheet();

    QXmlStreamWriter writer(device);

    writer.writeStartDocument(QStringLiteral("1.0"), true);
    writer.writeStartElement(QStringLiteral("workbook"));
    writer.writeAttribute(QStringLiteral("xmlns"),
                          QStringLiteral("http://schemas.openxmlformats.org/spreadsheetml/2006/main"));
    writer.writeAttribute(QStringLiteral("xmlns:r"),
                          QStringLiteral("http://schemas.openxmlformats.org/officeDocument/2006/relationships"));

    // 1. fileVersion
    writer.writeEmptyElement(QStringLiteral("fileVersion"));
    writeAttribute(writer, QLatin1String("appName"), d->appName);
    writeAttribute(writer, QLatin1String("lastEdited"), d->lastEdited);
    writeAttribute(writer, QLatin1String("lowestEdited"), d->lowestEdited);
    writeAttribute(writer, QLatin1String("rupBuild"), d->rupBuild);
    writeAttribute(writer, QLatin1String("codeName"), d->appCodeName);
    //2. fileSharing
    if (d->readOnlyRecommended.has_value() || !d->userName.isEmpty() || !d->algorithmName.isEmpty()
        || !d->hashValue.isEmpty() || !d->saltValue.isEmpty() || !d->spinCount.has_value()) {
        writer.writeEmptyElement(QLatin1String("fileSharing"));
        writeAttribute(writer, QLatin1String("readOnlyRecommended"), d->readOnlyRecommended);
        writeAttribute(writer, QLatin1String("userName"), d->userName);
        writeAttribute(writer, QLatin1String("algorithmName"), d->algorithmName);
        writeAttribute(writer, QLatin1String("hashValue"), d->hashValue);
        writeAttribute(writer, QLatin1String("saltValue"), d->saltValue);
        writeAttribute(writer, QLatin1String("spinCount"), d->spinCount);
    }
    //3. workbookPr
    writer.writeEmptyElement(QStringLiteral("workbookPr"));
    writeAttribute(writer, QLatin1String("date1904"), d->date1904);
    if (d->showObjects.has_value())
        writeAttribute(writer, QLatin1String("showObjects"), toString(d->showObjects.value()));
    writeAttribute(writer, QLatin1String("showBorderUnselectedTables"), d->showBorderUnselectedTables);
    writeAttribute(writer, QLatin1String("filterPrivacy"), d->filterPrivacy);
    writeAttribute(writer, QLatin1String("promptedSolutions"), d->promptedSolutions);
    writeAttribute(writer, QLatin1String("showInkAnnotation"), d->showInkAnnotation);
    writeAttribute(writer, QLatin1String("backupFile"), d->backupFile);
    writeAttribute(writer, QLatin1String("saveExternalLinkValues"), d->saveExternalLinkValues);
    if (d->updateLinks.has_value())
        writeAttribute(writer, QLatin1String("updateLinks"), toString(d->updateLinks.value()));
    writeAttribute(writer, QLatin1String("codeName"), d->codeName);
    writeAttribute(writer, QLatin1String("hidePivotFieldList"), d->hidePivotFieldList);
    writeAttribute(writer, QLatin1String("showPivotChartFilter"), d->showPivotChartFilter);
    writeAttribute(writer, QLatin1String("allowRefreshQuery"), d->allowRefreshQuery);
    writeAttribute(writer, QLatin1String("publishItems"), d->publishItems);
    writeAttribute(writer, QLatin1String("checkCompatibility"), d->checkCompatibility);
    writeAttribute(writer, QLatin1String("autoCompressPictures"), d->autoCompressPictures);
    writeAttribute(writer, QLatin1String("refreshAllConnections"), d->refreshAllConnections);
    writeAttribute(writer, QLatin1String("defaultThemeVersion"), d->defaultThemeVersion.value_or(124226));

    // 4. workbookProtection
    d->protection.write(writer);
    // 5. bookViews
    writer.writeStartElement(QStringLiteral("bookViews"));
    if (d->views.isEmpty()) d->views << WorkbookView{};
    for (const auto &view: qAsConst(d->views))
        view.write(writer);
    writer.writeEndElement(); //bookViews
    // 6. sheets
    writer.writeStartElement(QStringLiteral("sheets"));
    int worksheetIndex = 0;
    int chartsheetIndex = 0;
    for (int i = 0; i < d->sheets.size(); ++i) {
        QSharedPointer<AbstractSheet> sheet = d->sheets[i];
        writer.writeEmptyElement(QStringLiteral("sheet"));
        writer.writeAttribute(QStringLiteral("name"), sheet->name());
        writer.writeAttribute(QStringLiteral("sheetId"), QString::number(sheet->id()));
        writer.writeAttribute(QStringLiteral("state"), AbstractSheet::toString(sheet->visibility()));

        if (sheet->type() == AbstractSheet::Type::Worksheet)
            d->relationships->addDocumentRelationship(QStringLiteral("/worksheet"),
                                                      QStringLiteral("worksheets/sheet%1.xml")
                                                          .arg(++worksheetIndex));
        else if (sheet->type() == AbstractSheet::Type::Chartsheet)
            d->relationships->addDocumentRelationship(QStringLiteral("/chartsheet"),
                                                      QStringLiteral("chartsheets/sheet%1.xml")
                                                          .arg(++chartsheetIndex));

        writer.writeAttribute(QStringLiteral("r:id"),
                              QStringLiteral("rId%1").arg(d->relationships->count()));
    }
    writer.writeEndElement(); //sheets
    // 7. functionGroups
    if (!d->functionGroups.isEmpty() || d->builtInGroupCount.has_value()) {
        writer.writeStartElement(QLatin1String("functionGroups"));
        for (const auto &g: qAsConst(d->functionGroups))
            writer.writeTextElement(QLatin1String("functionGroup"), g);
        writeAttribute(writer, QLatin1String("builtInGroupCount"), d->builtInGroupCount);
        writer.writeEndElement();
    }
    // 8. externalReferences
    if (d->externalLinks.size() > 0) {
        writer.writeStartElement(QStringLiteral("externalReferences"));
        for (int i = 0; i < d->externalLinks.size(); ++i) {
            writer.writeEmptyElement(QStringLiteral("externalReference"));
            d->relationships->addDocumentRelationship(QStringLiteral("/externalLink"),
                                                      QStringLiteral(
                                                          "externalLinks/externalLink%1.xml")
                                                          .arg(i + 1));
            writer.writeAttribute(QStringLiteral("r:id"),
                                  QStringLiteral("rId%1").arg(d->relationships->count()));
        }
        writer.writeEndElement(); //externalReferences
    }
    // 9. definedNames
    if (!d->definedNamesList.isEmpty()) {
        writer.writeStartElement(QStringLiteral("definedNames"));
        for (const auto &data : qAsConst(d->definedNamesList)) {
            writer.writeStartElement(QStringLiteral("definedName"));
            writer.writeAttribute(QStringLiteral("name"), data.name);
            writeAttribute(writer, QLatin1String("comment"), data.comment);
            writeAttribute(writer, QLatin1String("description"), data.description);
            if (data.sheetId != -1) {
                //find the local index of the sheet.
                for (int i = 0; i < d->sheets.size(); ++i) {
                    if (d->sheets[i]->id() == data.sheetId) {
                        writer.writeAttribute(QStringLiteral("localSheetId"), QString::number(i));
                        break;
                    }
                }
            }
            writeAttribute(writer, QLatin1String("hidden"), data.hidden);
            writeAttribute(writer, QLatin1String("workbookParameter"), data.workbookParameter);
            writer.writeCharacters(data.formula);
            writer.writeEndElement(); //definedName
        }
        writer.writeEndElement(); //definedNames
    }
    // 10. calcPr
    writer.writeStartElement(QStringLiteral("calcPr"));
    writeAttribute(writer, QLatin1String("calcId"), d->calcId);
    if (d->calcMode.has_value())
        writeAttribute(writer, QLatin1String("calcMode"), toString(d->calcMode.value()));
    writeAttribute(writer, QLatin1String("fullCalcOnLoad"), d->fullCalcOnLoad);
    if (d->refMode.has_value())
        writeAttribute(writer, QLatin1String("refMode"), toString(d->refMode.value()));
    writeAttribute(writer, QLatin1String("iterate"), d->iterate);
    writeAttribute(writer, QLatin1String("iterateCount"), d->iterateCount);
    writeAttribute(writer, QLatin1String("iterateDelta"), d->iterateDelta);
    writeAttribute(writer, QLatin1String("fullPrecision"), d->fullPrecision);
    writeAttribute(writer, QLatin1String("calcCompleted"), d->calcCompleted);
    writeAttribute(writer, QLatin1String("calcOnSave"), d->calcOnSave);
    writeAttribute(writer, QLatin1String("concurrentCalc"), d->concurrentCalc);
    writeAttribute(writer, QLatin1String("concurrentManualCount"), d->concurrentManualCount);
    writeAttribute(writer, QLatin1String("forceFullCalc"), d->forceFullCalc);
    writer.writeEndElement(); //calcPr

    // 11. oleSize
    if (d->oleSize.isValid()) {
        writer.writeEmptyElement(QLatin1String("oleSize"));
        writeAttribute(writer, QLatin1String("ref"), d->oleSize.toString());
    }
    // 12. customWorkbookViews
    //TODO:
    // 13. pivotCaches
    //TODO:
    // 14. smartTagPr
    //TODO:
    // 15. smartTagTypes
    //TODO:
    // 16. webPublishing
    if (d->css.has_value() || d->thicket.has_value() || d->longFileNames.has_value()
        || d->vml.has_value() || d->allowPng.has_value() || d->dpi.has_value()
        || !d->targetScreenSize.isEmpty() || !d->characterSet.isEmpty()) {
        writer.writeEmptyElement(QLatin1String("webPublishing"));
        writeAttribute(writer, QLatin1String("css"), d->css);
        writeAttribute(writer, QLatin1String("thicket"), d->thicket);
        writeAttribute(writer, QLatin1String("longFileNames"), d->longFileNames);
        writeAttribute(writer, QLatin1String("vml"), d->vml);
        writeAttribute(writer, QLatin1String("allowPng"), d->allowPng);
        writeAttribute(writer, QLatin1String("targetScreenSize"), d->targetScreenSize);
        writeAttribute(writer, QLatin1String("dpi"), d->dpi);
        writeAttribute(writer, QLatin1String("characterSet"), d->characterSet);
    }
    // 17. fileRecoveryPr
    if (d->autoRecover.has_value() || d->crashSave.has_value() ||
        d->dataExtractLoad.has_value() || d->repairLoad.has_value()) {
        writer.writeEmptyElement(QLatin1String("fileRecoveryPr"));
        writeAttribute(writer, QLatin1String("autoRecover"), d->autoRecover);
        writeAttribute(writer, QLatin1String("crashSave"), d->crashSave);
        writeAttribute(writer, QLatin1String("dataExtractLoad"), d->dataExtractLoad);
        writeAttribute(writer, QLatin1String("repairLoad"), d->repairLoad);
    }
    // 18. webPublishObjects
    if (!d->webPublishObjects.isEmpty()) {
        writer.writeStartElement(QLatin1String("webPublishObjects"));
       // writeAttribute(writer, QLatin1String("count"), d->webPublishObjects.size());   //\QXlsx\source\xlsxworkbook.cpp</a>:906:23: error: call of overloaded 'writeAttribute(QXmlStreamWriter&amp;, QLatin1String, qsizetype)' is ambiguous
        for (const auto &w: qAsConst(d->webPublishObjects)) {
            writer.writeEmptyElement(QLatin1String("webPublishObject"));
            writeAttribute(writer, QLatin1String("id"), w.id);
            writeAttribute(writer, QLatin1String("divId"), w.divId);
            writeAttribute(writer, QLatin1String("sourceObject"), w.sourceObject);
            writeAttribute(writer, QLatin1String("destinationFile"), w.destinationFile);
            writeAttribute(writer, QLatin1String("title"), w.title);
            writeAttribute(writer, QLatin1String("autoRepublish"), w.autoRepublish);
        }
    }
    // 19. extLst
    d->extLst.write(writer, QLatin1String("extLst"));
    writer.writeEndElement(); //workbook
    writer.writeEndDocument();

    d->relationships->addDocumentRelationship(QStringLiteral("/theme"),
                                              QStringLiteral("theme/theme1.xml"));
    d->relationships->addDocumentRelationship(QStringLiteral("/styles"),
                                              QStringLiteral("styles.xml"));
    if (!sharedStrings()->isEmpty())
        d->relationships->addDocumentRelationship(QStringLiteral("/sharedStrings"),
                                                  QStringLiteral("sharedStrings.xml"));
}

bool Workbook::loadFromXmlFile(QIODevice *device)
{
    Q_D(Workbook);

    QXmlStreamReader reader(device);
    while (!reader.atEnd()) {
        QXmlStreamReader::TokenType token = reader.readNext();
        if (token == QXmlStreamReader::StartElement) {
            const auto &attributes = reader.attributes();
            if (reader.name() == QLatin1String("sheet")) {
                const auto &name = attributes.value(QLatin1String("name")).toString();

                int sheetId = attributes.value(QLatin1String("sheetId")).toInt();

                const auto &rId = attributes.value(QLatin1String("r:id")).toString();

                auto state = AbstractSheet::Visibility::Visible;
                if (attributes.hasAttribute(QLatin1String("state"))) {
                    AbstractSheet::fromString(attributes.value(QLatin1String("state")).toString(), state);
                }

                XlsxRelationship relationship = d->relationships->getRelationshipById(rId);

                auto type = AbstractSheet::Type::Worksheet;
                if (relationship.type.endsWith(QLatin1String("/worksheet")))
                    type = AbstractSheet::Type::Worksheet;
                else if (relationship.type.endsWith(QLatin1String("/chartsheet")))
                    type = AbstractSheet::Type::Chartsheet;
                else if (relationship.type.endsWith(QLatin1String("/dialogsheet")))
                    type = AbstractSheet::Type::Dialogsheet;
                else if (relationship.type.endsWith(QLatin1String("/xlMacrosheet")))
                    type = AbstractSheet::Type::Macrosheet;
                else
                    qWarning() << "unknown sheet type : " << relationship.type;

                AbstractSheet *sheet = addSheet(name, sheetId, type);
                sheet->setVisibility(state);
                QString strFilePath = filePath();

                // const QString fullPath = QDir::cleanPath(splitPath(strFilePath).constFirst() + QLatin1String("/") + relationship.target);
                const auto parts = splitPath(strFilePath);
                QString fullPath = QDir::cleanPath(parts.first() + QLatin1String("/")
                                                   + relationship.target);

                sheet->setFilePath(fullPath);
            }
            else if (reader.name() == QLatin1String("workbookProtection"))
                d->protection.read(reader);
            else if (reader.name() == QLatin1String("workbookPr")) {
                parseAttributeBool(attributes, QLatin1String("date1904"), d->date1904);
                if (attributes.hasAttribute(QLatin1String("showObjects"))) {
                    ShowObjects so;
                    fromString(attributes.value(QLatin1String("showObjects")).toString(), so);
                    d->showObjects = so;
                }
                parseAttributeBool(attributes, QLatin1String("showBorderUnselectedTables"), d->showBorderUnselectedTables);
                parseAttributeBool(attributes, QLatin1String("filterPrivacy"), d->filterPrivacy);
                parseAttributeBool(attributes, QLatin1String("promptedSolutions"), d->promptedSolutions);
                parseAttributeBool(attributes, QLatin1String("showInkAnnotation"), d->showInkAnnotation);
                parseAttributeBool(attributes, QLatin1String("backupFile"), d->backupFile);
                parseAttributeBool(attributes, QLatin1String("saveExternalLinkValues"), d->saveExternalLinkValues);
                if (attributes.hasAttribute(QLatin1String("updateLinks"))) {
                    UpdateLinks ul;
                    fromString(attributes.value(QLatin1String("updateLinks")).toString(), ul);
                    d->updateLinks = ul;
                }
                d->codeName = attributes.value(QLatin1String("codeName")).toString();
                parseAttributeBool(attributes, QLatin1String("hidePivotFieldList"), d->hidePivotFieldList);
                parseAttributeBool(attributes, QLatin1String("showPivotChartFilter"), d->showPivotChartFilter);
                parseAttributeBool(attributes, QLatin1String("allowRefreshQuery"), d->allowRefreshQuery);
                parseAttributeBool(attributes, QLatin1String("publishItems"), d->publishItems);
                parseAttributeBool(attributes, QLatin1String("checkCompatibility"), d->checkCompatibility);
                parseAttributeBool(attributes, QLatin1String("autoCompressPictures"), d->autoCompressPictures);
                parseAttributeBool(attributes, QLatin1String("refreshAllConnections"), d->refreshAllConnections);
                parseAttributeInt(attributes, QLatin1String("defaultThemeVersion"), d->defaultThemeVersion);
            }
            else if (reader.name() == QLatin1String("fileVersion")) {
                parseAttributeString(attributes, QLatin1String("appName"), d->appName);
                parseAttributeString(attributes, QLatin1String("lastEdited"), d->lastEdited);
                parseAttributeString(attributes, QLatin1String("lastEdited"), d->lastEdited);
                parseAttributeString(attributes, QLatin1String("lowestEdited"), d->lowestEdited);
                parseAttributeString(attributes, QLatin1String("rupBuild"), d->rupBuild);
                parseAttributeString(attributes, QLatin1String("codeName"), d->appCodeName);
            }
            else if (reader.name() == QLatin1String("fileRecoveryPr")) {
                parseAttributeBool(attributes, QLatin1String("autoRecover"), d->autoRecover);
                parseAttributeBool(attributes, QLatin1String("crashSave"), d->crashSave);
                parseAttributeBool(attributes, QLatin1String("dataExtractLoad"), d->dataExtractLoad);
                parseAttributeBool(attributes, QLatin1String("repairLoad"), d->repairLoad);
            }
            else if (reader.name() == QLatin1String("oleSize")) {
                d->oleSize = CellRange(attributes.value(QLatin1String("ref")).toString());
            }
            else if (reader.name() == QLatin1String("webPublishing")) {
                parseAttributeBool(attributes, QLatin1String("css"), d->css);
                parseAttributeBool(attributes, QLatin1String("thicket"), d->thicket);
                parseAttributeBool(attributes, QLatin1String("longFileNames"), d->longFileNames);
                parseAttributeBool(attributes, QLatin1String("vml"), d->vml);
                parseAttributeBool(attributes, QLatin1String("allowPng"), d->allowPng);
                parseAttributeString(attributes, QLatin1String("targetScreenSize"), d->targetScreenSize);
                parseAttributeInt(attributes, QLatin1String("dpi"), d->dpi);
                parseAttributeString(attributes, QLatin1String("characterSet"), d->characterSet);
            }
            else if (reader.name() == QLatin1String("fileSharing")) {
                parseAttributeBool(attributes, QLatin1String("readOnlyRecommended"), d->readOnlyRecommended);
                parseAttributeString(attributes, QLatin1String("userName"), d->userName);
                parseAttributeString(attributes, QLatin1String("algorithmName"), d->algorithmName);
                parseAttributeString(attributes, QLatin1String("hashValue"), d->hashValue);
                parseAttributeString(attributes, QLatin1String("saltValue"), d->saltValue);
                parseAttributeInt(attributes, QLatin1String("spinCount"), d->spinCount);
            }
            else if (reader.name() == QLatin1String("webPublishObject")) {
                WebPublishObject w;
                parseAttributeInt(attributes, QLatin1String("id"), w.id);
                parseAttributeString(attributes, QLatin1String("divId"), w.divId);
                parseAttributeString(attributes, QLatin1String("sourceObject"), w.sourceObject);
                parseAttributeString(attributes, QLatin1String("destinationFile"), w.destinationFile);
                parseAttributeString(attributes, QLatin1String("title"), w.title);
                parseAttributeBool(attributes, QLatin1String("autoRepublish"), w.autoRepublish);
                d->webPublishObjects << w;
            }
            else if (reader.name() == QLatin1String("workbookView")) {
                WorkbookView view;
                view.read(reader);
                d->views << view;
            }
            else if (reader.name() == QLatin1String("functionGroups")) {
                parseAttributeInt(attributes, QLatin1String("builtInGroupCount"), d->builtInGroupCount);
            }
            else if (reader.name() == QLatin1String("functionGroup")) {
                QString g = reader.readElementText();
                if (!g.isEmpty()) d->functionGroups << g;
            }
            else if (reader.name() == QLatin1String("externalReference")) {
                const QString rId = attributes.value(QLatin1String("r:id")).toString();
                XlsxRelationship relationship = d->relationships->getRelationshipById(rId);

                QSharedPointer<SimpleOOXmlFile> link(new SimpleOOXmlFile(F_LoadFromExists));

                const auto parts = splitPath(filePath());
                QString fullPath = QDir::cleanPath(parts.first() + QLatin1String("/")
                                                   + relationship.target);

                link->setFilePath(fullPath);
                d->externalLinks.append(link);
            }
            else if (reader.name() == QLatin1String("calcPr")) {
                parseAttributeInt(attributes, QLatin1String("calcId"), d->calcId);
                if (attributes.hasAttribute(QLatin1String("calcMode"))) {
                    CalculationMode m; fromString(attributes.value(QLatin1String("calcMode")).toString(), m);
                    d->calcMode = m;
                }
                parseAttributeBool(attributes, QLatin1String("fullCalcOnLoad"), d->fullCalcOnLoad);
                if (attributes.hasAttribute(QLatin1String("refMode"))) {
                    ReferenceMode m; fromString(attributes.value(QLatin1String("refMode")).toString(), m);
                    d->refMode = m;
                }
                parseAttributeBool(attributes, QLatin1String("iterate"), d->iterate);
                parseAttributeInt(attributes, QLatin1String("iterateCount"), d->iterateCount);
                parseAttributeDouble(attributes, QLatin1String("iterateDelta"), d->iterateDelta);
                parseAttributeBool(attributes, QLatin1String("fullPrecision"), d->fullPrecision);
                parseAttributeBool(attributes, QLatin1String("calcCompleted"), d->calcCompleted);
                parseAttributeBool(attributes, QLatin1String("calcOnSave"), d->calcOnSave);
                parseAttributeBool(attributes, QLatin1String("concurrentCalc"), d->concurrentCalc);
                parseAttributeInt(attributes, QLatin1String("concurrentManualCount"), d->concurrentManualCount);
                parseAttributeBool(attributes, QLatin1String("forceFullCalc"), d->forceFullCalc);
            }
            else if (reader.name() == QLatin1String("definedName")) {
                DefinedName data;

                data.name = attributes.value(QLatin1String("name")).toString();
                parseAttributeString(attributes, QLatin1String("comment"), data.comment);
                parseAttributeString(attributes, QLatin1String("description"), data.description);
                parseAttributeBool(attributes, QLatin1String("hidden"), data.hidden);
                parseAttributeBool(attributes, QLatin1String("workbookParameter"), data.workbookParameter);
                if (attributes.hasAttribute(QLatin1String("localSheetId"))) {
                    int localId = attributes.value(QLatin1String("localSheetId")).toInt();
                    int sheetId = d->sheets.at(localId)->id();
                    data.sheetId = sheetId;
                }
                data.formula = reader.readElementText();
                d->definedNamesList.append(data);
            }
            else if (reader.name() == QLatin1String("extLst"))
                d->extLst.read(reader);
        }
    }
    return true;
}

/*!
 * \internal
 */
QList<QWeakPointer<MediaFile> > Workbook::mediaFiles() const
{
    Q_D(const Workbook);

    return d->mediaFiles;
}

/*!
 * \internal
 */
bool Workbook::addMediaFile(QSharedPointer<MediaFile> media, bool force)
{
    Q_D(Workbook);

    if (!force) {
        for (int i = 0; i < d->mediaFiles.size(); ++i) {
            if (auto mediaFile = d->mediaFiles[i].lock()) {
                if (mediaFile->hashKey() == media->hashKey()) {
                    media->setIndex(i);
                    return false;
                }
            }
        }
    }

    media->setIndex(d->mediaFiles.size());
    d->mediaFiles.append(media);
    return true;
}

void Workbook::removeMediaFile(QSharedPointer<MediaFile> media)
{
    Q_D(Workbook);
    d->mediaFiles.removeOne(media);
    //reindex
    int index = 0;
    for (auto &media : d->mediaFiles) {
        if (auto mf = media.lock())
            mf->setIndex(index++);
    }
}

/*!
 * \internal
 */
QList<QWeakPointer<Chart> > Workbook::chartFiles() const
{
    Q_D(const Workbook);

    return d->chartFiles;
}

/*!
 * \internal
 */
void Workbook::addChartFile(const QSharedPointer<Chart> &chart)
{
    Q_D(Workbook);

    if (!d->chartFiles.contains(chart))
        d->chartFiles.append(chart);
}

/*!
 * \internal
 */
void Workbook::removeChartFile(const QSharedPointer<Chart> &chart)
{
    Q_D(Workbook);

    d->chartFiles.removeOne(chart);
}

} // namespace QXlsx
