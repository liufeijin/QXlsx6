// xlsxworkbook.cpp

#include <QtGlobal>
#include <QXmlStreamWriter>
#include <QXmlStreamReader>
#include <QFile>
#include <QBuffer>
#include <QDir>
#include <QtDebug>

#include "xlsxworkbook.h"
#include "xlsxworkbook_p.h"
#include "xlsxsharedstrings_p.h"
#include "xlsxworksheet.h"
#include "xlsxchartsheet.h"
#include "xlsxstyles_p.h"
#include "xlsxmediafile_p.h"
#include "xlsxutility_p.h"
#include "xlsxchart.h"

namespace QXlsx {

WorkbookPrivate::WorkbookPrivate(Workbook *q, Workbook::CreateFlag flag) :
    AbstractOOXmlFilePrivate(q, flag)
{
    sharedStrings = QSharedPointer<SharedStrings> (new SharedStrings(flag));
    styles = QSharedPointer<Styles>(new Styles(flag));
    theme = QSharedPointer<Theme>(new Theme(flag));

    x_window = 240;
    y_window = 15;
    window_width = 16095;
    window_height = 9660;

    strings_to_numbers_enabled = false;
    strings_to_hyperlinks_enabled = true;
    html_to_richstring_enabled = false;
    date1904 = false;
    defaultDateFormat = QStringLiteral("yyyy-mm-dd");
    activesheetIndex = 0;
    firstsheet = 0;
    table_count = 0;
}

Workbook::Workbook(CreateFlag flag)
    : AbstractOOXmlFile(new WorkbookPrivate(this, flag))
{

}

Workbook::~Workbook()
{
}

bool Workbook::isDate1904() const
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

/*!
 * \brief Create a defined name in the workbook.
 * \param name The defined name
 * \param formula The cell or range that the defined name refers to.
 * \param comment
 * \param scope The name of one worksheet, or empty which means global scope.
 * \return Return false if the name invalid.
 */
bool Workbook::defineName(const QString &name, const QString &formula, const QString &comment, const QString &scope)
{
    Q_D(Workbook);

    //Remove the = sign from the formula if it exists.
    QString formulaString = formula;
    if (formulaString.startsWith(QLatin1Char('=')))
        formulaString = formula.mid(1);

    int id=-1;
    if (!scope.isEmpty()) {
        for (int i=0; i<d->sheets.size(); ++i) {
            if (d->sheets[i]->name() == scope) {
                id = d->sheets[i]->id();
                break;
            }
        }
    }

    d->definedNamesList.append(XlsxDefineNameData(name, formulaString, comment, id));
    return true;
}

AbstractSheet *Workbook::addSheet(const QString &name, AbstractSheet::Type type)
{
    Q_D(Workbook);
    return insertSheet(d->sheets.size(), name, type);
}

/*!
 * \internal
 */
QStringList Workbook::worksheetNames() const
{
    Q_D(const Workbook);
    return d->sheetNames;
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
           sheet = new Worksheet(name, sheetId, this, F_LoadFromExists); break;
        case AbstractSheet::Type::Chartsheet:
           sheet = new Chartsheet(name, sheetId, this, F_LoadFromExists); break;
        default: {
            qWarning("unsupported sheet type.");
            break;
        }
    }

    if (sheet) {
        d->sheets.append(QSharedPointer<AbstractSheet>(sheet));
        d->sheetNames.append(name);
    }
    return sheet;
}

AbstractSheet *Workbook::insertSheet(int index, const QString &name, AbstractSheet::Type type)
{
    Q_D(Workbook);
    if (index > d->lastSheetId){
        //User tries to insert, where no sheet has gone before.
        return nullptr;
    }
    QString sheetName = createSafeSheetName(name);

    if (!sheetName.isEmpty()) {
        //If user given an already in-used name, we should not continue any more!
        if (d->sheetNames.contains(sheetName))
            return 0;
    } else {
        if (type == AbstractSheet::Type::Worksheet) {
            do {
                ++d->lastWorksheetIndex;
                sheetName = QStringLiteral("Sheet%1").arg(d->lastWorksheetIndex);
            } while (d->sheetNames.contains(sheetName));
        } else if (type == AbstractSheet::Type::Chartsheet) {
            do {
                ++d->lastChartsheetIndex;
                sheetName = QStringLiteral("Chart%1").arg(d->lastChartsheetIndex);
            } while (d->sheetNames.contains(sheetName));
        } else {
            qWarning("unsupported sheet type.");
            return 0;
        }
    }

    ++d->lastSheetId;

    AbstractSheet *sheet = NULL;
    if ( type == AbstractSheet::Type::Worksheet )
    {
        sheet = new Worksheet(sheetName, d->lastSheetId, this, F_NewFromScratch);
    }
    else if ( type == AbstractSheet::Type::Chartsheet )
    {
        sheet = new Chartsheet(sheetName, d->lastSheetId, this, F_NewFromScratch);
    }
    else
    {
        qWarning("unsupported sheet type.");
        Q_ASSERT(false);
    }

    d->sheets.insert(index, QSharedPointer<AbstractSheet>(sheet));
    d->sheetNames.insert(index, sheetName);
    d->activesheetIndex = index;

    return sheet;
}

AbstractSheet *Workbook::activeSheet() const
{
    Q_D(const Workbook);
    if (d->sheets.isEmpty())
        const_cast<Workbook*>(this)->addSheet();
    return d->sheets[d->activesheetIndex].data();
}

bool Workbook::setActiveSheet(int index)
{
    Q_D(Workbook);
    if (index < 0 || index >= d->sheets.size()) {
        //warning
        return false;
    }
    d->activesheetIndex = index;
    return true;
}

/*!
 * Rename the worksheet at the \a index to \a newName.
 */
bool Workbook::renameSheet(int index, const QString &newName)
{
    Q_D(Workbook);
    QString name = createSafeSheetName(newName);
    if (index < 0 || index >= d->sheets.size())
        return false;

    //If user given an already in-used name, return false
    for (int i=0; i<d->sheets.size(); ++i) {
        if (d->sheets[i]->name() == name)
            return false;
    }

    d->sheets[index]->setName(name);
    d->sheetNames[index] = name;
    return true;
}

/*!
 * Remove the worksheet at pos \a index.
 */
bool Workbook::deleteSheet(int index)
{
    Q_D(Workbook);
    if (d->sheets.size() <= 1)
        return false;
    if (index < 0 || index >= d->sheets.size())
        return false;
    d->sheets.removeAt(index);
    d->sheetNames.removeAt(index);
    return true;
}

/*!
 * Moves the worksheet form \a srcIndex to \a distIndex.
 */
bool Workbook::moveSheet(int srcIndex, int distIndex)
{
    Q_D(Workbook);
    if (srcIndex == distIndex)
        return false;

    if (srcIndex < 0 || srcIndex >= d->sheets.size())
        return false;

    QSharedPointer<AbstractSheet> sheet = d->sheets.takeAt(srcIndex);
    d->sheetNames.takeAt(srcIndex);
    if (distIndex >= 0 || distIndex <= d->sheets.size()) {
        d->sheets.insert(distIndex, sheet);
        d->sheetNames.insert(distIndex, sheet->name());
    } else {
        d->sheets.append(sheet);
        d->sheetNames.append(sheet->name());
    }
    return true;
}

bool Workbook::copySheet(int index, const QString &newName)
{
    Q_D(Workbook);
    if (index < 0 || index >= d->sheets.size())
        return false;

    QString worksheetName = createSafeSheetName(newName);
    if (!newName.isEmpty()) {
        //If user given an already in-used name, we should not continue any more!
        if (d->sheetNames.contains(newName))
            return false;
    } else {
        int copy_index = 1;
        do {
            ++copy_index;
            worksheetName = QStringLiteral("%1(%2)").arg(d->sheets[index]->name()).arg(copy_index);
        } while (d->sheetNames.contains(worksheetName));
    }

    ++d->lastSheetId;
    AbstractSheet *sheet = d->sheets[index]->copy(worksheetName, d->lastSheetId);
    d->sheets.append(QSharedPointer<AbstractSheet> (sheet));
    d->sheetNames.append(sheet->name());

    return true; // #162
}

/*!
 * Returns count of worksheets.
 */
int Workbook::sheetCount() const
{
    Q_D(const Workbook);
    return d->sheets.count();
}

/*!
 * Returns the sheet object at index \a sheetIndex.
 */
AbstractSheet *Workbook::sheet(int index) const
{
    Q_D(const Workbook);
    if (index < 0 || index >= d->sheets.size())
        return 0;
    return d->sheets.at(index).data();
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
    for (int i=0; i<d->sheets.size(); ++i) {
        QSharedPointer<AbstractSheet> sheet = d->sheets[i];
        if (sheet->drawing())
        ds.append(sheet->drawing());
    }

    return ds;
}

/*!
 * \internal
 */
QList<QSharedPointer<AbstractSheet> > Workbook::getSheetsByTypes(AbstractSheet::Type type) const
{
    Q_D(const Workbook);
    QList<QSharedPointer<AbstractSheet> > list;
    for (int i=0; i<d->sheets.size(); ++i) {
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
    writer.writeAttribute(QStringLiteral("xmlns"), QStringLiteral("http://schemas.openxmlformats.org/spreadsheetml/2006/main"));
    writer.writeAttribute(QStringLiteral("xmlns:r"), QStringLiteral("http://schemas.openxmlformats.org/officeDocument/2006/relationships"));

    writer.writeEmptyElement(QStringLiteral("fileVersion"));
    writer.writeAttribute(QStringLiteral("appName"), QStringLiteral("xl"));
    writer.writeAttribute(QStringLiteral("lastEdited"), QStringLiteral("4"));
    writer.writeAttribute(QStringLiteral("lowestEdited"), QStringLiteral("4"));
    writer.writeAttribute(QStringLiteral("rupBuild"), QStringLiteral("4505"));
//    writer.writeAttribute(QStringLiteral("codeName"), QStringLiteral("{37E998C4-C9E5-D4B9-71C8-EB1FF731991C}"));

    writer.writeEmptyElement(QStringLiteral("workbookPr"));
    if (d->date1904)
        writer.writeAttribute(QStringLiteral("date1904"), QStringLiteral("1"));
    writer.writeAttribute(QStringLiteral("defaultThemeVersion"), QStringLiteral("124226"));

    writer.writeStartElement(QStringLiteral("bookViews"));
    writer.writeEmptyElement(QStringLiteral("workbookView"));
    writer.writeAttribute(QStringLiteral("xWindow"), QString::number(d->x_window));
    writer.writeAttribute(QStringLiteral("yWindow"), QString::number(d->y_window));
    writer.writeAttribute(QStringLiteral("windowWidth"), QString::number(d->window_width));
    writer.writeAttribute(QStringLiteral("windowHeight"), QString::number(d->window_height));
    //Store the firstSheet when it isn't the default
    //For example, when "the first sheet 0 is hidden", the first sheet will be 1
    if (d->firstsheet > 0)
        writer.writeAttribute(QStringLiteral("firstSheet"), QString::number(d->firstsheet + 1));
    //Store the activeTab when it isn't the first sheet
    if (d->activesheetIndex > 0)
        writer.writeAttribute(QStringLiteral("activeTab"), QString::number(d->activesheetIndex));
    writer.writeEndElement();//bookViews

    writer.writeStartElement(QStringLiteral("sheets"));
    int worksheetIndex = 0;
    int chartsheetIndex = 0;
    for (int i=0; i<d->sheets.size(); ++i) {
        QSharedPointer<AbstractSheet> sheet = d->sheets[i];
        writer.writeEmptyElement(QStringLiteral("sheet"));
        writer.writeAttribute(QStringLiteral("name"), sheet->name());
        writer.writeAttribute(QStringLiteral("sheetId"), QString::number(sheet->id()));
        if (sheet->visibility() == AbstractSheet::Visibility::Hidden)
            writer.writeAttribute(QStringLiteral("state"), QStringLiteral("hidden"));
        else if (sheet->visibility() == AbstractSheet::Visibility::VeryHidden)
            writer.writeAttribute(QStringLiteral("state"), QStringLiteral("veryHidden"));

        if (sheet->type() == AbstractSheet::Type::Worksheet)
            d->relationships->addDocumentRelationship(QStringLiteral("/worksheet"), QStringLiteral("worksheets/sheet%1.xml").arg(++worksheetIndex));
        else if (sheet->type() == AbstractSheet::Type::Chartsheet)
            d->relationships->addDocumentRelationship(QStringLiteral("/chartsheet"), QStringLiteral("chartsheets/sheet%1.xml").arg(++chartsheetIndex));

        writer.writeAttribute(QStringLiteral("r:id"), QStringLiteral("rId%1").arg(d->relationships->count()));
    }
    writer.writeEndElement();//sheets

    if (d->externalLinks.size() > 0) {
        writer.writeStartElement(QStringLiteral("externalReferences"));
        for (int i=0; i<d->externalLinks.size(); ++i) {
            writer.writeEmptyElement(QStringLiteral("externalReference"));
            d->relationships->addDocumentRelationship(QStringLiteral("/externalLink"), QStringLiteral("externalLinks/externalLink%1.xml").arg(i+1));
            writer.writeAttribute(QStringLiteral("r:id"), QStringLiteral("rId%1").arg(d->relationships->count()));
        }
        writer.writeEndElement();//externalReferences
    }

    if (!d->definedNamesList.isEmpty()) {
        writer.writeStartElement(QStringLiteral("definedNames"));
        for (const XlsxDefineNameData &data : qAsConst(d->definedNamesList)) {
            writer.writeStartElement(QStringLiteral("definedName"));
            writer.writeAttribute(QStringLiteral("name"), data.name);
            if (!data.comment.isEmpty())
                writer.writeAttribute(QStringLiteral("comment"), data.comment);
            if (data.sheetId != -1) {
                //find the local index of the sheet.
                for (int i=0; i<d->sheets.size(); ++i) {
                    if (d->sheets[i]->id() == data.sheetId) {
                        writer.writeAttribute(QStringLiteral("localSheetId"), QString::number(i));
                        break;
                    }
                }
            }
            writer.writeCharacters(data.formula);
            writer.writeEndElement();//definedName
        }
        writer.writeEndElement();//definedNames
    }

    writer.writeStartElement(QStringLiteral("calcPr"));
    writer.writeAttribute(QStringLiteral("calcId"), QStringLiteral("124519"));
    writer.writeEndElement(); //calcPr

    writer.writeEndElement();//workbook
    writer.writeEndDocument();

    d->relationships->addDocumentRelationship(QStringLiteral("/theme"), QStringLiteral("theme/theme1.xml"));
    d->relationships->addDocumentRelationship(QStringLiteral("/styles"), QStringLiteral("styles.xml"));
    if (!sharedStrings()->isEmpty())
        d->relationships->addDocumentRelationship(QStringLiteral("/sharedStrings"), QStringLiteral("sharedStrings.xml"));
}

bool Workbook::loadFromXmlFile(QIODevice *device)
{
    Q_D(Workbook);

    QXmlStreamReader reader(device);
    while (!reader.atEnd()) {
         QXmlStreamReader::TokenType token = reader.readNext();
         if (token == QXmlStreamReader::StartElement) {
             if (reader.name() == QLatin1String("sheet")) {
                 QXmlStreamAttributes attributes = reader.attributes();

                 const auto& name = attributes.value(QLatin1String("name")).toString();

                 int sheetId = attributes.value(QLatin1String("sheetId")).toInt();

                 const auto& rId = attributes.value(QLatin1String("r:id")).toString();

                 const auto& stateString = attributes.value(QLatin1String("state"));

                 auto state = AbstractSheet::Visibility::Visible;
                 if (stateString == QLatin1String("hidden"))
                     state = AbstractSheet::Visibility::Hidden;
                 else if (stateString == QLatin1String("veryHidden"))
                     state = AbstractSheet::Visibility::VeryHidden;

                 XlsxRelationship relationship = d->relationships->getRelationshipById(rId);

                 auto type = AbstractSheet::Type::Worksheet;
                 if (relationship.type.endsWith(QLatin1String("/worksheet")))
                 {
                     type = AbstractSheet::Type::Worksheet;
                 }
                 else if (relationship.type.endsWith(QLatin1String("/chartsheet")))
                 {
                     type = AbstractSheet::Type::Chartsheet;
                 }
                 else if (relationship.type.endsWith(QLatin1String("/dialogsheet")))
                 {
                     type = AbstractSheet::Type::Dialogsheet;
                 }
                 else if (relationship.type.endsWith(QLatin1String("/xlMacrosheet")))
                 {
                     type = AbstractSheet::Type::Macrosheet;
                 }
                 else
                 {
                     qWarning() << "unknown sheet type : " << relationship.type ;
                 }

                 AbstractSheet *sheet = addSheet(name, sheetId, type);
                 sheet->setVisibility(state);
                 QString strFilePath = filePath();

                 // const QString fullPath = QDir::cleanPath(splitPath(strFilePath).constFirst() + QLatin1String("/") + relationship.target);
                 const auto parts = splitPath(strFilePath);
                 QString fullPath = QDir::cleanPath(parts.first() + QLatin1String("/") + relationship.target);

                 sheet->setFilePath(fullPath);
             }
             else if (reader.name() == QLatin1String("workbookPr"))
             {
                QXmlStreamAttributes attrs = reader.attributes();
                if (attrs.hasAttribute(QLatin1String("date1904")))
                    d->date1904 = true;
             }
             else if (reader.name() == QLatin1String("bookviews"))
             {
                while (!(reader.name() == QLatin1String("bookviews") &&
                         reader.tokenType() == QXmlStreamReader::EndElement))
                {
                    reader.readNextStartElement();
                    if (reader.tokenType() == QXmlStreamReader::StartElement)
                    {
                        if (reader.name() == QLatin1String("workbookView"))
                        {
                            QXmlStreamAttributes attrs = reader.attributes();
                            if (attrs.hasAttribute(QLatin1String("xWindow")))
                                d->x_window = attrs.value(QLatin1String("xWindow")).toInt();
                            if (attrs.hasAttribute(QLatin1String("yWindow")))
                                d->y_window = attrs.value(QLatin1String("yWindow")).toInt();
                            if (attrs.hasAttribute(QLatin1String("windowWidth")))
                                d->window_width = attrs.value(QLatin1String("windowWidth")).toInt();
                            if (attrs.hasAttribute(QLatin1String("windowHeight")))
                                d->window_height = attrs.value(QLatin1String("windowHeight")).toInt();
                            if (attrs.hasAttribute(QLatin1String("firstSheet")))
                                d->firstsheet = attrs.value(QLatin1String("firstSheet")).toInt();
                            if (attrs.hasAttribute(QLatin1String("activeTab")))
                                d->activesheetIndex = attrs.value(QLatin1String("activeTab")).toInt();
                        }
                    }
                }
             }
             else if (reader.name() == QLatin1String("externalReference"))
             {
                 QXmlStreamAttributes attributes = reader.attributes();
                 const QString rId = attributes.value(QLatin1String("r:id")).toString();
                 XlsxRelationship relationship = d->relationships->getRelationshipById(rId);

                 QSharedPointer<SimpleOOXmlFile> link(new SimpleOOXmlFile(F_LoadFromExists));

                 const auto parts = splitPath(filePath());
                 QString fullPath = QDir::cleanPath(parts.first() + QLatin1String("/") + relationship.target);

                 link->setFilePath(fullPath);
                 d->externalLinks.append(link);
             } else if (reader.name() == QLatin1String("definedName")) {
                 QXmlStreamAttributes attrs = reader.attributes();
                 XlsxDefineNameData data;

                 data.name = attrs.value(QLatin1String("name")).toString();
                 if (attrs.hasAttribute(QLatin1String("comment")))
                     data.comment = attrs.value(QLatin1String("comment")).toString();
                 if (attrs.hasAttribute(QLatin1String("localSheetId"))) {
                     int localId = attrs.value(QLatin1String("localSheetId")).toInt();
                     int sheetId = d->sheets.at(localId)->id();
                     data.sheetId = sheetId;
                 }
                 data.formula = reader.readElementText();
                 d->definedNamesList.append(data);
             }
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
        for (int i=0; i<d->mediaFiles.size(); ++i) {
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
    for (auto &media: d->mediaFiles) {
        if (auto mf = media.lock())  mf->setIndex(index++);
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

}
