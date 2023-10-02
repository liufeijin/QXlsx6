// xlsxworksheet_p.h

#ifndef XLSXWORKSHEET_P_H
#define XLSXWORKSHEET_P_H

#include <QtGlobal>
#include <QObject>
#include <QString>
#include <QVector>
#include <QImage>
#include <QSharedPointer>

#include <QRegularExpression>

#include "xlsxworksheet.h"
#include "xlsxabstractsheet_p.h"
#include "xlsxcell.h"
#include "xlsxdatavalidation.h"
#include "xlsxconditionalformatting.h"
#include "xlsxcellformula.h"
#include "xlsxautofilter.h"

class QXmlStreamWriter;
class QXmlStreamReader;

namespace QXlsx {

const int XLSX_ROW_MAX = 1048576;
const int XLSX_COLUMN_MAX = 16384;
const int XLSX_STRING_MAX = 32767;
constexpr const double XLSX_DEFAULT_ROW_HEIGHT = 14.4;

class SharedStrings;

struct XlsxHyperlinkData
{
    enum LinkType
    {
        External,
        Internal
    };

    XlsxHyperlinkData(LinkType linkType=External, const QString &target=QString(), const QString &location=QString()
            , const QString &display=QString(), const QString &tip=QString())
        :linkType(linkType), target(target), location(location), display(display), tooltip(tip)
    {

    }

    LinkType linkType;
    QString target; //For External link
    QString location;
    QString display;
    QString tooltip;
};

// ECMA-376 Part1 18.3.1.81
struct SheetFormatProperties
{
    int baseColWidth = 8; //we make it non-optional to have some base value we can use elsewhere.
    std::optional<bool> customHeight; // = false;
    std::optional<double> defaultColWidth; // = 8.43f;
    double defaultRowHeight = 15; //or 14.4, required according to ECMA376
    std::optional<int> outlineLevelCol; // = 0;
    std::optional<int> outlineLevelRow; // = 0;
    std::optional<bool> thickBottom; // = false;
    std::optional<bool> thickTop; //= false;
    std::optional<bool> zeroHeight; // = false;
    //returns calculated default column width
    double defaultColumnWidth() const;
    bool isValid() const;
    void read(QXmlStreamReader &reader);
    void write(QXmlStreamWriter &writer, const QLatin1String &name) const;
};

struct XlsxRowInfo
{
    std::optional<double> height;
    Format format;
    std::optional<bool> hidden;
    std::optional<int> outlineLevel;
    std::optional<bool> collapsed;

    bool isValid() const
    {
        if (height.has_value()) return true;
        if (format.isValid()) return true;
        if (hidden.has_value()) return true;
        if (outlineLevel.has_value()) return true;
        if (collapsed.has_value()) return true;
        return false;
    }
};

//TODO: convert to explicitly shareable to reduce memory
struct XlsxColumnInfo
{
    bool operator==(const XlsxColumnInfo& other) const {
        if (width != other.width) return false;
        if (format != other.format) return false;
        if (outlineLevel != other.outlineLevel) return false;
        if (hidden != other.hidden) return false;
        if (collapsed != other.collapsed) return false;
        return true;
    }
    bool operator!=(const XlsxColumnInfo& other) const {
        return !operator==(other);
    }
    std::optional<double> width;
    Format format;
    std::optional<int> outlineLevel;
    std::optional<bool> hidden;
    std::optional<bool> collapsed;
};

class WorksheetPrivate : public AbstractSheetPrivate
{
    Q_DECLARE_PUBLIC(Worksheet)

public:
    WorksheetPrivate(Worksheet *p, Worksheet::CreateFlag flag);
    ~WorksheetPrivate();

public:
    bool rowValid(int row) const;
    bool columnValid(int column) const;
    bool addRowToDimensions(int row);
    bool addColumnToDimensions(int column);
    Format cellFormat(int row, int col) const;
    QString generateDimensionString() const;
    QMap<int, QString> calculateSpans() const;
    void validateDimension();

    void saveXmlSheetData(QXmlStreamWriter &writer) const;
    void saveXmlCellData(QXmlStreamWriter &writer, int row, int col, std::shared_ptr<Cell> cell) const;
    void saveXmlMergeCells(QXmlStreamWriter &writer) const;
    void saveXmlHyperlinks(QXmlStreamWriter &writer) const;
    void saveXmlDrawings(QXmlStreamWriter &writer) const;
    void saveXmlDataValidations(QXmlStreamWriter &writer) const;

    void loadXmlSheetData(QXmlStreamReader &reader);
    void loadXmlColumnsInfo(QXmlStreamReader &reader);
    void loadXmlMergeCells(QXmlStreamReader &reader);
    void loadXmlDataValidations(QXmlStreamReader &reader);
    void loadXmlSheetViews(QXmlStreamReader &reader);
    void loadXmlHyperlinks(QXmlStreamReader &reader);
    void loadXmlCell(QXmlStreamReader &reader);

    bool isColumnRangeValid(int colFirst, int colLast) const;
    QList<QPair<int,int>> getIntervals() const;

    SharedStrings *sharedStrings() const;

public:
    QMap<int, QMap<int, std::shared_ptr<Cell> > > cellTable;

    QMap<int, QMap<int, QString> > comments;
    QMap<int, QMap<int, QSharedPointer<XlsxHyperlinkData> > > urlTable;
    QList<CellRange> merges;
    QMap<int, QSharedPointer<XlsxRowInfo> > rowsInfo;
    QMap<int, XlsxColumnInfo> colsInfo;

    std::optional<bool> disableValidationPrompts; //default = false;
    std::optional<int> dataValidationXWindow;
    std::optional<int> dataValidationYWindow;
    QList<DataValidation> dataValidationsList;
    QList<ConditionalFormatting> conditionalFormattingList;

    QMap<int, CellFormula> sharedFormulaMap; // shared formula map

    CellRange dimension;

    SheetFormatProperties sheetFormatProperties;
    AutoFilter autofilter;

    QRegularExpression urlPattern {QStringLiteral("^([fh]tt?ps?://)|(mailto:)|(file://)")};
private:

    static double calculateColWidth(int characters);
};

}
#endif // XLSXWORKSHEET_P_H
