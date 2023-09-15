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

class QXmlStreamWriter;
class QXmlStreamReader;

namespace QXlsx {

const int XLSX_ROW_MAX = 1048576;
const int XLSX_COLUMN_MAX = 16384;
const int XLSX_STRING_MAX = 32767;

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
struct XlsxSheetFormatProps
{
    XlsxSheetFormatProps(int baseColWidth = 8,
                         bool customHeight = false,
                         double defaultColWidth = 8.430f, // https://learn.microsoft.com/en-us/office/troubleshoot/excel/determine-column-widths
                         double defaultRowHeight = 15,
                         quint8 outlineLevelCol = 0,
                         quint8 outlineLevelRow = 0,
                         bool thickBottom = false,
                         bool thickTop = false,
                         bool zeroHeight = false) :
        baseColWidth(baseColWidth),
        customHeight(customHeight),
        defaultColWidth(defaultColWidth),
        defaultRowHeight(defaultRowHeight),
        outlineLevelCol(outlineLevelCol),
        outlineLevelRow(outlineLevelRow),
        thickBottom(thickBottom),
        thickTop(thickTop),
        zeroHeight(zeroHeight) {
    }

    int baseColWidth;
    bool customHeight;
    double defaultColWidth = 0.0;
    double defaultRowHeight;
    int outlineLevelCol;
    int outlineLevelRow;
    bool thickBottom;
    bool thickTop;
    bool zeroHeight;
};

struct SheetProperties
{
    Color tabColor;
    std::optional<bool> applyStyles;
    std::optional<bool> summaryBelow;
    std::optional<bool> summaryRight;
    std::optional<bool> showOutlineSymbols;
    std::optional<bool> autoPageBreaks;
    std::optional<bool> fitToPage;
    std::optional<bool> syncHorizontal;
    std::optional<bool> syncVertical;
    QString syncRef;
    std::optional<bool> transitionEvaluation;
    std::optional<bool> transitionEntry;
    std::optional<bool> published;
    QString codeName;
    std::optional<bool> filterMode;
    std::optional<bool> enableFormatConditionsCalculation;
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
    void calculateSpans() const;
    void validateDimension();

    void saveXmlSheetData(QXmlStreamWriter &writer) const;
    void saveXmlCellData(QXmlStreamWriter &writer, int row, int col, std::shared_ptr<Cell> cell) const;
    void saveXmlMergeCells(QXmlStreamWriter &writer) const;
    void saveXmlHyperlinks(QXmlStreamWriter &writer) const;
    void saveXmlDrawings(QXmlStreamWriter &writer) const;
    void saveXmlDataValidations(QXmlStreamWriter &writer) const;

    int rowPixelsSize(int row) const;
    int colPixelsSize(int col) const;

    void loadXmlSheetData(QXmlStreamReader &reader);
    void loadXmlColumnsInfo(QXmlStreamReader &reader);
    void loadXmlMergeCells(QXmlStreamReader &reader);
    void loadXmlDataValidations(QXmlStreamReader &reader);
    void loadXmlSheetFormatProps(QXmlStreamReader &reader);
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

    QList<DataValidation> dataValidationsList;
    QList<ConditionalFormatting> conditionalFormattingList;

    QMap<int, CellFormula> sharedFormulaMap; // shared formula map

    CellRange dimension;
//    int previous_row = 0;

    mutable QMap<int, QString> rowSpans;
    QMap<int, double> rowSizes;
    QMap<int, double> colSizes;

    int outlineRowLevel = 0;
    int outlineColLevel = 0;

    double defaultRowHeight = 14.4;
    bool defaultRowZeroed = false;

    XlsxSheetFormatProps sheetFormatProps;

    QList<SheetView> sheetViews;

    QRegularExpression urlPattern {QStringLiteral("^([fh]tt?ps?://)|(mailto:)|(file://)")};

private:

    static double calculateColWidth(int characters);
};

}
#endif // XLSXWORKSHEET_P_H
