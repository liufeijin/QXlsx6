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
    double defaultColWidth;
    double defaultRowHeight;
    quint8 outlineLevelCol;
    quint8 outlineLevelRow;
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
    XlsxRowInfo(double height=0, const Format &format=Format(), bool hidden=false) :
        customHeight(false), height(height), format(format), hidden(hidden), outlineLevel(0)
      , collapsed(false)
    {

    }

    bool customHeight;
    double height;
    Format format;
    bool hidden;
    int outlineLevel;
    bool collapsed;
};

struct XlsxColumnInfo
{
    XlsxColumnInfo( int firstColumn, // = 0,
                    int lastColumn, // = 1,
                    double width = 0,
                    const Format &format = Format(),
                    bool hidden = false)
        : width(width),
          format(format),
          firstColumn(firstColumn),
          lastColumn(lastColumn),
          outlineLevel(0),
          customWidth(false),
          hidden(hidden),
          collapsed(false)
    {

    }

    std::optional<double> width;
    Format format;
    int firstColumn;
    int lastColumn;
    int outlineLevel;
    bool customWidth;
    bool hidden;
    bool collapsed;
};

class WorksheetPrivate : public AbstractSheetPrivate
{
    Q_DECLARE_PUBLIC(Worksheet)

public:
    WorksheetPrivate(Worksheet *p, Worksheet::CreateFlag flag);
    ~WorksheetPrivate();

public:
    int checkDimensions(int row, int col, bool ignore_row=false, bool ignore_col=false);
    Format cellFormat(int row, int col) const;
    QString generateDimensionString() const;
    void calculateSpans() const;
    void splitColsInfo(int colFirst, int colLast);
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

    //NOTE: this method inserts missing rows into rowsInfo, as it is used exclusively in
    //setRowsHeight, setRowsFormat, setRowsHidden methods.
    //I guess if we need to change format of rows that are yet empty, this is the best solution.
    QList<QSharedPointer<XlsxRowInfo> > getRowInfoList(int rowFirst, int rowLast);
    QList<QSharedPointer<XlsxColumnInfo> > getColumnInfoList(int colFirst, int colLast) const;
    QList<int> getColumnIndexes(int colFirst, int colLast);
    bool isColumnRangeValid(int colFirst, int colLast);

    SharedStrings *sharedStrings() const;

public:
    QMap<int, QMap<int, std::shared_ptr<Cell> > > cellTable;

    QMap<int, QMap<int, QString> > comments;
    QMap<int, QMap<int, QSharedPointer<XlsxHyperlinkData> > > urlTable;
    QList<CellRange> merges;
    QMap<int, QSharedPointer<XlsxRowInfo> > rowsInfo;
    QMap<int, QSharedPointer<XlsxColumnInfo> > colsInfo;
    QMap<int, QSharedPointer<XlsxColumnInfo> > colsInfoHelper;

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

    int defaultRowHeight = 15;
    bool defaultRowZeroed = false;

    XlsxSheetFormatProps sheetFormatProps;

    bool windowProtection = false;
    bool showFormulas = false;
    bool showGridLines = true;
    bool showRowColHeaders = true;
    bool showZeros = true;
    bool rightToLeft = false;
    bool tabSelected = false;
    bool showRuler = false;
    bool showOutlineSymbols = true;
    bool showWhiteSpace = true;

    QRegularExpression urlPattern {QStringLiteral("^([fh]tt?ps?://)|(mailto:)|(file://)")};

private:

    static double calculateColWidth(int characters);
};

}
#endif // XLSXWORKSHEET_P_H
