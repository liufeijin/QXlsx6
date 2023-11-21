// xlsxworksheet.cpp

#include <QtGlobal>
#include <QVariant>
#include <QDateTime>
#include <QDate>
#include <QTime>
#include <QPoint>
#include <QFile>
#include <QUrl>
#include <QDebug>
#include <QBuffer>
#include <QXmlStreamWriter>
#include <QXmlStreamReader>
#include <QTextDocument>
#include <QDir>
#include <QMapIterator>
#include <QMap>
#include <QFontMetricsF>

#include <cmath>

#include "xlsxrichstring.h"
#include "xlsxcellreference.h"
#include "xlsxworksheet.h"
#include "xlsxworksheet_p.h"
#include "xlsxworkbook.h"
#include "xlsxformat.h"
#include "xlsxutility_p.h"
#include "xlsxsharedstrings_p.h"
#include "xlsxdrawing_p.h"
#include "xlsxstyles_p.h"
#include "xlsxcell.h"
#include "xlsxcellrange.h"
#include "xlsxconditionalformatting_p.h"
#include "xlsxdrawinganchor_p.h"
#include "xlsxchart.h"
#include "xlsxcellformula.h"
#include "xlsxmain.h"

namespace QXlsx {

WorksheetPrivate::WorksheetPrivate(Worksheet *p, Worksheet::CreateFlag flag) : AbstractSheetPrivate(p, flag)

{
}

WorksheetPrivate::~WorksheetPrivate()
{
}

bool WorksheetPrivate::rowValid(int row) const
{
    return (row < XLSX_ROW_MAX && row > 0);
}

bool WorksheetPrivate::columnValid(int column) const
{
    return (column < XLSX_COLUMN_MAX && column > 0);
}

bool WorksheetPrivate::addRowToDimensions(int row)
{
    if (!rowValid(row)) return false;
    if (row < dimension.firstRow() || dimension.firstRow() == -1) dimension.setFirstRow(row);
    if (row > dimension.lastRow()) dimension.setLastRow(row);
    return true;
}

bool WorksheetPrivate::addColumnToDimensions(int column)
{
    if (!columnValid(column)) return false;
    if (column < dimension.firstColumn() || dimension.firstColumn() == -1) dimension.setFirstColumn(column);
    if (column > dimension.lastColumn()) dimension.setLastColumn(column);
    return true;
}

/*
  Calculate the "spans" attribute of the <row> tag. This is an
  XLSX optimisation and isn't strictly required. However, it
  makes comparing files easier. The span is the same for each
  block of 16 rows.
 */
QMap<int, QString> WorksheetPrivate::calculateSpans() const
{
    QMap<int, QString> rowSpans;
    int span_min = XLSX_COLUMN_MAX+1;
    int span_max = -1;

    for (int row_num = dimension.firstRow(); row_num <= dimension.lastRow(); row_num++) {
        auto it = cellTable.constFind(row_num);
        if (it != cellTable.constEnd()) {
            for (int col_num = dimension.firstColumn(); col_num <= dimension.lastColumn(); col_num++) {
                if (it->contains(col_num)) {
                    if (span_max == -1) {
                        span_min = col_num;
                        span_max = col_num;
                    } else {
                        if (col_num < span_min)
                            span_min = col_num;
                        else if (col_num > span_max)
                            span_max = col_num;
                    }
                }
            }
        }
        auto cIt = comments.constFind(row_num);
        if (cIt != comments.constEnd()) {
            for (int col_num = dimension.firstColumn(); col_num <= dimension.lastColumn(); col_num++) {
                if (cIt->contains(col_num)) {
                    if (span_max == -1) {
                        span_min = col_num;
                        span_max = col_num;
                    } else {
                        if (col_num < span_min)
                            span_min = col_num;
                        else if (col_num > span_max)
                            span_max = col_num;
                    }
                }
            }
        }

        if (row_num%16 == 0 || row_num == dimension.lastRow()) {
            if (span_max != -1) {
                rowSpans[row_num / 16] = QString("%1:%2").arg(span_min).arg(span_max);
                span_min = XLSX_COLUMN_MAX+1;
                span_max = -1;
            }
        }
    }
    return rowSpans;
}


QString WorksheetPrivate::generateDimensionString() const
{
    //invalid CellRange returns empty string, so we need additional check
    if (!dimension.isValid())
        return QLatin1String("A1");
    else
        return dimension.toString();
}

/*!
 * \internal
 */
Worksheet::Worksheet(const QString &name, int id, Workbook *workbook, CreateFlag flag)
    :AbstractSheet(name, id, workbook, new WorksheetPrivate(this, flag))
{
    if (!workbook) //For unit test propose only. Ignore the memery leak.
        d_func()->workbook = new Workbook(flag);
}

/*!
 * \internal
 *
 * Make a copy of this sheet.
 */

Worksheet *Worksheet::copy(const QString &distName, int distId) const
{
    Q_D(const Worksheet);
    Worksheet *sheet = new Worksheet(distName, distId, d->workbook, F_NewFromScratch);
    WorksheetPrivate *sheet_d = sheet->d_func();

    sheet_d->dimension = d->dimension;
    if (d->drawing) {
        sheet_d->drawing = std::make_shared<Drawing>(sheet, F_NewFromScratch);
        for (auto anchor: qAsConst(d->drawing->anchors)) {
            anchor->copyTo(sheet_d->drawing.get());
        }
    }
    sheet_d->sheetState = d->sheetState;
    sheet_d->type = d->type;
    sheet_d->headerFooter = d->headerFooter;
    sheet_d->pageMargins = d->pageMargins;
    sheet_d->pageSetup = d->pageSetup;
    sheet_d->pictureFile = d->pictureFile;
    sheet_d->sheetProtection = d->sheetProtection;

    QMapIterator<int, QMap<int, std::shared_ptr<Cell> > > it(d->cellTable);
    while (it.hasNext())
    {
        it.next();
        int row = it.key();
        QMapIterator<int, std::shared_ptr<Cell> > it2(it.value());
        while (it2.hasNext())
        {
            it2.next();
            int col = it2.key();

            auto cell = std::make_shared<Cell>(it2.value().get(), sheet);
//            cell->d_ptr->parent = sheet;

//            if (cell->type() == Cell::Type::SharedString)
//                d->workbook->sharedStrings()->addSharedString(cell->d_ptr->richString);

            sheet_d->cellTable[row][col] = cell;
        }
    }

    sheet_d->merges = d->merges;
    sheet_d->rowsInfo = d->rowsInfo;
    sheet_d->colsInfo = d->colsInfo;
    sheet_d->dataValidationsList = d->dataValidationsList;
    sheet_d->conditionalFormattingList = d->conditionalFormattingList;
    sheet_d->sheetViews = d->sheetViews;

    return sheet;
}

Worksheet::~Worksheet()
{
}

std::optional<bool> Worksheet::isWindowProtected() const
{
    Q_D(const Worksheet);
    if (d->sheetViews.isEmpty()) return {};
    return d->sheetViews.last().windowProtection;
}

void Worksheet::setWindowProtected(bool protect)
{
    Q_D(Worksheet);
    if (d->sheetViews.isEmpty()) d->sheetViews << SheetView();
    d->sheetViews.last().windowProtection = protect;
}

std::optional<bool> Worksheet::isFormulasVisible() const
{
    Q_D(const Worksheet);
    if (d->sheetViews.isEmpty()) return {};
    return d->sheetViews.last().showFormulas;
}

void Worksheet::setFormulasVisible(bool visible)
{
    Q_D(Worksheet);
    if (d->sheetViews.isEmpty()) d->sheetViews << SheetView();
    d->sheetViews.last().showFormulas = visible;
}

std::optional<bool> Worksheet::isGridLinesVisible() const
{
    Q_D(const Worksheet);
    if (d->sheetViews.isEmpty()) return {};
    return d->sheetViews.last().showGridLines;
}

void Worksheet::setGridLinesVisible(bool visible)
{
    Q_D(Worksheet);
    if (d->sheetViews.isEmpty()) d->sheetViews << SheetView();
    d->sheetViews.last().showGridLines = visible;
}

std::optional<bool> Worksheet::isRowColumnHeadersVisible() const
{
    Q_D(const Worksheet);
    if (d->sheetViews.isEmpty()) return {};
    return d->sheetViews.last().showRowColHeaders;
}

void Worksheet::setRowColumnHeadersVisible(bool visible)
{
    Q_D(Worksheet);
    if (d->sheetViews.isEmpty()) d->sheetViews << SheetView();
    d->sheetViews.last().showRowColHeaders = visible;
}

std::optional<bool> Worksheet::isRightToLeft() const
{
    Q_D(const Worksheet);
    if (d->sheetViews.isEmpty()) return {};
    return d->sheetViews.last().rightToLeft;
}

void Worksheet::setRightToLeft(bool enable)
{
    Q_D(Worksheet);
    if (d->sheetViews.isEmpty()) d->sheetViews << SheetView();
    d->sheetViews.last().rightToLeft = enable;
}

std::optional<bool> Worksheet::isZerosVisible() const
{
    Q_D(const Worksheet);
    if (d->sheetViews.isEmpty()) return {};
    return d->sheetViews.last().showZeros;
}

void Worksheet::setZerosVisible(bool visible)
{
    Q_D(Worksheet);
    if (d->sheetViews.isEmpty()) d->sheetViews << SheetView();
    d->sheetViews.last().showZeros = visible;
}

std::optional<bool> Worksheet::isRulerVisible() const
{
    Q_D(const Worksheet);
    if (d->sheetViews.isEmpty()) return {};
    return d->sheetViews.last().showRuler;
}

void Worksheet::setRulerVisible(bool visible)
{
    Q_D(Worksheet);
    if (d->sheetViews.isEmpty()) d->sheetViews << SheetView();
    d->sheetViews.last().showRuler = visible;
}

std::optional<bool> Worksheet::isOutlineSymbolsVisible() const
{
    Q_D(const Worksheet);
    if (d->sheetViews.isEmpty()) return {};
    return d->sheetViews.last().showOutlineSymbols;
}

void Worksheet::setOutlineSymbolsVisible(bool visible)
{
    Q_D(Worksheet);
    if (d->sheetViews.isEmpty()) d->sheetViews << SheetView();
    d->sheetViews.last().showOutlineSymbols = visible;
}

std::optional<bool> Worksheet::isPageMarginsVisible() const
{
    Q_D(const Worksheet);
    if (d->sheetViews.isEmpty()) return {};
    return d->sheetViews.last().showWhiteSpace;
}

void Worksheet::setPageMarginsVisible(bool visible)
{
    Q_D(Worksheet);
    if (d->sheetViews.isEmpty()) d->sheetViews << SheetView();
    d->sheetViews.last().showWhiteSpace = visible;
}

std::optional<bool> Worksheet::isDefaultGridColorUsed() const
{
    Q_D(const Worksheet);
    if (d->sheetViews.isEmpty()) return {};
    return d->sheetViews.last().defaultGridColor;
}

void Worksheet::setDefaultGridColorUsed(bool value)
{
    Q_D(Worksheet);
    if (d->sheetViews.isEmpty()) d->sheetViews << SheetView();
    d->sheetViews.last().defaultGridColor = value;
}

std::optional<SheetView::Type> Worksheet::viewType() const
{
    Q_D(const Worksheet);
    if (d->sheetViews.isEmpty()) return {};
    return d->sheetViews.last().type;
}

void Worksheet::setViewType(SheetView::Type type)
{
    Q_D(Worksheet);
    if (d->sheetViews.isEmpty()) d->sheetViews << SheetView();
    d->sheetViews.last().type = type;
}

CellReference Worksheet::viewTopLeftCell() const
{
    Q_D(const Worksheet);
    if (d->sheetViews.isEmpty()) return {};
    return d->sheetViews.last().topLeftCell;
}

void Worksheet::setViewTopLeftCell(const CellReference &ref)
{
    Q_D(Worksheet);
    if (d->sheetViews.isEmpty()) d->sheetViews << SheetView();
    d->sheetViews.last().topLeftCell = ref;
}

std::optional<int> Worksheet::viewColorIndex() const
{
    Q_D(const Worksheet);
    if (d->sheetViews.isEmpty()) return {};
    return d->sheetViews.last().colorId;
}

void Worksheet::setViewColorIndex(int index)
{
    Q_D(Worksheet);
    if (d->sheetViews.isEmpty()) d->sheetViews << SheetView();
    if (index >=0 && index <= 64) d->sheetViews.last().colorId = index;
}



CellReference Worksheet::activeCell() const
{
    Q_D(const Worksheet);
    if (d->sheetViews.isEmpty()) return {};
    return d->sheetViews.last().selection.activeCell;
}

void Worksheet::setActiveCell(const CellReference &activeCell)
{
    Q_D(Worksheet);
    if (d->sheetViews.isEmpty()) d->sheetViews << SheetView();
    d->sheetViews.last().selection.activeCell = activeCell;
}

QList<CellRange> Worksheet::selectedRanges() const
{
    Q_D(const Worksheet);
    if (d->sheetViews.isEmpty()) return {};
    return d->sheetViews.last().selection.selectedRanges;
}

bool Worksheet::addSelection(const CellRange &range)
{
    Q_D(Worksheet);
    if (!range.isValid()) return false;

    if (d->sheetViews.isEmpty()) d->sheetViews << SheetView();
    if (d->sheetViews.last().selection.selectedRanges.contains(range)) return false;
    d->sheetViews.last().selection.selectedRanges << range;
    return true;
}

bool Worksheet::removeSelection(const CellRange &range)
{
    Q_D(Worksheet);
    if (!range.isValid()) return false;

    if (d->sheetViews.isEmpty()) d->sheetViews << SheetView();
    return d->sheetViews.last().selection.selectedRanges.removeOne(range);
}

void Worksheet::clearSelection()
{
    Q_D(Worksheet);
    if (d->sheetViews.isEmpty()) d->sheetViews << SheetView();
    d->sheetViews.last().selection.selectedRanges.clear();
    d->sheetViews.last().selection.activeCell = CellReference();
    d->sheetViews.last().selection.activeCellId.reset();
}

Selection Worksheet::selection() const
{
    Q_D(const Worksheet);
    if (d->sheetViews.isEmpty()) return {};
    return d->sheetViews.last().selection;
}

Selection &Worksheet::selection()
{
    Q_D(Worksheet);
    if (d->sheetViews.isEmpty()) d->sheetViews << SheetView();
    return d->sheetViews.last().selection;
}

void Worksheet::setSelection(const Selection &selection)
{
    Q_D(Worksheet);
    if (d->sheetViews.isEmpty()) d->sheetViews << SheetView();
    d->sheetViews.last().selection = selection;
}

void Worksheet::setPrintScale(int scale)
{
    Q_D(Worksheet);
    d->pageSetup.setScale(scale);
}

std::optional<int> Worksheet::printScale() const
{
    Q_D(const Worksheet);
    return d->pageSetup.scale();
}

std::optional<PageSetup::PageOrder> Worksheet::pageOrder() const
{
    Q_D(const Worksheet);
    return d->pageSetup.pageOrder();
}

void Worksheet::setPageOrder(PageSetup::PageOrder pageOrder)
{
    Q_D(Worksheet);
    d->pageSetup.setPageOrder(pageOrder);
}

std::optional<int> Worksheet::fitToWidth() const
{
    Q_D(const Worksheet);
    return d->pageSetup.fitToWidth();
}

void Worksheet::setFitToWidth(int pages)
{
    Q_D(Worksheet);
    if (pages > 0) d->pageSetup.setFitToWidth(pages);
}

std::optional<int> Worksheet::fitToHeight() const
{
    Q_D(const Worksheet);
    return d->pageSetup.fitToHeight();
}

void Worksheet::setFitToHeight(int pages)
{
    Q_D(Worksheet);
    if (pages > 0) d->pageSetup.setFitToHeight(pages);
}

std::optional<PageSetup::PrintError> Worksheet::printErrors() const
{
    Q_D(const Worksheet);
    return d->pageSetup.printErrors();
}

void Worksheet::setPrintErrors(PageSetup::PrintError mode)
{
    Q_D(Worksheet);
    d->pageSetup.setPrintErrors(mode);
}

std::optional<PageSetup::CellComments> Worksheet::printCellComments() const
{
    Q_D(const Worksheet);
    return d->pageSetup.printCellComments();
}

void Worksheet::setPrintCellComments(PageSetup::CellComments mode)
{
    Q_D(Worksheet);
    d->pageSetup.setPrintCellComments(mode);
}

std::optional<bool> Worksheet::printGridLines() const
{
    Q_D(const Worksheet);
    return d->pageSetup.printGridLines();
}

void Worksheet::setPrintGridLines(bool printGridLines)
{
    Q_D(Worksheet);
    d->pageSetup.setPrintGridLines(printGridLines);
}

std::optional<bool> Worksheet::printHeadings() const
{
    Q_D(const Worksheet);
    return d->pageSetup.printHeadings();
}

void Worksheet::setPrintHeadings(bool printHeadings)
{
    Q_D(Worksheet);
    d->pageSetup.setPrintHeadings(printHeadings);
}

std::optional<bool> Worksheet::printHorizontalCentered() const
{
    Q_D(const Worksheet);
    return d->pageSetup.printHorizontalCentered();
}

void Worksheet::setPrintHorizontalCentered(bool printHorizontalCentered)
{
    Q_D(Worksheet);
    d->pageSetup.setPrintHorizontalCentered(printHorizontalCentered);
}

std::optional<bool> Worksheet::printVerticalCentered() const
{
    Q_D(const Worksheet);
    return d->pageSetup.printVerticalCentered();
}

void Worksheet::setPrintVerticalCentered(bool printVerticalCentered)
{
    Q_D(Worksheet);
    d->pageSetup.setPrintVerticalCentered(printVerticalCentered);
}

AutoFilter &Worksheet::autofilter()
{
    Q_D(Worksheet);
    return d->autofilter;
}

AutoFilter Worksheet::autofilter() const
{
    Q_D(const Worksheet);
    return d->autofilter;
}

void Worksheet::clearAutofilter()
{
    Q_D(Worksheet);
    d->autofilter = {};
}

void Worksheet::setAutofilter(const AutoFilter &autofilter)
{
    Q_D(Worksheet);
    d->autofilter = autofilter;
}

bool Worksheet::write(int row, int column, const QVariant &value, const Format &format)
{
    Q_D(Worksheet);

    if (value.isNull()) return writeBlank(row, column, format);

    if (value.userType() == QMetaType::QString) {
        //String
        QString token = value.toString();
        bool ok;

        if (token.startsWith(QLatin1String("=")))//convert to formula
            return writeFormula(row, column, CellFormula(token), format);

        if (d->workbook->isStringsToHyperlinksEnabled() && token.contains(d->urlPattern)) //convert to url
            return writeHyperlink(row, column, QUrl(token));

        else if (d->workbook->isStringsToNumbersEnabled() && (value.toDouble(&ok), ok)) //Try convert string to number if the flag enabled.
            return writeNumeric(row, column, value.toDouble(), format);

        //normal string now
        return writeString(row, column, token, format);
    }

    if (value.userType() == qMetaTypeId<RichString>())
        return writeString(row, column, value.value<RichString>(), format);

    if (value.userType() == QMetaType::Int || value.userType() == QMetaType::UInt
               || value.userType() == QMetaType::LongLong || value.userType() == QMetaType::ULongLong
               || value.userType() == QMetaType::Double || value.userType() == QMetaType::Float) //Number
        return writeNumeric(row, column, value.toDouble(), format);

    if (value.userType() == QMetaType::Bool)//Bool
        return writeBool(row,column, value.toBool(), format);

    if (value.userType() == QMetaType::QDateTime ) // dev67
        //DateTime, Date
        //  note that QTime can't convert to QDateTime (why?)
        return writeDateTime(row, column, value.toDateTime(), format);

    if ( value.userType() == QMetaType::QDate ) // dev67
        return writeDate(row, column, value.toDate(), format);

    if (value.userType() == QMetaType::QTime)//Time
        return writeTime(row, column, value.toTime(), format);

    if (value.userType() == QMetaType::QUrl)//Url
        return writeHyperlink(row, column, value.toUrl(), format);

    //Wrong type
    return false;
}

bool Worksheet::setFormat(const CellRange &range, const Format &format)
{
    if (!range.isValid() || !format.isValid()) return false;
    for (int row = range.firstRow(); row <= range.lastRow(); ++row) {
        for (int col = range.firstColumn(); col <= range.lastColumn(); ++col)
            setFormat(row, col, format);
    }
    return true;
}

bool Worksheet::setFormat(int row, int column, const Format &format)
{
    Q_D(Worksheet);

    if (!d->addRowToDimensions(row)) return false;
    if (!d->addColumnToDimensions(column)) return false;

    Format fmt = format.isValid() ? format : d->cellFormat(row, column);
    d->workbook->styles()->addXfFormat(fmt);
    if (!d->cellTable[row][column])
        d->cellTable[row][column] = std::make_shared<Cell>(QVariant{}, Cell::Type::Number, fmt, this);
    else
        d->cellTable[row][column]->setFormat(fmt);
    return true;
}

bool Worksheet::clearFormat(const CellRange &range)
{
    bool applied = false;
    for (int row = range.firstRow(); row <= range.lastRow(); ++row) {
        for (int col = range.firstColumn(); col <= range.lastColumn(); ++col) {
            if (auto c = cell(row, col)) {
                c->setFormat({});
                applied = true;
            }
        }
    }
    return applied;
}

bool Worksheet::clearFormat(int row, int column)
{
    if (auto c = cell(row, column)) {
        c->setFormat({});
        return true;
    }
    return false;
}

void Worksheet::clearFormat()
{
    clearFormat(dimension());
}

Format Worksheet::format(const CellReference &cell) const
{
    if (auto c = this->cell(cell)) return c->format();
    return {};
}

Format Worksheet::format(int row, int column) const
{
    if (auto c = cell(row, column)) return c->format();
    return {};
}

bool Worksheet::write(const CellReference &row_column, const QVariant &value, const Format &format)
{
    if (!row_column.isValid())
        return false;

    return write(row_column.row(), row_column.column(), value, format);
}

QVariant Worksheet::read(const CellReference &cell) const
{
    if (!cell.isValid())
        return QVariant();

    return read(cell.row(), cell.column());
}

QVariant Worksheet::read(int row, int column) const
{
    Q_D(const Worksheet);

    Cell *c = cell(row, column);
    if (!c)
        return QVariant();

    if (c->hasFormula()) {
        auto fType = c->formula().type().value_or(CellFormula::Type::Normal);
        if (fType == CellFormula::Type::Normal)
            return QVariant(QLatin1String("=")+c->formula().text());

        if (fType == CellFormula::Type::Shared) {
            if (!c->formula().text().isEmpty())
                return QVariant(QLatin1String("=")+c->formula().text());

            int si = c->formula().sharedIndex().value_or(-1);
            const CellFormula &rootFormula = d->sharedFormulaMap[ si ];
            CellReference rootCellRef = rootFormula.reference().topLeft();
            QString rootFormulaText = rootFormula.text();
            QString newFormulaText = convertSharedFormula(rootFormulaText, rootCellRef, CellReference(row, column));
            return QVariant(QLatin1String("=")+newFormulaText);
        }
    }

    if (c->isDateTime())
        return c->dateTime();

    return c->value();
}

Cell *Worksheet::cell(const CellReference &ref) const
{
    if (!ref.isValid())
        return nullptr;

    return cell(ref.row(), ref.column());
}

Cell *Worksheet::cell(int row, int col) const
{
    Q_D(const Worksheet);
    auto it = d->cellTable.constFind(row);
    if (it == d->cellTable.constEnd())
        return nullptr;
    if (!it->contains(col))
        return nullptr;

    return (*it)[col].get();
}

Format WorksheetPrivate::cellFormat(int row, int col) const
{
    auto it = cellTable.constFind(row);
    if (it == cellTable.constEnd())
        return Format();
    if (!it->contains(col))
        return Format();
    return (*it)[col]->format();
}

bool Worksheet::writeString(const CellReference &row_column, const RichString &value, const Format &format)
{
    if (!row_column.isValid())
        return false;

    return writeString(row_column.row(), row_column.column(), value, format);
}

bool Worksheet::writeString(int row, int column, const RichString &value, const Format &format)
{
    Q_D(Worksheet);
//    QString content = value.toPlainString();
    if (!d->addRowToDimensions(row)) return false;
    if (!d->addColumnToDimensions(column)) return false;

//    if (content.size() > d->xls_strmax) {
//        content = content.left(d->xls_strmax);
//        error = -2;
//    }

    d->sharedStrings()->addSharedString(value);
    Format fmt = format.isValid() ? format : d->cellFormat(row, column);
    if (value.fragmentCount() == 1 && value.fragmentFormat(0).isValid())
        fmt.mergeFormat(value.fragmentFormat(0));
    d->workbook->styles()->addXfFormat(fmt);
    auto cell = std::make_shared<Cell>(value.toPlainString(), Cell::Type::SharedString, fmt, this, 0, value);
//    cell->d_ptr->richString = value;
    d->cellTable[row][column] = cell;
    return true;
}

bool Worksheet::writeString(const CellReference &row_column, const QString &value, const Format &format)
{
    if (!row_column.isValid())
        return false;

    return writeString(row_column.row(), row_column.column(), value, format);
}

bool Worksheet::writeString(int row, int column, const QString &value, const Format &format)
{
    Q_D(Worksheet);
    if (!d->addRowToDimensions(row)) return false;
    if (!d->addColumnToDimensions(column)) return false;

    RichString rs;
    if (d->workbook->isHtmlToRichStringEnabled() && Qt::mightBeRichText(value))
        rs.setHtml(value);
    else
        rs.addFragment(value, Format());

    return writeString(row, column, rs, format);
}

bool Worksheet::writeInlineString(const CellReference &row_column, const QString &value, const Format &format)
{
    if (!row_column.isValid())
        return false;

    return writeInlineString(row_column.row(), row_column.column(), value, format);
}

bool Worksheet::writeInlineString(int row, int column, const QString &value, const Format &format)
{
    Q_D(Worksheet);
    //int error = 0;
    QString content = value;
    if (!d->addRowToDimensions(row)) return false;
    if (!d->addColumnToDimensions(column)) return false;

    if (value.size() > XLSX_STRING_MAX) {
        content = value.left(XLSX_STRING_MAX);
    }

    Format fmt = format.isValid() ? format : d->cellFormat(row, column);
    d->workbook->styles()->addXfFormat(fmt);
    d->cellTable[row][column] = std::make_shared<Cell>(value, Cell::Type::InlineString, fmt, this);
    return true;
}

bool Worksheet::writeNumeric(const CellReference &row_column, double value, const Format &format)
{
    if (!row_column.isValid())
        return false;

    return writeNumeric(row_column.row(), row_column.column(), value, format);
}

bool Worksheet::writeNumeric(int row, int column, double value, const Format &format)
{
    Q_D(Worksheet);
    if (!d->addRowToDimensions(row)) return false;
    if (!d->addColumnToDimensions(column)) return false;

    Format fmt = format.isValid() ? format : d->cellFormat(row, column);
    d->workbook->styles()->addXfFormat(fmt);
    d->cellTable[row][column] = std::make_shared<Cell>(value, Cell::Type::Number, fmt, this);
    return true;
}


bool Worksheet::writeFormula(const CellReference &row_column, const CellFormula &formula, const Format &format, double result)
{
    if (!row_column.isValid())
        return false;

    return writeFormula(row_column.row(), row_column.column(), formula, format, result);
}

bool Worksheet::writeFormula(int row, int column, const CellFormula &formula_, const Format &format, double result)
{
    Q_D(Worksheet);

    if (!d->addRowToDimensions(row)) return false;
    if (!d->addColumnToDimensions(column)) return false;

    Format fmt = format.isValid() ? format : d->cellFormat(row, column);
    d->workbook->styles()->addXfFormat(fmt);

    CellFormula formula = formula_;
    if (!formula.needsRecalculation().has_value())
        formula.setNeedsRecalculation(true);
    if (formula.type().value_or(CellFormula::Type::Normal) == CellFormula::Type::Shared)  {
        //Assign proper shared index for shared formula
        int si = 0;
        while ( d->sharedFormulaMap.contains(si) )
            ++si;
        formula.setSharedIndex(si);
        d->sharedFormulaMap[si] = formula;
    }

    auto data = std::make_shared<Cell>(result, Cell::Type::Number, fmt, this);
    data->setFormula(formula);
    d->cellTable[row][column] = data;

    CellRange range = formula.reference();
    if (formula.type().value_or(CellFormula::Type::Normal) == CellFormula::Type::Shared) {
        CellFormula sf(QString(), CellFormula::Type::Shared);
        if (formula.sharedIndex().has_value())
            sf.setSharedIndex(formula.sharedIndex().value());
        for (int r=range.firstRow(); r<=range.lastRow(); ++r) {
            for (int c=range.firstColumn(); c<=range.lastColumn(); ++c) {
                if (!(r==row && c==column)) {
                    if (Cell *ce = cell(r, c)) {
                        ce->setFormula(sf);
                    } else {
                        auto newCell = std::make_shared<Cell>(result, Cell::Type::Number, fmt, this);
                        newCell->setFormula(sf);
                        d->cellTable[r][c] = newCell;
                    }
                }
            }
        }
    }

    return true;
}

bool Worksheet::writeBlank(const CellReference &row_column, const Format &format)
{
    if (!row_column.isValid())
        return false;

    return writeBlank(row_column.row(), row_column.column(), format);
}

bool Worksheet::writeBlank(int row, int column, const Format &format)
{
    Q_D(Worksheet);
    if (!d->addRowToDimensions(row)) return false;
    if (!d->addColumnToDimensions(column)) return false;

    Format fmt = format.isValid() ? format : d->cellFormat(row, column);
    d->workbook->styles()->addXfFormat(fmt);

    //Note: NumberType with an invalid QVariant value means blank.
    d->cellTable[row][column] = std::make_shared<Cell>(QVariant{}, Cell::Type::Number, fmt, this);

    return true;
}

bool Worksheet::writeBool(const CellReference &row_column, bool value, const Format &format)
{
    if (!row_column.isValid())
        return false;

    return writeBool(row_column.row(), row_column.column(), value, format);
}

bool Worksheet::writeBool(int row, int column, bool value, const Format &format)
{
    Q_D(Worksheet);
    if (!d->addRowToDimensions(row)) return false;
    if (!d->addColumnToDimensions(column)) return false;

    Format fmt = format.isValid() ? format : d->cellFormat(row, column);
    d->workbook->styles()->addXfFormat(fmt);
    d->cellTable[row][column] = std::make_shared<Cell>(value, Cell::Type::Boolean, fmt, this);

    return true;
}

bool Worksheet::writeDateTime(const CellReference &row_column, const QDateTime &dt, const Format &format)
{
    if (!row_column.isValid())
        return false;

    return writeDateTime(row_column.row(), row_column.column(), dt, format);
}

bool Worksheet::writeDateTime(int row, int column, const QDateTime &dt, const Format &format)
{
    Q_D(Worksheet);
    if (!d->addRowToDimensions(row)) return false;
    if (!d->addColumnToDimensions(column)) return false;

    Format fmt = format.isValid() ? format : d->cellFormat(row, column);
    if (!fmt.isValid() || !fmt.isDateTimeFormat())
        fmt.setNumberFormat(d->workbook->defaultDateFormat());
    d->workbook->styles()->addXfFormat(fmt);

    double value = datetimeToNumber(dt, d->workbook->date1904().value_or(false));

    d->cellTable[row][column] = std::make_shared<Cell>(value, Cell::Type::Number, fmt, this);

    return true;
}

// dev67
bool Worksheet::writeDate(const CellReference &row_column, const QDate &dt, const Format &format)
{
    if (!row_column.isValid())
        return false;

    return writeDate(row_column.row(), row_column.column(), dt, format);
}

// dev67
bool Worksheet::writeDate(int row, int column, const QDate &dt, const Format &format)
{
    Q_D(Worksheet);
    if (!d->addRowToDimensions(row)) return false;
    if (!d->addColumnToDimensions(column)) return false;

    Format fmt = format.isValid() ? format : d->cellFormat(row, column);

    if (!fmt.isValid() || !fmt.isDateTimeFormat())
        fmt.setNumberFormat(d->workbook->defaultDateFormat());

    d->workbook->styles()->addXfFormat(fmt);

    double value = datetimeToNumber(QDateTime(dt, QTime(0,0,0)), d->workbook->date1904().value_or(false));

    d->cellTable[row][column] = std::make_shared<Cell>(value, Cell::Type::Number, fmt, this);

    return true;
}

bool Worksheet::writeTime(const CellReference &row_column, const QTime &t, const Format &format)
{
    if (!row_column.isValid())
        return false;

    return writeTime(row_column.row(), row_column.column(), t, format);
}

bool Worksheet::writeTime(int row, int column, const QTime &t, const Format &format)
{
    Q_D(Worksheet);
    if (!d->addRowToDimensions(row)) return false;
    if (!d->addColumnToDimensions(column)) return false;

    Format fmt = format.isValid() ? format : d->cellFormat(row, column);
    if (!fmt.isValid() || !fmt.isDateTimeFormat())
        fmt.setNumberFormat(QLatin1String("hh:mm:ss"));
    d->workbook->styles()->addXfFormat(fmt);

    d->cellTable[row][column] = std::make_shared<Cell>(timeToNumber(t), Cell::Type::Number, fmt, this);

    return true;
}

bool Worksheet::writeHyperlink(const CellReference &row_column, const QUrl &url, const Format &format, const QString &display, const QString &tip)
{
    if (!row_column.isValid())
        return false;

    return writeHyperlink(row_column.row(), row_column.column(), url, format, display, tip);
}

bool Worksheet::writeHyperlink(int row, int column, const QUrl &url, const Format &format, const QString &display, const QString &tip)
{
    Q_D(Worksheet);
    if (!d->addRowToDimensions(row)) return false;
    if (!d->addColumnToDimensions(column)) return false;

    //int error = 0;

    QString urlString = url.toString();

    //Generate proper display string
    QString displayString = display.isEmpty() ? urlString : display;
    if (displayString.startsWith(QLatin1String("mailto:")))
        displayString.replace(QLatin1String("mailto:"), QString());
    if (displayString.size() > XLSX_STRING_MAX) {
        displayString = displayString.left(XLSX_STRING_MAX);
        //error = -2;
    }

    /*
      Location within target. If target is a workbook (or this workbook)
      this shall refer to a sheet and cell or a defined name. Can also
      be an HTML anchor if target is HTML file.

      c:\temp\file.xlsx#Sheet!A1
      http://a.com/aaa.html#aaaaa
    */
    QString locationString;
    if (url.hasFragment()) {
        locationString = url.fragment();
        urlString = url.toString(QUrl::RemoveFragment);
    }

    Format fmt = format.isValid() ? format : d->cellFormat(row, column);
    //Given a default style for hyperlink
    if (!fmt.isValid()) {
        fmt.setVerticalAlignment(Format::AlignVCenter);
        fmt.setFontColor(Qt::blue);
        fmt.setFontUnderline(Format::FontUnderlineSingle);
    }
    d->workbook->styles()->addXfFormat(fmt);

    //Write the hyperlink string as normal string.
    d->sharedStrings()->addSharedString(displayString);
    d->cellTable[row][column] = std::make_shared<Cell>(displayString, Cell::Type::SharedString, fmt, this);

    //Store the hyperlink data in a separate table
    d->urlTable[row][column] = QSharedPointer<XlsxHyperlinkData>(new XlsxHyperlinkData(XlsxHyperlinkData::External, urlString, locationString, QString(), tip));

    return true;
}

bool Worksheet::addDataValidation(const DataValidation &validation)
{
    Q_D(Worksheet);
    if (!validation.isValid() || validation.ranges().isEmpty() || validation.type() == DataValidation::Type::None)
        return false;

    d->dataValidationsList.append(validation);
    return true;
}

bool Worksheet::addDataValidation(const CellRange &range, DataValidation::Type type,
                                  const QString &formula1, std::optional<DataValidation::Predicate> predicate,
                                  const QString &formula2, bool strict)
{
    DataValidation v(type, formula1, predicate, formula2);
    v.addRange(range);
    v.setErrorMessageVisible(strict);
    return addDataValidation(v);
}

bool Worksheet::addDataValidation(const CellRange &range, const CellRange &allowableValues, bool strict)
{
    DataValidation v(DataValidation::Type::List, allowableValues.toString(true, true));
    v.setErrorMessageVisible(strict);
    v.addRange(range);
    return addDataValidation(v);
}

bool Worksheet::addDataValidation(const CellRange &range, const QTime &time1,
                                  std::optional<DataValidation::Predicate> predicate, const QTime &time2, bool strict)
{
    DataValidation v(time1, predicate, time2);
    v.addRange(range);
    v.setErrorMessageVisible(strict);
    return addDataValidation(v);
}

bool Worksheet::addDataValidation(const CellRange &range, const QDate &date1,
                                  std::optional<DataValidation::Predicate> predicate,
                                  const QDate &date2, bool strict)
{
    DataValidation v(date1, predicate, date2);
    v.addRange(range);
    v.setErrorMessageVisible(strict);
    return addDataValidation(v);
}

bool Worksheet::addDataValidation(const CellRange &range, int len1,
                                  std::optional<DataValidation::Predicate> predicate,
                                  std::optional<int> len2, bool strict)
{
    DataValidation v(len1, predicate, len2);
    v.addRange(range);
    v.setErrorMessageVisible(strict);
    return addDataValidation(v);
}

std::optional<bool> Worksheet::dataValidationPromptsDisabled() const
{
     Q_D(const Worksheet);
     return d->disableValidationPrompts;
}

void Worksheet::setDataValidationPromptsDisabled(bool disabled)
{
    Q_D(Worksheet);
    d->disableValidationPrompts = disabled;
}

void Worksheet::clearDataValidation()
{
    Q_D(Worksheet);
    d->dataValidationsList.clear();
}

bool Worksheet::hasDataValidation() const
{
    Q_D(const Worksheet);
    return !d->dataValidationsList.isEmpty();
}

QList<DataValidation> Worksheet::dataValidationRules() const
{
    Q_D(const Worksheet);
    return d->dataValidationsList;
}

DataValidation Worksheet::dataValidation(int index) const
{
    Q_D(const Worksheet);
    return d->dataValidationsList.value(index);
}

DataValidation &Worksheet::dataValidation(int index)
{
    Q_D(Worksheet);
    return d->dataValidationsList[index];
}

bool Worksheet::removeDataValidation(int index)
{
    Q_D(Worksheet);
    if (index < 0 || index >= d->dataValidationsList.size()) return false;
    d->dataValidationsList.removeAt(index);
    return true;
}

int Worksheet::dataValidationsCount() const
{
    Q_D(const Worksheet);
    return d->dataValidationsList.size();
}

void Worksheet::clearConditionalformatting()
{
    Q_D(Worksheet);
    d->conditionalFormattingList.clear();
}

bool Worksheet::hasConditionalFormatting() const
{
    Q_D(const Worksheet);
    return !d->conditionalFormattingList.isEmpty();
}
int Worksheet::conditionalFormattingCount() const
{
    Q_D(const Worksheet);
    return d->conditionalFormattingList.size();
}
QList<ConditionalFormatting> Worksheet::conditionalFormattingRules() const
{
    Q_D(const Worksheet);
    return d->conditionalFormattingList;
}

ConditionalFormatting Worksheet::conditionalFormatting(int index) const
{
    Q_D(const Worksheet);
    return d->conditionalFormattingList.value(index);
}
ConditionalFormatting &Worksheet::conditionalFormatting(int index)
{
    Q_D(Worksheet);
    return d->conditionalFormattingList[index];
}

bool Worksheet::removeConditionalFormatting(int index)
{
    Q_D(Worksheet);
    if (index < 0 || index >= d->conditionalFormattingList.size()) return false;
    d->conditionalFormattingList.removeAt(index);
    return true;
}

bool Worksheet::addConditionalFormatting(const ConditionalFormatting &cf)
{
    Q_D(Worksheet);
    if (cf.ranges().isEmpty())
        return false;

    for (const auto &rule: cf.d->cfRules) {
        if (!rule->dxfFormat.isEmpty())
            d->workbook->styles()->addDxfFormat(rule->dxfFormat);
    }
    d->conditionalFormattingList.append(cf);
    return true;
}

int Worksheet::insertImage(const CellReference &ref, const QImage &image)
{
    if (!ref.isValid()) return -1;
    return insertImage(ref.row(), ref.column(), image);
}

int Worksheet::insertImage(int row, int column, const QImage &image)
{
    Q_D(Worksheet);

    if (!image.isNull()) {
        if (!d->drawing)
            d->drawing = std::make_shared<Drawing>(this, F_NewFromScratch);

        DrawingOneCellAnchor* anchor = new DrawingOneCellAnchor(d->drawing.get());

        /*
        The size are expressed as English Metric Units (EMUs).
        EMU is 1/360 000 of centimiter.
        */
        anchor->from = XlsxMarker(row, column, 0, 0);
        int dpiX = int(double(image.dotsPerMeterX()) / 39.3700787); //calculate dpi from dpm
        int dpiY = int(double(image.dotsPerMeterY()) / 39.3700787);
        anchor->ext = QPair<Coordinate, Coordinate>(Coordinate::fromPixels(image.width(), dpiX),
                                                    Coordinate::fromPixels(image.height(), dpiY));
        anchor->setObjectPicture(image);
        return d->drawing->anchors.size()-1;
    }
    return -1;
}

QImage Worksheet::image(int index) const
{
    Q_D(const Worksheet);

    if (index < 0) return {};

    if (d->drawing) {
        int currentIndex = -1;
        for (auto anchor: qAsConst(d->drawing->anchors)) {
            if (anchor && anchor->isPicture()) {
                currentIndex++;
                if (currentIndex == index) {
                    QImage img;
                    anchor->getObjectPicture(img);
                    return img;
                }
            }
        }
    }

   return {};
}

QImage Worksheet::image(int row, int column) const
{
    Q_D(const Worksheet);

    if (d->drawing) {
        for (auto anchor: qAsConst(d->drawing->anchors)) {
            if (anchor && anchor->row() == row && anchor->col() == column) {
                QImage img;
                anchor->getObjectPicture(img);
                return img;
            }
        }
    }
    return {};
}

int Worksheet::imagesCount() const
{
    Q_D(const Worksheet);
    int count = 0;
    if (d->drawing) {
        for (auto anchor: qAsConst(d->drawing->anchors)) {
            if (anchor && anchor->isPicture())
                count++;
        }
    }
    return count;
}

bool Worksheet::removeImage(int row, int column)
{
    Q_D(Worksheet);
    if (d->drawing) {
        auto anchor = std::find_if(std::begin(d->drawing->anchors),
                                   std::end(d->drawing->anchors),
                                   [row, column](auto anchor)
        {return anchor && anchor->row() == row && anchor->col() == column;});
        if (anchor != std::end(d->drawing->anchors)) {
            bool result = (*anchor)->removeObjectPicture();
            delete *anchor;
            d->drawing->anchors.erase(anchor);

            return result;
        }
    }
    return false;
}

bool Worksheet::removeImage(int index)
{
    Q_D(Worksheet);

    if (index < 0) return false;
    if (d->drawing) {
        int currentIndex = -1;
        for (auto anchor: qAsConst(d->drawing->anchors)) {
            if (anchor && anchor->isPicture()) {
                currentIndex++;
                if (currentIndex == index) {
                    bool result = anchor->removeObjectPicture();
                    delete anchor;
                    d->drawing->anchors.removeOne(anchor);
                    return result;
                }
            }
        }
    }

    return false;
}

bool Worksheet::changeImage(int index, const QString &fileName, bool keepSize)
{
    Q_D(const Worksheet);

    if (index < 0) return false;
    int currentIndex = -1;
    if (d->drawing) {
        for (auto anchor: qAsConst(d->drawing->anchors)) {
            if (anchor && anchor->isPicture()) {
                currentIndex++;
                if (currentIndex == index) break;
            }
        }
    }
    if (currentIndex > -1) {
        QImage newpic(fileName);
        if (newpic.isNull()) return false;

        const QString suffix = fileName.mid(fileName.lastIndexOf(QLatin1Char('.'))+1);
        QString mimetypemy;
        if(QString::compare(QLatin1String("jpg"), suffix, Qt::CaseInsensitive)==0)
           mimetypemy=QStringLiteral("image/jpeg");
        if(QString::compare(QLatin1String("bmp"), suffix, Qt::CaseInsensitive)==0)
           mimetypemy=QStringLiteral("image/bmp");
        if(QString::compare(QLatin1String("gif"), suffix, Qt::CaseInsensitive)==0)
           mimetypemy=QStringLiteral("image/gif");
        if(QString::compare(QLatin1String("png"), suffix, Qt::CaseInsensitive)==0)
           mimetypemy=QStringLiteral("image/png");

        QByteArray ba;
        QBuffer buffer(&ba);
        buffer.setBuffer(&ba);
        buffer.open(QIODevice::WriteOnly);
        newpic.save(&buffer,suffix.toLocal8Bit().data());

        if (auto anchor = dynamic_cast<DrawingOneCellAnchor*>(d->drawing->anchors[currentIndex])) {
            anchor->setObjectPicture(newpic);
            if (!keepSize) {
                int dpiX = int(double(newpic.dotsPerMeterX()) / 39.3700787); //calculate dpi from dpm
                int dpiY = int(double(newpic.dotsPerMeterY()) / 39.3700787);
                anchor->ext = QPair<Coordinate, Coordinate>(Coordinate::fromPixels(newpic.width(), dpiX),
                                                            Coordinate::fromPixels(newpic.height(), dpiY));
            }
            return true;
        }
    }

    return false;
}

Chart *Worksheet::insertChart(int row, int column, const QSize &size)
{
    Q_D(Worksheet);

    if (!d->drawing)
        d->drawing = std::make_shared<Drawing>(this, F_NewFromScratch);

    if (!d->addRowToDimensions(row)) return nullptr;
    if (!d->addColumnToDimensions(column)) return nullptr;

    DrawingOneCellAnchor *anchor = new DrawingOneCellAnchor(d->drawing.get());

    /*
        The size are expressed as English Metric Units (EMUs). There are
        12,700 EMUs per point. Therefore, 12,700 * 3 /4 = 9,525 EMUs per
        pixel
    */
    anchor->from = XlsxMarker(row, column, 0, 0);
    anchor->ext = QPair<Coordinate, Coordinate>(Coordinate(qint64(size.width() * 9525)),
                                                Coordinate(qint64(size.height() * 9525)));

    QSharedPointer<Chart> chart = QSharedPointer<Chart>(new Chart(this, F_NewFromScratch));
    anchor->setObjectGraphicFrame(chart);

    return chart.data();
}

Chart *Worksheet::insertChart(int row, int column, Coordinate width, Coordinate height)
{
    Q_D(Worksheet);

    if (!d->drawing)
        d->drawing = std::make_shared<Drawing>(this, F_NewFromScratch);

    if (!d->addRowToDimensions(row)) return nullptr;
    if (!d->addColumnToDimensions(column)) return nullptr;

    DrawingOneCellAnchor *anchor = new DrawingOneCellAnchor(d->drawing.get());

    anchor->from = XlsxMarker(row, column, 0, 0);
    anchor->ext = QPair<Coordinate, Coordinate>(width, height);

    QSharedPointer<Chart> chart = QSharedPointer<Chart>(new Chart(this, F_NewFromScratch));
    anchor->setObjectGraphicFrame(chart);

    return chart.data();
}

Chart *Worksheet::chart(int index) const
{
    Q_D(const Worksheet);

    if (index < 0) return nullptr;

    if (d->drawing) {
        int currentIndex = -1;
        for (auto anchor: qAsConst(d->drawing->anchors)) {
            if (anchor && anchor->chart()) {
                currentIndex++;
                if (currentIndex == index)
                    return anchor->chart().get();
            }
        }
    }

    return nullptr;
}

Chart *Worksheet::chart(int row, int column) const
{
    Q_D(const Worksheet);
    if (d->drawing) {
        for (auto anchor: qAsConst(d->drawing->anchors)) {
            if (anchor && anchor->row() == row && anchor->col() == column) {
                return anchor->chart().get();
            }
        }
    }
    return nullptr;
}

bool Worksheet::removeChart(int row, int column)
{
    Q_D(Worksheet);
    if (d->drawing) {
        auto anchor = std::find_if(std::begin(d->drawing->anchors),
                                   std::end(d->drawing->anchors),
                                   [row, column](auto anchor)
        {return anchor && anchor->row() == row && anchor->col() == column;});
        if (anchor != std::end(d->drawing->anchors)) {
            bool result = (*anchor)->removeObjectGraphicFrame();
            delete *anchor;
            d->drawing->anchors.erase(anchor);

            return result;
        }
    }
    return false;
}

bool Worksheet::removeChart(int index)
{
    Q_D(Worksheet);

    if (index < 0) return false;
    if (d->drawing) {
        int currentIndex = -1;
        for (auto anchor: qAsConst(d->drawing->anchors)) {
            if (anchor && anchor->chart()) {
                currentIndex++;
                if (currentIndex == index) {
                    bool result = anchor->removeObjectGraphicFrame();
                    delete anchor;
                    d->drawing->anchors.removeOne(anchor);
                    return result;
                }
            }
        }
    }

    return false;
}

int Worksheet::chartsCount() const
{
    Q_D(const Worksheet);
    int count = 0;

    if (d->drawing) {
        for (auto anchor: qAsConst(d->drawing->anchors)) {
            if (anchor && anchor->isChart())
                count++;
        }
    }
    return count;
}

bool Worksheet::mergeCells(const CellRange &range, const Format &format)
{
    Q_D(Worksheet);
    if (range.rowCount() < 2 && range.columnCount() < 2)
        return false;

    if (!d->addRowToDimensions(range.firstRow())) return false;
    if (!d->addColumnToDimensions(range.firstColumn())) return false;

    if (format.isValid())
        d->workbook->styles()->addXfFormat(format);

    for (int row = range.firstRow(); row <= range.lastRow(); ++row) {
        for (int col = range.firstColumn(); col <= range.lastColumn(); ++col) {
            if (row == range.firstRow() && col == range.firstColumn()) {
                if (Cell *c = cell(row, col)) {
                    if (format.isValid())
                        c->setFormat(format);
                }
                else
                    writeBlank(row, col, format);
            }
            else
                writeBlank(row, col, format);
        }
    }

    d->merges.append(range);
    return true;
}

bool Worksheet::unmergeCells(const CellRange &range)
{
    Q_D(Worksheet);
    return d->merges.removeOne(range);
}

QList<CellRange> Worksheet::mergedCells() const
{
    Q_D(const Worksheet);
    return d->merges;
}

/*!
 * \internal
 */
void Worksheet::saveToXmlFile(QIODevice *device) const
{
    Q_D(const Worksheet);
    d->relationships->clear();

    QXmlStreamWriter writer(device);

    writer.writeStartDocument(QLatin1String("1.0"), true);
    writer.writeStartElement(QLatin1String("worksheet"));
    writer.writeAttribute(QLatin1String("xmlns"), QLatin1String("http://schemas.openxmlformats.org/spreadsheetml/2006/main"));
    writer.writeAttribute(QLatin1String("xmlns:r"), QLatin1String("http://schemas.openxmlformats.org/officeDocument/2006/relationships"));

    //for Excel 2010
    //    writer.writeAttribute("xmlns:mc", "http://schemas.openxmlformats.org/markup-compatibility/2006");
    //    writer.writeAttribute("xmlns:x14ac", "http://schemas.microsoft.com/office/spreadsheetml/2009/9/ac");
    //    writer.writeAttribute("mc:Ignorable", "x14ac");

    //1. sheetPr - CT_SheetPr
    auto sp = d->sheetProperties;
    if (d->autofilter.isValid()) sp.filterMode = true;
    if (sp.isValid()) sp.write(writer, QLatin1String("sheetPr"));

    //2. dimension
    writer.writeStartElement(QLatin1String("dimension"));
    writer.writeAttribute(QLatin1String("ref"), d->generateDimensionString());
    writer.writeEndElement();//dimension

    //3. sheetViews
    auto views = d->sheetViews;
    if (views.isEmpty()) views << SheetView();
    writer.writeStartElement(QLatin1String("sheetViews"));
    for (auto &view: views)
        if (view.isValid()) view.write(writer, QLatin1String("sheetView"));
    writer.writeEndElement();//sheetViews

    //4. sheetFormatPr
    d->sheetFormatProperties.write(writer, QLatin1String("sheetFormatPr"));

    //5. cols
    if (!d->colsInfo.empty()) {
        writer.writeStartElement(QLatin1String("cols"));
        const auto intervals = d->getIntervals();
        for (auto interval: intervals) {
            auto col_info = d->colsInfo.value(interval.first);
            writer.writeStartElement(QLatin1String("col"));
            writer.writeAttribute(QLatin1String("min"), QString::number(interval.first));
            writer.writeAttribute(QLatin1String("max"), QString::number(interval.second));
            if (col_info.width.has_value()) {
                writer.writeAttribute(QLatin1String("width"), QString::number(col_info.width.value(), 'g', 15));
                writeAttribute(writer, QLatin1String("customWidth"), true);
            }
            if (!col_info.format.isEmpty())
                writer.writeAttribute(QLatin1String("style"), QString::number(col_info.format.xfIndex()));
            writeAttribute(writer, QLatin1String("hidden"), col_info.hidden);

            writeAttribute(writer, QLatin1String("outlineLevel"), col_info.outlineLevel);
            writeAttribute(writer, QLatin1String("collapsed"), col_info.collapsed);
            writer.writeEndElement();//col
        }
        writer.writeEndElement();//cols
    }
    //6. sheetData
    writer.writeStartElement(QLatin1String("sheetData"));
    if (d->dimension.isValid())
        d->saveXmlSheetData(writer);
    writer.writeEndElement();//sheetData

    //7. sheetCalcPr
    if (d->fullCalcOnLoad.has_value()) {
        writer.writeEmptyElement(QLatin1String("sheetCalcPr"));
        writeAttribute(writer, QLatin1String("fullCalcOnLoad"), d->fullCalcOnLoad);
    }

    //8. sheet protection
    d->sheetProtection.write(writer);

    //9. protectedRanges

    //10. scenarios

    //11. autoFilter
    d->autofilter.write(writer, QLatin1String("autoFilter"));

    //12. sortState
    d->sortState.write(writer, QLatin1String("sortState"));

    //13. dataConsolidate

    //14. customSheetViews

    //15. mergeCells
    d->saveXmlMergeCells(writer);

    //16. phoneticPr

    //17. conditionalFormatting
    for (const ConditionalFormatting &cf : qAsConst(d->conditionalFormattingList))
        cf.saveToXml(writer);

    //18. dataValidations
    d->saveXmlDataValidations(writer);

    //19. hyperlinks
    d->saveXmlHyperlinks(writer);

    //20. printOptions
    d->pageSetup.writeWorksheet(writer, QLatin1String("printOptions"));

    //21. pageMargins
    d->pageMargins.write(writer);

    //22. pageSetup
    d->pageSetup.writeWorksheet(writer, QLatin1String("pageSetup"));

    //23. headerFooter
    d->headerFooter.write(writer, QLatin1String("headerFooter"));

    //24. rowBreaks

    //25. colBreaks

    //26. customProperties

    //27. cellWatches

    //28. ignoredErrors

    //29. smartTags

    //30. drawing
    d->saveXmlDrawings(writer);

    //31. drawingHF

    //32. picture
    if (d->pictureFile) {
        d->relationships->addDocumentRelationship(QLatin1String("/image"), QString("../media/image%1.%2")
                                                   .arg(d->pictureFile->index()+1)
                                                   .arg(d->pictureFile->suffix()));
        writer.writeEmptyElement(QLatin1String("picture"));
        writer.writeAttribute(QLatin1String("r:id"), QString("rId%1").arg(d->relationships->count()));
    }

    //33. oleObjects

    //34. controls

    //35. webPublishItems

    //36. tableParts

    //37. extLst
    d->extLst.write(writer, "extLst");

    writer.writeEndElement(); // worksheet
    writer.writeEndDocument();
}

void WorksheetPrivate::saveXmlSheetData(QXmlStreamWriter &writer) const
{
    QMap<int, QString> rowSpans = calculateSpans();
    for (int row = dimension.firstRow(); row <= dimension.lastRow(); row++) {
        auto ctIt = cellTable.constFind(row);
        auto riIt = rowsInfo.constFind(row);
        auto cmIt = comments.constFind(row);
        if (ctIt == cellTable.constEnd() && riIt == rowsInfo.constEnd() && cmIt == comments.constEnd()) {
            //Only process rows with cell data / comments / formatting
            continue;
        }

        int span_index = (row-1) / 16;
        QString span;
        auto rsIt = rowSpans.constFind(span_index);
        if (rsIt != rowSpans.constEnd())
            span = rsIt.value();

        writer.writeStartElement(QLatin1String("row"));
        writer.writeAttribute(QLatin1String("r"), QString::number(row));

        writeAttribute(writer, QLatin1String("spans"), span);

        if (riIt != rowsInfo.constEnd()) {
            QSharedPointer<XlsxRowInfo> rowInfo = riIt.value();
            if (!rowInfo->format.isEmpty()) {
                writer.writeAttribute(QLatin1String("s"), QString::number(rowInfo->format.xfIndex()));
                writeAttribute(writer, QLatin1String("customFormat"), true);
            }

            if (rowInfo->height.has_value()) {
                writeAttribute(writer, QLatin1String("ht"), rowInfo->height);
                writeAttribute(writer, QLatin1String("customHeight"), true);
            }

            writeAttribute(writer, QLatin1String("hidden"), rowInfo->hidden);
            if (rowInfo->outlineLevel.value_or(0) > 0)
                writeAttribute(writer, QLatin1String("outlineLevel"), rowInfo->outlineLevel);
            writeAttribute(writer, QLatin1String("collapsed"), rowInfo->collapsed);
        }

        //Write cell data if row contains filled cells
        if (ctIt != cellTable.constEnd()) {
            for (int col_num = dimension.firstColumn(); col_num <= dimension.lastColumn(); col_num++) {
                if (ctIt->contains(col_num))
                    saveXmlCellData(writer, row, col_num, (*ctIt)[col_num]);
            }
        }
        writer.writeEndElement(); //row
    }
}

void WorksheetPrivate::saveXmlCellData(QXmlStreamWriter &writer, int row, int col, std::shared_ptr<Cell> cell) const
{
    //This is the innermost loop so efficiency is important.
    QString cell_pos = CellReference(row, col).toString();

    writer.writeStartElement(QLatin1String("c"));
    writer.writeAttribute(QLatin1String("r"), cell_pos);

    //Style used by the cell, row or col
    if (!cell->format().isEmpty())
        writer.writeAttribute(QLatin1String("s"), QString::number(cell->format().xfIndex()));
    else if (auto rIt = rowsInfo.constFind(row); rIt != rowsInfo.constEnd() && !(*rIt)->format.isEmpty())
        writer.writeAttribute(QLatin1String("s"), QString::number((*rIt)->format.xfIndex()));
    else if (auto cIt = colsInfo.constFind(col); cIt != colsInfo.constEnd() && !(*cIt).format.isEmpty())
        writer.writeAttribute(QLatin1String("s"), QString::number((*cIt).format.xfIndex()));

    switch (cell->type()) {
        case Cell::Type::SharedString: { // 's'
            auto sst_idx = sharedStrings()->getSharedStringIndex(cell->isRichString()
                                                                     ? cell->richString()
                                                                     : RichString(cell->value().toString()));
            if (sst_idx.has_value()) {
                writer.writeAttribute(QLatin1String("t"), QLatin1String("s"));
                writer.writeTextElement(QLatin1String("v"), QString::number(sst_idx.value()));
            }
            break;
        }
        case Cell::Type::InlineString: {// 'inlineStr'
            writer.writeAttribute(QLatin1String("t"), QLatin1String("inlineStr"));
            writer.writeStartElement(QLatin1String("is"));
            if (cell->isRichString()) {
                //Rich text string
                const RichString &string = cell->richString();
                for (int i=0; i<string.fragmentCount(); ++i) {
                    writer.writeStartElement(QLatin1String("r"));
                    if (string.fragmentFormat(i).hasFontData())
                    {
                        writer.writeStartElement(QLatin1String("rPr"));
                        //TODO
                        writer.writeEndElement();// rPr
                    }
                    writer.writeStartElement(QLatin1String("t"));
                    if (isSpaceReserveNeeded(string.fragmentText(i)))
                        writer.writeAttribute(QLatin1String("xml:space"), QLatin1String("preserve"));
                    writer.writeCharacters(string.fragmentText(i));
                    writer.writeEndElement();// t
                    writer.writeEndElement(); // r
                }
            }
            else {
                writer.writeStartElement(QLatin1String("t"));
                QString string = cell->value().toString();
                if (isSpaceReserveNeeded(string))
                    writer.writeAttribute(QLatin1String("xml:space"), QLatin1String("preserve"));
                writer.writeCharacters(string);
                writer.writeEndElement(); // t
            }
            writer.writeEndElement();//is
            break;
        }
        case Cell::Type::Number: {// 'n'
            writer.writeAttribute(QLatin1String("t"), QLatin1String("n")); // dev67

            if (cell->hasFormula())
                cell->formula().saveToXml(writer);

            if (cell->value().isValid()) {   //note that, invalid value means 'v' is blank
                double value = cell->value().toDouble();
                writer.writeTextElement(QLatin1String("v"), QString::number(value, 'g', 15));
            }
            break;
        }
        case Cell::Type::String: {// 'str'
            writer.writeAttribute(QLatin1String("t"), QLatin1String("str"));
            if (cell->hasFormula())
                cell->formula().saveToXml(writer);

            writer.writeTextElement(QLatin1String("v"), cell->value().toString());
            break;
        }
        case Cell::Type::Boolean: {// 'b'
            writer.writeAttribute(QLatin1String("t"), QLatin1String("b"));

            // dev34
            if (cell->hasFormula())
                cell->formula().saveToXml(writer);

            writer.writeTextElement(QLatin1String("v"), cell->value().toBool() ? QLatin1String("1") : QLatin1String("0"));
            break;
        }
        case Cell::Type::Date: { // 'd'
            // dev67
//            double num = cell->value().toDouble();
//            if (!q->workbook()->isDate1904() && num > 60) // for mac os excel
//                num = num - 1;

            // number type. see for 18.18.11 ST_CellType (Cell Type) more information.
            writer.writeAttribute(QLatin1String("t"), QLatin1String("n"));
            writer.writeTextElement(QLatin1String("v"), cell->value().toString() );
            break;
        }
        case Cell::Type::Error: {// 'e'
            writer.writeAttribute(QLatin1String("t"), QLatin1String("e"));
            writer.writeTextElement(QLatin1String("v"), cell->value().toString() );
            break;
        }
        default: { //Cell::CustomType
            if (cell->hasFormula())
                cell->formula().saveToXml(writer);

            if (cell->value().isValid()) {   //note that, invalid value means 'v' is blank
                double value = cell->value().toDouble();
                writer.writeTextElement(QLatin1String("v"), QString::number(value, 'g', 15));
            }
        }
    }
    writer.writeEndElement(); // c
}

void WorksheetPrivate::saveXmlMergeCells(QXmlStreamWriter &writer) const
{
    if (merges.isEmpty())
        return;

    writer.writeStartElement(QLatin1String("mergeCells"));
    writer.writeAttribute(QLatin1String("count"), QString::number(merges.size()));

    for (const CellRange &range : qAsConst(merges)) {
        writer.writeEmptyElement(QLatin1String("mergeCell"));
        writer.writeAttribute(QLatin1String("ref"), range.toString());
    }

    writer.writeEndElement(); //mergeCells
}

void WorksheetPrivate::saveXmlDataValidations(QXmlStreamWriter &writer) const
{
    if (dataValidationsList.isEmpty())
        return;

    writer.writeStartElement(QLatin1String("dataValidations"));
    writer.writeAttribute(QLatin1String("count"), QString::number(dataValidationsList.size()));
    writeAttribute(writer, QLatin1String("disablePrompts"), disableValidationPrompts);
    writeAttribute(writer, QLatin1String("xWindow"), dataValidationXWindow);
    writeAttribute(writer, QLatin1String("yWindow"), dataValidationYWindow);

    for (const DataValidation &validation : qAsConst(dataValidationsList))
        validation.write(writer);

    writer.writeEndElement(); //dataValidations
}

void WorksheetPrivate::saveXmlHyperlinks(QXmlStreamWriter &writer) const
{
    if (urlTable.isEmpty())
        return;

    writer.writeStartElement(QLatin1String("hyperlinks"));

    for (auto it1 = urlTable.constBegin(); it1 != urlTable.constEnd(); ++it1) {
        int row = it1.key();
        for (auto it2 = it1.value().constBegin(); it2 != it1.value().constEnd(); ++it2) {
            int col = it2.key();
            auto data = it2.value();

            // dev57
            // writer.writeEmptyElement(QLatin1String("hyperlink"));
            writer.writeStartElement(QLatin1String("hyperlink"));
            writer.writeAttribute(QLatin1String("ref"), CellReference(row, col).toString()); // required field

            if (data->linkType == XlsxHyperlinkData::External) {
                // Update relationships
                relationships->addWorksheetRelationship(QLatin1String("/hyperlink"), data->target, QLatin1String("External"));
                writer.writeAttribute(QLatin1String("r:id"), QString("rId%1").arg(relationships->count()));
            }

            writeAttribute(writer, QLatin1String("location"), data->location);
            writeAttribute(writer, QLatin1String("display"), data->display);
            writeAttribute(writer, QLatin1String("tooltip"), data->tooltip);

            // dev57
            writer.writeEndElement(); // hyperlink
        }
    }

    writer.writeEndElement(); // hyperlinks
}

void WorksheetPrivate::saveXmlDrawings(QXmlStreamWriter &writer) const
{
    if (!drawing)
        return;

    int idx = workbook->drawings().indexOf(drawing.get());
    relationships->addWorksheetRelationship(QLatin1String("/drawing"), QString("../drawings/drawing%1.xml").arg(idx+1));

    writer.writeEmptyElement(QLatin1String("drawing"));
    writer.writeAttribute(QLatin1String("r:id"), QString("rId%1").arg(relationships->count()));
}

bool WorksheetPrivate::isColumnRangeValid(int colFirst, int colLast) const
{
    if (colFirst > colLast)
        return false;
    return columnValid(colFirst) && columnValid(colLast);
}

QList<QPair<int, int> > WorksheetPrivate::getIntervals() const
{
    QList<QPair<int, int> > result;

    if (!colsInfo.isEmpty()) {
        auto it = colsInfo.constBegin();
        auto val = it.value();
        auto firstCol = it.key();
        auto currentCol = it.key();
        it++;

        for (; ; ++it) {
            if (it == colsInfo.constEnd()) {
                result << QPair<int,int>{firstCol, currentCol};
                break;
            }
            currentCol++;
            if (it.value() != val || currentCol != it.key()) {
                result << QPair<int,int>{firstCol, --currentCol};
                val = it.value();
                firstCol = it.key();
                currentCol = it.key();
            }
        }
    }

    return result;
}

bool Worksheet::setColumnWidth(const CellRange &range, double width)
{
    if (!range.isValid())
        return false;

    return setColumnWidth(range.firstColumn(), range.lastColumn(), width);
}

bool Worksheet::setColumnFormat(const CellRange& range, const Format &format)
{
    if (!range.isValid())
        return false;

    return setColumnFormat(range.firstColumn(), range.lastColumn(), format);
}

bool Worksheet::setColumnHidden(const CellRange &range, bool hidden)
{
    if (!range.isValid())
        return false;

    return setColumnHidden(range.firstColumn(), range.lastColumn(), hidden);
}

bool Worksheet::setColumnWidth(int colFirst, int colLast, double width)
{
    Q_D(Worksheet);

    if (!d->addColumnToDimensions(colFirst)) return false;
    if (!d->addColumnToDimensions(colLast)) return false;

    for (int i=colFirst; i<=colLast; ++i)
        d->colsInfo[i].width = width;

    return true;
}

bool Worksheet::setColumnWidth(int col, double width)
{
    return setColumnWidth(col, col, width);
}

bool Worksheet::setColumnFormat(int colFirst, int colLast, const Format &format)
{
    Q_D(Worksheet);

    if (!d->addColumnToDimensions(colFirst)) return false;
    if (!d->addColumnToDimensions(colLast)) return false;

    for (int i=colFirst; i<=colLast; ++i)
        d->colsInfo[i].format = format;

    d->workbook->styles()->addXfFormat(format);
    return true;
}

bool Worksheet::setColumnFormat(int col, const Format &format)
{
    return setColumnFormat(col, col, format);
}

bool Worksheet::setColumnHidden(int colFirst, int colLast, bool hidden)
{
    Q_D(Worksheet);

    if (!d->addColumnToDimensions(colFirst)) return false;
    if (!d->addColumnToDimensions(colLast)) return false;

    for (int i=colFirst; i<=colLast; ++i)
        d->colsInfo[i].hidden = hidden;

    return true;
}

bool Worksheet::setColumnHidden(int col, bool hidden)
{
    return setColumnHidden(col, col, hidden);
}

std::optional<double> Worksheet::columnWidth(int column) const
{
    Q_D(const Worksheet);

    if (d->colsInfo.contains(column))
        return d->colsInfo.value(column).width;

    return {};
}

Format Worksheet::columnFormat(int column) const
{
    Q_D(const Worksheet);

    if (d->colsInfo.contains(column))
        return d->colsInfo.value(column).format;

    return Format();
}

std::optional<bool> Worksheet::isColumnHidden(int column) const
{
    Q_D (const Worksheet);

    if (d->colsInfo.contains(column))
        return d->colsInfo.value(column).hidden;

    return {};
}

bool Worksheet::setRowHeight(int rowFirst, int rowLast, double height)
{
    Q_D(Worksheet);

    if (!d->addRowToDimensions(rowFirst)) return false;
    if (!d->addRowToDimensions(rowLast)) return false;

    for (int row = rowFirst; row <= rowLast; ++row) {
        if (!d->rowsInfo.contains(row))
            d->rowsInfo[row] = QSharedPointer<XlsxRowInfo>(new XlsxRowInfo());
        d->rowsInfo[row]->height = height;
        d->addRowToDimensions(row);
    }

    return true;
}

bool Worksheet::setRowHeight(int row, double height)
{
    return setRowHeight(row, row, height);
}

bool Worksheet::setRowFormat(int rowFirst,int rowLast, const Format &format)
{
    Q_D(Worksheet);

    if (!d->addRowToDimensions(rowFirst)) return false;
    if (!d->addRowToDimensions(rowLast)) return false;

    for (int row = rowFirst; row <= rowLast; ++row) {
        if (!d->rowsInfo.contains(row))
            d->rowsInfo[row] = QSharedPointer<XlsxRowInfo>(new XlsxRowInfo());
        d->rowsInfo[row]->format = format;
        d->addRowToDimensions(row);
    }
    d->workbook->styles()->addXfFormat(format);

    return true;
}

bool Worksheet::setRowFormat(int row, const Format &format)
{
    return setRowFormat(row, row, format);
}

bool Worksheet::setRowHidden(int rowFirst,int rowLast, bool hidden)
{
    Q_D(Worksheet);

    if (!d->addRowToDimensions(rowFirst)) return false;
    if (!d->addRowToDimensions(rowLast)) return false;

    for (int row = rowFirst; row <= rowLast; ++row) {
        if (!d->rowsInfo.contains(row))
            d->rowsInfo[row] = QSharedPointer<XlsxRowInfo>(new XlsxRowInfo());
        d->rowsInfo[row]->hidden = hidden;
        d->addRowToDimensions(row);
    }

    return true;
}


std::optional<double> Worksheet::rowHeight(int row) const
{
    Q_D(const Worksheet);

    if (d->rowsInfo.contains(row))
        return d->rowsInfo.value(row)->height;
    return {};
}

Format Worksheet::rowFormat(int row) const
{
    Q_D(const Worksheet);

    if (d->rowsInfo.contains(row))
        return d->rowsInfo.value(row)->format;
    return {};
}

std::optional<bool> Worksheet::isRowHidden(int row) const
{
    Q_D(const Worksheet);
    if (d->rowsInfo.contains(row))
        return d->rowsInfo.value(row)->hidden;
    return {};
}

bool Worksheet::groupRows(int rowFirst, int rowLast, bool collapsed)
{
    Q_D(Worksheet);

    if (!d->addRowToDimensions(rowFirst)) return false;
    if (!d->addRowToDimensions(rowLast)) return false;

    //calculate current max outline level
    int maximumOutlineLevel = 0;
    for (int i = rowFirst; i <= rowLast; ++i) {
        if (d->rowsInfo.contains(i))
            maximumOutlineLevel = std::max(d->rowsInfo.value(i)->outlineLevel.value_or(0),
                                           maximumOutlineLevel);
    }
    if (maximumOutlineLevel >= 7) return false;

    for (int row=rowFirst; row<=rowLast; ++row) {
        auto it = d->rowsInfo.find(row);
        if (it != d->rowsInfo.end()) {
            (*it)->outlineLevel = (*it)->outlineLevel.value_or(0) + 1;
        } else {
            QSharedPointer<XlsxRowInfo> info(new XlsxRowInfo);
            info->outlineLevel = info->outlineLevel.value_or(0) + 1;
            it = d->rowsInfo.insert(row, info);
        }
        if (collapsed)
            (*it)->hidden = true;
    }
    if (collapsed) {
        auto it = d->rowsInfo.find(rowLast+1);
        if (it == d->rowsInfo.end())
            it = d->rowsInfo.insert(rowLast+1, QSharedPointer<XlsxRowInfo>(new XlsxRowInfo));
        (*it)->collapsed = true;
    }
    d->sheetFormatProperties.outlineLevelRow = maximumOutlineLevel+1;
    return true;
}

bool Worksheet::groupColumns(const CellRange &range, bool collapsed)
{
    if (!range.isValid())
        return false;

    return groupColumns(range.firstColumn(), range.lastColumn(), collapsed);
}

bool Worksheet::groupColumns(int colFirst, int colLast, bool collapsed)
{
    Q_D(Worksheet);

    //calculate current max outline level
    int maximumOutlineLevel = 0;
    for (int i = colFirst; i <= colLast; ++i) {
        if (d->colsInfo.contains(i))
            maximumOutlineLevel = std::max(d->colsInfo.value(i).outlineLevel.value_or(0),
                                           maximumOutlineLevel);
    }
    if (maximumOutlineLevel >= 7) return false;

    if (!d->addColumnToDimensions(colFirst)) return false;
    if (!d->addColumnToDimensions(colLast)) return false;

    for (int i = colFirst; i <= colLast; ++i) {
        auto &val = d->colsInfo[i];
        val.outlineLevel = val.outlineLevel.value_or(0) + 1;
        if (collapsed) val.hidden = true;
        else {
            //we need it for Excel to not hide grouped columns
            if (!val.width.has_value()) val.width = d->sheetFormatProperties.defaultColumnWidth();
            if (!val.hidden.has_value()) val.hidden = false;
        }
    }
    if (collapsed) {
        auto &val = d->colsInfo[colLast+1];
        val.collapsed = true;
        if (!val.width.has_value()) val.width = d->sheetFormatProperties.defaultColumnWidth();
        if (!val.hidden.has_value()) val.hidden = false;
    }
    d->sheetFormatProperties.outlineLevelCol = maximumOutlineLevel+1;
    return true;
}

CellRange Worksheet::dimension() const
{
    Q_D(const Worksheet);
    return d->dimension;
}

std::optional<bool> Worksheet::isFormatConditionsCalculationEnabled() const
{
    Q_D(const Worksheet);
    return d->sheetProperties.enableFormatConditionsCalculation;
}

void Worksheet::setFormatConditionsCalculationEnabled(bool enabled)
{
    Q_D(Worksheet);
    d->sheetProperties.enableFormatConditionsCalculation = enabled;
}
std::optional<bool> Worksheet::isSyncedHorizontal() const
{
    Q_D(const Worksheet);
    return d->sheetProperties.syncHorizontal;
}

void Worksheet::setSyncedHorizontal(bool sync)
{
    Q_D(Worksheet);
    d->sheetProperties.syncHorizontal = sync;
}

std::optional<bool> Worksheet::isSyncedVertical() const
{
    Q_D(const Worksheet);
    return d->sheetProperties.syncVertical;
}

void Worksheet::setSyncedVertical(bool sync)
{
    Q_D(Worksheet);
    d->sheetProperties.syncVertical = sync;
}

CellReference Worksheet::topLeftAnchor() const
{
    Q_D(const Worksheet);
    return d->sheetProperties.syncRef;
}

void Worksheet::setTopLeftAnchor(const CellReference &ref)
{
    Q_D(Worksheet);
    d->sheetProperties.syncRef = ref;
}

std::optional<bool> Worksheet::showAutoPageBreaks() const
{
    Q_D(const Worksheet);
    return d->sheetProperties.autoPageBreaks;
}

void Worksheet::setShowAutoPageBreaks(bool show)
{
    Q_D(Worksheet);
    d->sheetProperties.autoPageBreaks = show;
}

std::optional<bool> Worksheet::fitToPage() const
{
    Q_D(const Worksheet);
    return d->sheetProperties.fitToPage;
}

void Worksheet::setFitToPage(bool value)
{
    Q_D(Worksheet);
    d->sheetProperties.fitToPage = value;
}

std::optional<bool> Worksheet::thickBottomBorder() const
{
    Q_D(const Worksheet);
    return d->sheetFormatProperties.thickBottom;
}

void Worksheet::setThickBottomBorder(bool thick)
{
    Q_D(Worksheet);
    d->sheetFormatProperties.thickBottom = thick;
}

std::optional<bool> Worksheet::thickTopBorder() const
{
    Q_D(const Worksheet);
    return d->sheetFormatProperties.thickTop;
}

void Worksheet::setThickTopBorder(bool thick)
{
    Q_D(Worksheet);
    d->sheetFormatProperties.thickTop = thick;
}

std::optional<bool> Worksheet::rowsHiddenByDefault() const
{
    Q_D(const Worksheet);
    return d->sheetFormatProperties.zeroHeight;
}

void Worksheet::setRowsHiddenByDefault(bool hidden)
{
    Q_D(Worksheet);
    d->sheetFormatProperties.zeroHeight = hidden;
}

double Worksheet::defaultRowHeight() const
{
    Q_D(const Worksheet);
    return d->sheetFormatProperties.defaultRowHeight;
}

void Worksheet::setDefaultRowHeight(double height)
{
    Q_D(Worksheet);
    d->sheetFormatProperties.defaultRowHeight = height;
    if (d->sheetFormatProperties.defaultRowHeight != XLSX_DEFAULT_ROW_HEIGHT)
        d->sheetFormatProperties.customHeight = true;
}

double Worksheet::defaultColumnWidth() const
{
    Q_D(const Worksheet);
    return d->sheetFormatProperties.defaultColumnWidth();
}

void Worksheet::setDefaultColumnWidth(double width)
{
    Q_D(Worksheet);
    d->sheetFormatProperties.defaultColWidth = width;
}

void WorksheetPrivate::loadXmlSheetData(QXmlStreamReader &reader)
{
    Q_ASSERT(reader.name() == QLatin1String("sheetData"));

    const auto &name = reader.name();

    //since row numbers are optional, we need to track the current row
    int currentRow = 0;
    while (!reader.atEnd())    {
        auto token = reader.readNext();
        if (token == QXmlStreamReader::StartElement) {
            const auto &a = reader.attributes();

            if (reader.name() == QLatin1String("row")) {
                currentRow++;
                QSharedPointer<XlsxRowInfo> info(new XlsxRowInfo);
                if (a.hasAttribute(QLatin1String("customFormat")) &&
                    a.hasAttribute(QLatin1String("s"))) {
                    int idx = a.value(QLatin1String("s")).toInt();
                    info->format = workbook->styles()->xfFormat(idx);
                }
                bool customHeight;
                parseAttributeBool(a, QLatin1String("customHeight"), customHeight);
                if (customHeight) info->height = sheetFormatProperties.defaultRowHeight;
                parseAttributeDouble(a, QLatin1String("ht"), info->height);
                parseAttributeBool(a, QLatin1String("hidden"), info->hidden);
                parseAttributeBool(a, QLatin1String("collapsed"), info->collapsed);
                parseAttributeInt(a, QLatin1String("outlineLevel"), info->outlineLevel);

                //"r" is optional too.
                if (a.hasAttribute(QLatin1String("r")))
                    currentRow = a.value(QLatin1String("r")).toInt();
                if (info->isValid())
                    rowsInfo[currentRow] = info;
            }
            else if (reader.name() == QLatin1String("c"))  //Cell
                loadXmlCell(reader);
        }
        else if (token == QXmlStreamReader::EndElement && reader.name() == name)
            break;
    }
}

void WorksheetPrivate::loadXmlCell(QXmlStreamReader &reader)
{
    Q_Q(Worksheet);

    const auto &name = reader.name();
    const auto &a = reader.attributes();

    CellReference pos(a.value(QLatin1String("r")).toString());

    //get format
    Format format;
    qint32 styleIndex = -1;
    if (a.hasAttribute(QLatin1String("s"))) {// Style (defined in the styles.xml file)
        //"s" == style index
        styleIndex = a.value(QLatin1String("s")).toInt();
        format = workbook->styles()->xfFormat(styleIndex);
    }

    auto cellType = Cell::Type::Custom;

    if (a.hasAttribute(QLatin1String("t"))) {// Type
        auto typeString = a.value(QLatin1String("t")).toString();
        Cell::fromString(typeString, cellType);
    }

    if (Cell::isDateType(cellType, format)) cellType = Cell::Type::Date;

    // create a heap of new cell
    auto cell = std::make_shared<Cell>(QVariant{}, cellType, format, q, styleIndex);

    while (!reader.atEnd())    {
        auto token = reader.readNext();
        if (token == QXmlStreamReader::StartElement) {
            if (reader.name() == QLatin1String("f")) {// formula
                CellFormula formula;
                formula.loadFromXml(reader);
                if (formula.type().value_or(CellFormula::Type::Normal) == CellFormula::Type::Shared
                    && !formula.text().isEmpty()) {
                    int si = formula.sharedIndex().value_or(-1);
                    sharedFormulaMap[si] = formula;
                }
                cell->setFormula(formula);
            }
            else if (reader.name() == QLatin1String("v")) {// Value
                QString value = reader.readElementText();
                if (cellType == Cell::Type::SharedString) {
                    int sst_idx = value.toInt();
                    sharedStrings()->incRefByStringIndex(sst_idx);
                    RichString rs = sharedStrings()->getSharedString(sst_idx);
                    QString strPlainString = rs.toPlainString();
                    cell->setValue(strPlainString);
                    if (rs.isRichString())
                        cell->setRichString(rs);
                }
                else if (cellType == Cell::Type::Number) {
                    cell->setValue(value.toDouble());
                }
                else if (cellType == Cell::Type::Boolean) {
                    cell->setValue(fromST_Boolean(value));
                }
                else  if (cellType == Cell::Type::Date) {
                    // [dev54] DateType

                    double dValue = value.toDouble(); // days from 1900(or 1904)
                    cell->setValue(dValue); // dev67
                }
                else {
                    // ELSE type
                    cell->setValue(value);
                }
            }
            else if (reader.name() == QLatin1String("is")) {
                //TODO: add rich text read support
                cell->setValue(reader.readElementText());
            }
            else if (reader.name() == QLatin1String("extLst"))
                reader.skipCurrentElement();
        }
        else if (token == QXmlStreamReader::EndElement && reader.name() == name)
            break;
    }
    cellTable[pos.row()][pos.column()] = cell;
}

void WorksheetPrivate::loadXmlColumnsInfo(QXmlStreamReader &reader)
{
    Q_ASSERT(reader.name() == QLatin1String("cols"));

    while (!reader.atEnd()) {
        auto token = reader.readNext();
        if (token == QXmlStreamReader::StartElement) {
            if (reader.name() == QLatin1String("col")) {
                XlsxColumnInfo info;
                const auto &a = reader.attributes();
                int firstColumn;
                int lastColumn;

                parseAttributeInt(a, QLatin1String("min"), firstColumn);
                parseAttributeInt(a, QLatin1String("max"), lastColumn);

                //Flag indicating that the column width for the affected column(s) is different from the
                // default or has been manually set
                bool customWidth = false;
                parseAttributeBool(a, QLatin1String("customWidth"), customWidth);
                if (customWidth) info.width = sheetFormatProperties.defaultColumnWidth();

                //Note, node may have "width" without "customWidth"
                // [dev54]
                parseAttributeDouble(a, QLatin1String("width"), info.width);
                parseAttributeBool(a, QLatin1String("hidden"), info.hidden);
                parseAttributeBool(a, QLatin1String("collapsed"), info.collapsed);

                if (a.hasAttribute(QLatin1String("style"))) {
                    int idx = a.value(QLatin1String("style")).toInt();
                    info.format = workbook->styles()->xfFormat(idx);
                }
                parseAttributeInt(a, QLatin1String("outlineLevel"), info.outlineLevel);
                // qDebug() << "[debug] " << __FUNCTION__ << min << max << info->width << hasWidth;
                for (int i=firstColumn; i<=lastColumn; ++i) {
                    colsInfo.insert(i, info);
                }
            }
        }
        else if (token == QXmlStreamReader::EndElement && reader.name() == QLatin1String("cols"))
            break;
    }
}

void WorksheetPrivate::loadXmlMergeCells(QXmlStreamReader &reader)
{
    // issue #173 https://github.com/QtExcel/QXlsx/issues/173
    auto a = reader.attributes();
    int count = a.value(QLatin1String("count")).toInt();

    auto name = reader.name();
    while (!reader.atEnd()) {
        auto token = reader.readNext();
        if (token == QXmlStreamReader::StartElement) {
            if (reader.name() == QLatin1String("mergeCell")) {
                QString rangeStr = reader.attributes().value(QLatin1String("ref")).toString();
                merges.append(CellRange(rangeStr));
            }
        }
        else if (token == QXmlStreamReader::EndElement && reader.name() == name)
            break;
    }

    if (count != merges.size())
        qWarning("WorksheetPrivate::loadXmlMergeCells: read merge cells error");
}

void WorksheetPrivate::loadXmlDataValidations(QXmlStreamReader &reader)
{
    Q_ASSERT(reader.name() == QLatin1String("dataValidations"));
    QXmlStreamAttributes attributes = reader.attributes();
    int count = attributes.value(QLatin1String("count")).toInt();
    parseAttributeBool(attributes, QLatin1String("disablePrompts"), disableValidationPrompts);
    parseAttributeInt(attributes, QLatin1String("xWindow"), dataValidationXWindow);
    parseAttributeInt(attributes, QLatin1String("yWindow"), dataValidationYWindow);

    while (!reader.atEnd() && !(reader.name() == QLatin1String("dataValidations")
            && reader.tokenType() == QXmlStreamReader::EndElement)) {
        reader.readNextStartElement();
        if (reader.tokenType() == QXmlStreamReader::StartElement
                && reader.name() == QLatin1String("dataValidation")) {
            DataValidation dv;
            dv.read(reader);
            dataValidationsList.append(dv);
        }
    }

    if (dataValidationsList.size() != count)
        qDebug("read data validation error");
}

void WorksheetPrivate::loadXmlSheetViews(QXmlStreamReader &reader)
{
    const auto &name = reader.name();

    while (!reader.atEnd()) {
        const auto token = reader.readNext();
        if (token == QXmlStreamReader::StartElement) {
            if (reader.name() == QLatin1String("sheetView")) {
                SheetView view;
                view.read(reader);
                sheetViews << view;
            }
        }
        else if (token == QXmlStreamReader::EndElement && reader.name() == name)
            break;
    }
}

double WorksheetPrivate::calculateColWidth(int characters)
{
//    Default column width measured as the number of characters of the maximum digit width
//    of the normal style's font.

//    If the user has not set this manually, then it can be calculated:
//    defaultColWidth = baseColumnWidth + {margin padding (2 pixels on each side, totalling
//    4 pixels)} + {gridline (1pixel)}

    //Calibri 11 has digits width of 7px at 96dpi, so margins width will be 5/7 characters
    return characters + 5.0/7.0;
}

void WorksheetPrivate::loadXmlHyperlinks(QXmlStreamReader &reader)
{
    Q_ASSERT(reader.name() == QLatin1String("hyperlinks"));

    while (!reader.atEnd() && !(reader.name() == QLatin1String("hyperlinks")
            && reader.tokenType() == QXmlStreamReader::EndElement)) {
        reader.readNextStartElement();
        if (reader.tokenType() == QXmlStreamReader::StartElement && reader.name() == QLatin1String("hyperlink")) {
            QXmlStreamAttributes attrs = reader.attributes();
            CellReference pos(attrs.value(QLatin1String("ref")).toString());
            if (pos.isValid()) { //Valid
                QSharedPointer<XlsxHyperlinkData> link(new XlsxHyperlinkData);
                link->display = attrs.value(QLatin1String("display")).toString();
                link->tooltip = attrs.value(QLatin1String("tooltip")).toString();
                link->location = attrs.value(QLatin1String("location")).toString();

                if (attrs.hasAttribute(QLatin1String("r:id"))) {
                    link->linkType = XlsxHyperlinkData::External;
                    XlsxRelationship ship = relationships->getRelationshipById(attrs.value(QLatin1String("r:id")).toString());
                    link->target = ship.target;
                } else {
                    link->linkType = XlsxHyperlinkData::Internal;
                }

                urlTable[pos.row()][pos.column()] = link;
            }
        }
    }
}

bool Worksheet::loadFromXmlFile(QIODevice *device)
{
    Q_D(Worksheet);

    QXmlStreamReader reader(device);
    while (!reader.atEnd()) {
        auto token = reader.readNext();
        if (token == QXmlStreamReader::StartElement) {
            const auto &a = reader.attributes();
            if (reader.name() == QLatin1String("sheetPr"))
                d->sheetProperties.read(reader);
            else if (reader.name() == QLatin1String("dimension")) {
                QString range = a.value(QLatin1String("ref")).toString();
                d->dimension = CellRange(range);
            }
            else if (reader.name() == QLatin1String("sheetViews"))
                d->loadXmlSheetViews(reader);
            else if (reader.name() == QLatin1String("sheetFormatPr"))
                d->sheetFormatProperties.read(reader);
            else if (reader.name() == QLatin1String("cols"))
                d->loadXmlColumnsInfo(reader);
            else if (reader.name() == QLatin1String("sheetData"))
                d->loadXmlSheetData(reader);
            else if (reader.name() == QLatin1String("sheetProtection")) {
                SheetProtection s;
                s.read(reader);
                d->sheetProtection = s;
            }
            else if (reader.name() == QLatin1String("mergeCells"))
                d->loadXmlMergeCells(reader);
            else if (reader.name() == QLatin1String("dataValidations"))
                d->loadXmlDataValidations(reader);
            else if (reader.name() == QLatin1String("conditionalFormatting")) {
                ConditionalFormatting cf;
                cf.loadFromXml(reader, workbook()->styles());
                d->conditionalFormattingList.append(cf);
            }
            else if (reader.name() == QLatin1String("hyperlinks"))
                d->loadXmlHyperlinks(reader);
            else if(reader.name() == QLatin1String("pageSetup"))
                d->pageSetup.read(reader);
            else if(reader.name() == QLatin1String("pageMargins"))
                d->pageMargins.read(reader);
            else if(reader.name() == QLatin1String("headerFooter"))
                d->headerFooter.read(reader);
            else if (reader.name() == QLatin1String("autoFilter"))
                d->autofilter.read(reader);
            else if (reader.name() == QLatin1String("sortState"))
                d->sortState.read(reader);
            else if (reader.name() == QLatin1String("drawing")) {
                QString rId = a.value(QLatin1String("r:id")).toString();
                QString name = d->relationships->getRelationshipById(rId).target;

                const auto parts = splitPath(filePath());
                QString path = QDir::cleanPath(parts.first() + QLatin1String("/") + name);

                d->drawing = std::make_shared<Drawing>(this, F_LoadFromExists);
                d->drawing->setFilePath(path);
            }
            else if (reader.name() == QLatin1String("picture")) {
                QString rId = a.value(QLatin1String("r:id")).toString();
                QString name = d->relationships->getRelationshipById(rId).target;

                const auto parts = splitPath(filePath());
                QString path = QDir::cleanPath(parts.first() + QLatin1String("/") + name);

                bool exist = false;
                const auto mfs = d->workbook->mediaFiles();
                for (const auto &mf : mfs) {
                    if (auto media = mf.lock(); media->fileName() == path) {
                        //already exist
                        exist = true;
                        d->pictureFile = media;
                        break;
                    }
                }
                if (!exist) {
                    d->pictureFile = QSharedPointer<MediaFile>(new MediaFile(path));
                    d->workbook->addMediaFile(d->pictureFile, true);
                }
            }
            else if (reader.name() == QLatin1String("extLst"))
                d->extLst.read(reader);
            else if (reader.name() == QLatin1String("sheetCalcPr")) {
                parseAttributeBool(reader.attributes(), QLatin1String("fullCalcOnLoad"), d->fullCalcOnLoad);
            }
        }
    }

    d->validateDimension();
    return true;
}

/*
 *  Documents imported from Google Docs do not contain dimension data.
 */
void WorksheetPrivate::validateDimension()
{
    if (dimension.isValid() || cellTable.isEmpty())
        return;

    const auto firstRow = cellTable.constBegin().key();

    const auto lastRow = (--cellTable.constEnd()).key();

    int firstColumn = -1;
    int lastColumn = -1;

    for (auto&& it = cellTable.constBegin(); it != cellTable.constEnd(); ++it) {
        Q_ASSERT(!it.value().isEmpty());

        if (firstColumn == -1 || it.value().constBegin().key() < firstColumn)
            firstColumn = it.value().constBegin().key();

        if (lastColumn == -1 || (--it.value().constEnd()).key() > lastColumn)
            lastColumn = (--it.value().constEnd()).key();
    }

    CellRange cr(firstRow, firstColumn, lastRow, lastColumn);

    if (cr.isValid())
        dimension = cr;
}

/*!
 * \internal
 *  Unit test can use this member to get sharedString object.
 */
SharedStrings *WorksheetPrivate::sharedStrings() const
{
    return workbook->sharedStrings();
}

bool Worksheet::autosizeColumnsWidth(int firstColumn, int lastColumn)
{
    CellRange r(1, firstColumn, INT_MAX, lastColumn);
    return autosizeColumnsWidth(r);
}

bool Worksheet::autosizeColumnsWidth(const CellRange &range)
{
    bool result = false;
    if (range.isValid()) {
        auto columnWidths = getMaximumColumnWidths(range.firstRow(), range.lastRow());

        for (auto it = columnWidths.constBegin(); it != columnWidths.constEnd(); ++it) {
            if (it.key() >= range.firstColumn() && it.key() <= range.lastColumn())
                result |= setColumnWidth(it.key(), it.key(), it.value());
        }
    }

    return result;
}

bool Worksheet::autosizeColumnsWidth()
{
    return autosizeColumnsWidth(1, INT_MAX);
}
bool Worksheet::autosizeColumnWidth(int column)
{
    return autosizeColumnsWidth(column, column);
}

std::optional<bool> Worksheet::fullCalculationOnLoad() const
{
    Q_D(const Worksheet);
    return d->fullCalcOnLoad;
}

void Worksheet::setFullCalculationOnLoad(bool value)
{
    Q_D(Worksheet);
    d->fullCalcOnLoad = value;
}

QMap<int, double> Worksheet::getMaximumColumnWidths(int firstRow, int lastRow)
{
    Q_D(const Worksheet);

    const int defaultPtSize = 11;    //Calibri 11

    QMap<int, double> colWidth;

    for (auto row = d->cellTable.constBegin(); row != d->cellTable.constEnd(); ++row) {
        int rowIndex = row.key();
        const auto &rowValue = row.value();

        for (auto col = rowValue.constBegin(); col != rowValue.constEnd(); ++col) {
            int colIndex = col.key();
            std::shared_ptr<Cell> ptrCell = col.value(); // value

            auto fs = ptrCell->format().fontSize();
            if (fs <= 0) fs = defaultPtSize;

            QString str = read(rowIndex, colIndex).toString();
            double w = str.length() * fs / defaultPtSize + 1; // width not perfect, but works reasonably well

            if (rowIndex >= firstRow && rowIndex <= lastRow) {
                if (w > colWidth.value(colIndex))
                    colWidth.insert(colIndex, w);
            }
        }
    }

    return colWidth;
}

bool SheetProperties::isValid() const
{
    if (tabColor.isValid()) return true;
    if (applyStyles.has_value()) return true;
    if (summaryBelow.has_value()) return true;
    if (summaryRight.has_value()) return true;
    if (showOutlineSymbols.has_value()) return true;
    if (autoPageBreaks.has_value()) return true;
    if (fitToPage.has_value()) return true;
    if (syncHorizontal.has_value()) return true;
    if (syncVertical.has_value()) return true;
    if (syncRef.isValid()) return true;
    if (transitionEvaluation.has_value()) return true;
    if (transitionEntry.has_value()) return true;
    if (published.has_value()) return true;
    if (!codeName.isEmpty()) return true;
    if (filterMode.has_value()) return true;
    if (enableFormatConditionsCalculation.has_value()) return true;

    return false;
}

void SheetProperties::read(QXmlStreamReader &reader)
{
    const auto &name = reader.name();

    const auto &a = reader.attributes();
    parseAttributeBool(a, QLatin1String("syncHorizontal"), syncHorizontal);
    parseAttributeBool(a, QLatin1String("syncVertical"), syncVertical);
    if (a.hasAttribute(QLatin1String("syncRef")))
        syncRef = CellReference::fromString(a.value(QLatin1String("syncRef")).toString());
    parseAttributeBool(a, QLatin1String("transitionEvaluation"), transitionEvaluation);
    parseAttributeBool(a, QLatin1String("transitionEntry"), transitionEntry);
    parseAttributeBool(a, QLatin1String("published"), published);
    parseAttributeString(a, QLatin1String("codeName"), codeName);
    parseAttributeBool(a, QLatin1String("filterMode"), filterMode);
    parseAttributeBool(a, QLatin1String("enableFormatConditionsCalculation"), enableFormatConditionsCalculation);

    while (!reader.atEnd()) {
        const auto token = reader.readNext();
        if (token == QXmlStreamReader::StartElement) {
            if (reader.name() == QLatin1String("tabColor"))
                tabColor.read(reader);
            else if (reader.name() == QLatin1String("outlinePr")) {
                const auto &a1 = reader.attributes();
                parseAttributeBool(a1, QLatin1String("applyStyles"), applyStyles);
                parseAttributeBool(a1, QLatin1String("summaryBelow"), summaryBelow);
                parseAttributeBool(a1, QLatin1String("summaryRight"), summaryRight);
                parseAttributeBool(a1, QLatin1String("showOutlineSymbols"), showOutlineSymbols);
            }
            else if (reader.name() == QLatin1String("pageSetUpPr")) {
                const auto &a1 = reader.attributes();
                parseAttributeBool(a1, QLatin1String("autoPageBreaks"), autoPageBreaks);
                parseAttributeBool(a1, QLatin1String("fitToPage"), fitToPage);
            }
        }
        else if (token == QXmlStreamReader::EndElement && reader.name() == name)
            break;
    }
}

void SheetProperties::write(QXmlStreamWriter &writer, const QLatin1String &name) const
{
    if (!isValid()) return;

    writer.writeStartElement(name);
    writeAttribute(writer, QLatin1String("syncHorizontal"), syncHorizontal);
    writeAttribute(writer, QLatin1String("syncVertical"), syncVertical);
    if (syncRef.isValid())
        writeAttribute(writer, QLatin1String("syncRef"), syncRef.toString());
    writeAttribute(writer, QLatin1String("transitionEvaluation"), transitionEvaluation);
    writeAttribute(writer, QLatin1String("transitionEntry"), transitionEntry);
    writeAttribute(writer, QLatin1String("published"), published);
    writeAttribute(writer, QLatin1String("codeName"), codeName);
    writeAttribute(writer, QLatin1String("filterMode"), filterMode);
    writeAttribute(writer, QLatin1String("enableFormatConditionsCalculation"), enableFormatConditionsCalculation);
    if (tabColor.isValid()) tabColor.write(writer, QLatin1String("tabColor"));
    if (applyStyles.has_value() || summaryBelow.has_value() || summaryRight.has_value() ||
        showOutlineSymbols.has_value()) {
        writer.writeEmptyElement(QLatin1String("outlinePr"));
        writeAttribute(writer, QLatin1String("applyStyles"), applyStyles);
        writeAttribute(writer, QLatin1String("summaryBelow"), summaryBelow);
        writeAttribute(writer, QLatin1String("summaryRight"), summaryRight);
        writeAttribute(writer, QLatin1String("showOutlineSymbols"), showOutlineSymbols);
    }
    if (autoPageBreaks.has_value() || fitToPage.has_value()) {
        writer.writeEmptyElement(QLatin1String("pageSetUpPr"));
        writeAttribute(writer, QLatin1String("autoPageBreaks"), autoPageBreaks);
        writeAttribute(writer, QLatin1String("fitToPage"), fitToPage);
    }
    writer.writeEndElement();
}

double SheetFormatProperties::defaultColumnWidth() const
{
    return defaultColWidth.value_or(double(baseColWidth) + 5.0/7.0);
}

bool SheetFormatProperties::isValid() const
{
    //baseColWidth was made non-optional, so we test it
    if (baseColWidth != 8) return true;
    if (customHeight.has_value()) return true; // = false;
    if (defaultColWidth.has_value()) return true; // = 8.43f;
    //sheetFormatPr is optional in worksheet, but defaultRowHeight is not,
    //so we ignore it: if any other parameter is present, then we write sheetFormatPr.
    //if (defaultRowHeight.has_value()) return true; // 14.4
    if (outlineLevelCol.has_value()) return true; // = 0;
    if (outlineLevelRow.has_value()) return true; // = 0;
    if (thickBottom.has_value()) return true; // = false;
    if (thickTop.has_value()) return true; //= false;
    if (zeroHeight.has_value()) return true; // = false;

    return false;
}

void SheetFormatProperties::read(QXmlStreamReader &reader)
{
    Q_ASSERT(reader.name() == QLatin1String("sheetFormatPr"));

    const QXmlStreamAttributes a = reader.attributes();

    parseAttributeInt(a, QLatin1String("baseColWidth"), baseColWidth);
    parseAttributeBool(a, QLatin1String("customHeight"), customHeight);
    if (a.hasAttribute(QLatin1String("defaultColWidth")))
        parseAttributeDouble(a, QLatin1String("defaultColWidth"), defaultColWidth);
    else if (baseColWidth != 8)
        defaultColWidth = double(baseColWidth) + 5.0/7.0;
    parseAttributeDouble(a, QLatin1String("defaultRowHeight"), defaultRowHeight);
    parseAttributeInt(a, QLatin1String("outlineLevelCol"), outlineLevelCol);
    parseAttributeInt(a, QLatin1String("outlineLevelRow"), outlineLevelRow);
    parseAttributeBool(a, QLatin1String("thickBottom"), thickBottom);
    parseAttributeBool(a, QLatin1String("thickTop"), thickTop);
    parseAttributeBool(a, QLatin1String("zeroHeight"), zeroHeight);
}

void SheetFormatProperties::write(QXmlStreamWriter &writer, const QLatin1String &name) const
{
//    if (!isValid()) return;

    writer.writeEmptyElement(name);
    if (baseColWidth != 8) writeAttribute(writer, QLatin1String("baseColWidth"), baseColWidth);
    writeAttribute(writer, QLatin1String("defaultColWidth"), defaultColWidth);
    writeAttribute(writer, QLatin1String("defaultRowHeight"), defaultRowHeight);
    writeAttribute(writer, QLatin1String("customHeight"), customHeight);
    writeAttribute(writer, QLatin1String("zeroHeight"), zeroHeight);
    writeAttribute(writer, QLatin1String("thickTop"), thickTop);
    writeAttribute(writer, QLatin1String("thickBottom"), thickBottom);
    writeAttribute(writer, QLatin1String("outlineLevelRow"), outlineLevelRow);
    writeAttribute(writer, QLatin1String("outlineLevelCol"), outlineLevelCol);
}

}
