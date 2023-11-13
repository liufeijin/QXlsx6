#include "xlsxpagesetup.h"
#include "xlsxutility_p.h"

namespace QXlsx {

class PageSetupPrivate : public QSharedData
{
public:
    PageSetupPrivate();
    PageSetupPrivate(const PageSetupPrivate& other);
    ~PageSetupPrivate();
    bool operator==(const PageSetupPrivate &other) const;
    std::optional<PageSetup::PaperSize> paperSize;
    QString paperWidth;
    QString paperHeight;
    std::optional<int> scale;
    std::optional<int> firstPageNumber;
    std::optional<int> fitToWidth;
    std::optional<int> fitToHeight;
    std::optional<PageSetup::PageOrder> pageOrder;
    std::optional<PageSetup::Orientation> orientation;
    std::optional<bool> usePrinterDefaults;
    std::optional<bool> blackAndWhite;
    std::optional<bool> draft;
    std::optional<bool> useFirstPageNumber;
    std::optional<int> horizontalDpi;
    std::optional<int> verticalDpi;
    std::optional<int> copies;
    QString printerId; //TODO: adding printer
    std::optional<PageSetup::PrintError> errors;
    std::optional<PageSetup::CellComments> cellComments;
    std::optional<bool> printGridLines;
    std::optional<bool> printHeadings;
    std::optional<bool> printHorizontalCentered;
    std::optional<bool> printVerticalCentered;
};

PageSetup::PageSetup()
{

}

PageSetup::PageSetup(const PageSetup &other) : d{other.d}
{

}

PageSetup &PageSetup::operator=(const PageSetup &other)
{
    if (*this != other) d = other.d;
    return *this;
}

PageSetup::~PageSetup()
{

}

bool PageSetup::operator==(const PageSetup &other) const
{
    if (d == other.d) return true;
    if (!d || !other.d) return false;
    return *this->d.constData() == *other.d.constData();
}

bool PageSetup::operator!=(const PageSetup &other) const
{
    return !operator==(other);
}

PageSetup::operator QVariant() const
{
    const auto& cref
#if QT_VERSION >= 0x060000 // Qt 6.0 or over
        = QMetaType::fromType<PageSetup>();
#else
        = qMetaTypeId<PageSetup>() ;
#endif
    return QVariant(cref, this);
}

bool PageSetup::isValid() const
{
    if (d) return true;
    return false;
}

std::optional<PageSetup::PaperSize> PageSetup::paperSize() const
{
    if (d) return d->paperSize;
    return {};
}

void PageSetup::setPaperSize(PaperSize paperSize)
{
    if (!d) d = new PageSetupPrivate{};
    d->paperSize = paperSize;
}

QString PageSetup::paperWidth() const
{
    if (d) return d->paperWidth;
    return {};
}

void PageSetup::setPaperWidth(const QString &paperWidth)
{
    if (!d) d = new PageSetupPrivate{};
    d->paperWidth = paperWidth;
}

QString PageSetup::paperHeight() const
{
    if (d) return d->paperHeight;
    return {};
}

void PageSetup::setPaperHeight(const QString &paperHeight)
{
    if (!d) d = new PageSetupPrivate{};
    d->paperHeight = paperHeight;
}

std::optional<int> PageSetup::scale() const
{
    if (d) return d->scale;
    return {};
}

void PageSetup::setScale(int scale)
{
    if (!d) d = new PageSetupPrivate{};
    d->scale = scale;
}

std::optional<int> PageSetup::firstPageNumber() const
{
    if (d) return d->firstPageNumber;
    return {};
}

void PageSetup::setFirstPageNumber(int firstPageNumber)
{
    if (!d) d = new PageSetupPrivate{};
    d->firstPageNumber = firstPageNumber;
}

std::optional<int> PageSetup::fitToWidth() const
{
    if (d) return d->fitToWidth;
    return {};
}

void PageSetup::setFitToWidth(int pages)
{
    if (!d) d = new PageSetupPrivate{};
    d->fitToWidth = pages;
}

std::optional<int> PageSetup::fitToHeight() const
{
    if (d) return d->fitToHeight;
    return {};
}

void PageSetup::setFitToHeight(int pages)
{
    if (!d) d = new PageSetupPrivate{};
    d->fitToHeight = pages;
}

std::optional<PageSetup::PageOrder> PageSetup::pageOrder() const
{
    if (d) return d->pageOrder;
    return {};
}

void PageSetup::setPageOrder(PageOrder pageOrder)
{
    if (!d) d = new PageSetupPrivate{};
    d->pageOrder = pageOrder;
}

std::optional<PageSetup::Orientation> PageSetup::orientation() const
{
    if (d) return d->orientation;
    return {};
}

void PageSetup::setOrientation(Orientation orientation)
{
    if (!d) d = new PageSetupPrivate{};
    d->orientation = orientation;
}

std::optional<bool> PageSetup::usePrinterDefaults() const
{
    if (d) return d->usePrinterDefaults;
    return {};
}

void PageSetup::setUsePrinterDefaults(bool use)
{
    if (!d) d = new PageSetupPrivate{};
    d->usePrinterDefaults = use;
}

std::optional<bool> PageSetup::printBlackAndWhite() const
{
    if (d) return d->blackAndWhite;
    return {};
}

void PageSetup::setPrintBlackAndWhite(bool blackAndWhite)
{
    if (!d) d = new PageSetupPrivate{};
    d->blackAndWhite = blackAndWhite;
}

std::optional<bool> PageSetup::printDraft() const
{
    if (d) return d->draft;
    return {};
}

void PageSetup::setPrintDraft(bool draft)
{
    if (!d) d = new PageSetupPrivate{};
    d->draft = draft;
}

std::optional<bool> PageSetup::useFirstPageNumber() const
{
    if (d) return d->useFirstPageNumber;
    return {};
}

void PageSetup::setUseFirstPageNumber(bool use)
{
    if (!d) d = new PageSetupPrivate{};
    d->useFirstPageNumber = use;
}

std::optional<int> PageSetup::horizontalDpi() const
{
    if (d) return d->horizontalDpi;
    return {};
}

void PageSetup::setHorizontalDpi(int dpi)
{
    if (!d) d = new PageSetupPrivate{};
    d->horizontalDpi = dpi;
}

std::optional<int> PageSetup::verticalDpi() const
{
    if (d) return d->verticalDpi;
    return {};
}

void PageSetup::setVerticalDpi(int dpi)
{
    if (!d) d = new PageSetupPrivate{};
    d->verticalDpi = dpi;
}

std::optional<int> PageSetup::copies() const
{
    if (d) return d->copies;
    return {};
}

void PageSetup::setCopies(int copies)
{
    if (!d) d = new PageSetupPrivate{};
    d->copies = copies;
}

QString PageSetup::printerId() const
{
    if (d) return d->printerId;
    return {};
}

std::optional<PageSetup::PrintError> PageSetup::printErrors() const
{
    if (d) return d->errors;
    return {};
}

void PageSetup::setPrintErrors(PrintError errors)
{
    if (!d) d = new PageSetupPrivate{};
    d->errors = errors;
}

std::optional<PageSetup::CellComments> PageSetup::printCellComments() const
{
    if (d) return d->cellComments;
    return {};
}

void PageSetup::setPrintCellComments(CellComments cellComments)
{
    if (!d) d = new PageSetupPrivate{};
    d->cellComments = cellComments;
}

std::optional<bool> PageSetup::printGridLines() const
{
    if (d) return d->printGridLines;
    return {};
}

void PageSetup::setPrintGridLines(bool printGridLines)
{
    if (!d) d = new PageSetupPrivate{};
    d->printGridLines = printGridLines;
}

std::optional<bool> PageSetup::printHeadings() const
{
    if (d) return d->printHeadings;
    return {};
}

void PageSetup::setPrintHeadings(bool printHeadings)
{
    if (!d) d = new PageSetupPrivate{};
    d->printHeadings = printHeadings;
}

std::optional<bool> PageSetup::printHorizontalCentered() const
{
    if (d) return d->printHorizontalCentered;
    return {};
}

void PageSetup::setPrintHorizontalCentered(bool centered)
{
    if (!d) d = new PageSetupPrivate{};
    d->printHorizontalCentered = centered;
}

std::optional<bool> PageSetup::printVerticalCentered() const
{
    if (d) return d->printVerticalCentered;
    return {};
}

void PageSetup::setPrintVerticalCentered(bool centered)
{
    if (!d) d = new PageSetupPrivate{};
    d->printVerticalCentered = centered;
}

void PageSetup::writeWorksheet(QXmlStreamWriter &writer, const QString &name) const
{
    if (!d) return;
    if (name == QLatin1String("pageSetup")) {
        writer.writeEmptyElement(QLatin1String("pageSetup"));
        if (d->paperSize.has_value()) {
            int val = static_cast<int>(d->paperSize.value());
            writer.writeAttribute(QLatin1String("paperSize"), QString::number(val));
        }
        writeAttribute(writer, QLatin1String("paperHeight"), d->paperHeight);
        writeAttribute(writer, QLatin1String("paperWidth"), d->paperWidth);
        writeAttribute(writer, QLatin1String("scale"), d->scale);
        writeAttribute(writer, QLatin1String("firstPageNumber"), d->firstPageNumber);
        writeAttribute(writer, QLatin1String("fitToWidth"), d->fitToWidth);
        writeAttribute(writer, QLatin1String("fitToHeight"), d->fitToHeight);
        if (d->pageOrder.has_value()) {
            switch (d->pageOrder.value()) {
            case PageOrder::DownThenOver: writeAttribute(writer, QLatin1String("pageOrder"), "downThenOver"); break;
            case PageOrder::OverThenDown: writeAttribute(writer, QLatin1String("pageOrder"), "overThenDown"); break;
            }
        }
        if (d->orientation.has_value()) {
            switch (d->orientation.value()) {
            case Orientation::Default: writeAttribute(writer, QLatin1String("orientation"), "default"); break;
            case Orientation::Portrait: writeAttribute(writer, QLatin1String("orientation"), "portrait"); break;
            case Orientation::Landscape: writeAttribute(writer, QLatin1String("orientation"), "landscape"); break;
            }
        }
        writeAttribute(writer, QLatin1String("usePrinterDefaults"), d->usePrinterDefaults);
        writeAttribute(writer, QLatin1String("blackAndWhite"), d->blackAndWhite);
        writeAttribute(writer, QLatin1String("draft"), d->draft);
        if (d->cellComments.has_value()) {
            switch (d->cellComments.value()) {
            case CellComments::AtEnd: writeAttribute(writer, QLatin1String("cellComments"), "atEnd"); break;
            case CellComments::AsDisplayed: writeAttribute(writer, QLatin1String("cellComments"), "asDisplayed"); break;
            case CellComments::DoNotPrint: writeAttribute(writer, QLatin1String("cellComments"), "none"); break;
            }
        }
        writeAttribute(writer, QLatin1String("useFirstPageNumber"), d->useFirstPageNumber);
        if (d->errors.has_value()) {
            switch (d->errors.value()) {
            case PrintError::Dash: writeAttribute(writer, QLatin1String("errors"), "dash"); break;
            case PrintError::Displayed: writeAttribute(writer, QLatin1String("errors"), "displayed"); break;
            case PrintError::Blank: writeAttribute(writer, QLatin1String("errors"), "blank"); break;
            case PrintError::NotAvailable: writeAttribute(writer, QLatin1String("errors"), "NA"); break;
            }
        }
        writeAttribute(writer, QLatin1String("horizontalDpi"), d->horizontalDpi);
        writeAttribute(writer, QLatin1String("verticalDpi"), d->verticalDpi);
        writeAttribute(writer, QLatin1String("copies"), d->copies);
        writeAttribute(writer, QLatin1String("r:id"), d->printerId);
    }
    if (name == QLatin1String("printOptions")) {
        writer.writeEmptyElement(name);
        writeAttribute(writer, QLatin1String("horizontalCentered"), d->printHorizontalCentered);
        writeAttribute(writer, QLatin1String("verticalCentered"), d->printVerticalCentered);
        writeAttribute(writer, QLatin1String("headings"), d->printHeadings);
        if (d->printGridLines.value_or(false)) {
            writeAttribute(writer, QLatin1String("gridLines"), true);
            writeAttribute(writer, QLatin1String("gridLinesSet"), true);
        }

    }
}

void PageSetup::writeChartsheet(QXmlStreamWriter &writer) const
{
    if (!d) return;

    writer.writeEmptyElement(QLatin1String("pageSetup"));
    if (d->paperSize.has_value()) {
        int val = static_cast<int>(d->paperSize.value());
        writer.writeAttribute(QLatin1String("paperSize"), QString::number(val));
    }
    writeAttribute(writer, QLatin1String("paperHeight"), d->paperHeight);
    writeAttribute(writer, QLatin1String("paperWidth"), d->paperWidth);
    writeAttribute(writer, QLatin1String("firstPageNumber"), d->firstPageNumber);
    if (d->orientation.has_value()) {
        switch (d->orientation.value()) {
        case Orientation::Default: writeAttribute(writer, QLatin1String("orientation"), "default"); break;
        case Orientation::Portrait: writeAttribute(writer, QLatin1String("orientation"), "portrait"); break;
        case Orientation::Landscape: writeAttribute(writer, QLatin1String("orientation"), "landscape"); break;
        }
    }
    writeAttribute(writer, QLatin1String("usePrinterDefaults"), d->usePrinterDefaults);
    writeAttribute(writer, QLatin1String("blackAndWhite"), d->blackAndWhite);
    writeAttribute(writer, QLatin1String("draft"), d->draft);
    writeAttribute(writer, QLatin1String("useFirstPageNumber"), d->useFirstPageNumber);
    writeAttribute(writer, QLatin1String("horizontalDpi"), d->horizontalDpi);
    writeAttribute(writer, QLatin1String("verticalDpi"), d->verticalDpi);
    writeAttribute(writer, QLatin1String("copies"), d->copies);
    writeAttribute(writer, QLatin1String("r:id"), d->printerId);
}

void PageSetup::read(QXmlStreamReader &reader)
{
    if (!d) d = new PageSetupPrivate{};
    const auto &a = reader.attributes();

    readPaperSize(reader);
    parseAttributeString(a, QLatin1String("paperHeight"), d->paperHeight);
    parseAttributeString(a, QLatin1String("paperWidth"), d->paperWidth);
    parseAttributeInt(a, QLatin1String("scale"), d->scale);
    parseAttributeInt(a, QLatin1String("firstPageNumber"), d->firstPageNumber);
    parseAttributeInt(a, QLatin1String("fitToWidth"), d->fitToWidth);
    parseAttributeInt(a, QLatin1String("fitToHeight"), d->fitToHeight);
    if (a.hasAttribute(QLatin1String("pageOrder"))) {
        auto s = a.value(QLatin1String("pageOrder"));
        if (s == QLatin1String("downThenOver")) d->pageOrder = PageOrder::DownThenOver;
        if (s == QLatin1String("overThenDown")) d->pageOrder = PageOrder::OverThenDown;
    }
    if (a.hasAttribute(QLatin1String("orientation"))) {
        auto s = a.value(QLatin1String("orientation"));
        if (s == QLatin1String("default")) d->orientation = Orientation::Default;
        if (s == QLatin1String("portrait")) d->orientation = Orientation::Portrait;
        if (s == QLatin1String("landscape")) d->orientation = Orientation::Landscape;
    }
    parseAttributeBool(a, QLatin1String("usePrinterDefaults"), d->usePrinterDefaults);
    parseAttributeBool(a, QLatin1String("blackAndWhite"), d->blackAndWhite);
    parseAttributeBool(a, QLatin1String("draft"), d->draft);
    if (a.hasAttribute(QLatin1String("cellComments"))) {
        auto s = a.value(QLatin1String("cellComments"));
        if (s == QLatin1String("atEnd")) d->cellComments = CellComments::AtEnd;
        if (s == QLatin1String("asDisplayed")) d->cellComments = CellComments::AsDisplayed;
        if (s == QLatin1String("none")) d->cellComments = CellComments::DoNotPrint;
    }
    parseAttributeBool(a, QLatin1String("useFirstPageNumber"), d->useFirstPageNumber);
    if (a.hasAttribute(QLatin1String("errors"))) {
        auto s = a.value(QLatin1String("errors"));
        if (s == QLatin1String("dash")) d->errors = PrintError::Dash;
        if (s == QLatin1String("displayed")) d->errors = PrintError::Displayed;
        if (s == QLatin1String("blank")) d->errors = PrintError::Blank;
        if (s == QLatin1String("NA")) d->errors = PrintError::NotAvailable;
    }
    parseAttributeInt(a, QLatin1String("horizontalDpi"), d->horizontalDpi);
    parseAttributeInt(a, QLatin1String("verticalDpi"), d->verticalDpi);
    parseAttributeInt(a, QLatin1String("copies"), d->copies);
    parseAttributeString(a, QLatin1String("r:id"), d->printerId);

    //printOptions
    parseAttributeBool(a, QLatin1String("horizontalCentered"), d->printHorizontalCentered);
    parseAttributeBool(a, QLatin1String("verticalCentered"), d->printVerticalCentered);
    parseAttributeBool(a, QLatin1String("headings"), d->printHeadings);

    bool printGrid1 = false;
    parseAttributeBool(a, QLatin1String("gridLines"), printGrid1);
    bool printGrid2 = false;
    parseAttributeBool(a, QLatin1String("gridLinesSet"), printGrid2);
    if (printGrid1 && printGrid2) d->printGridLines = true;
}

void PageSetup::readPaperSize(QXmlStreamReader &reader)
{
    if (!d) d = new PageSetupPrivate{};
    if (reader.attributes().hasAttribute(QLatin1String("paperSize"))) {
        auto val = reader.attributes().value(QLatin1String("paperSize")).toInt();
        if ((val >=1 && val <= 47) || (val >=50 && val <= 118))
            d->paperSize = static_cast<PaperSize>(val);
        else
            d->paperSize = PaperSize::Unknown;
    }
}

PageSetupPrivate::PageSetupPrivate()
{

}

PageSetupPrivate::PageSetupPrivate(const PageSetupPrivate &other) : QSharedData(other),
    paperSize{other.paperSize},
    paperWidth{other.paperWidth},
    paperHeight{other.paperHeight},
    scale{other.scale},
    firstPageNumber{other.firstPageNumber},
    fitToWidth{other.fitToWidth},
    fitToHeight{other.fitToHeight},
    pageOrder{other.pageOrder},
    orientation{other.orientation},
    usePrinterDefaults{other.usePrinterDefaults},
    blackAndWhite{other.blackAndWhite},
    draft{other.draft},
    useFirstPageNumber{other.useFirstPageNumber},
    horizontalDpi{other.horizontalDpi},
    verticalDpi{other.verticalDpi},
    copies{other.copies},
    printerId{other.printerId},
    errors{other.errors},
    cellComments{other.cellComments},
    printGridLines{other.printGridLines},
    printHeadings{other.printHeadings},
    printHorizontalCentered{other.printHorizontalCentered},
    printVerticalCentered{other.printVerticalCentered}
{

}

PageSetupPrivate::~PageSetupPrivate()
{

}

bool PageSetupPrivate::operator==(const PageSetupPrivate &other) const
{
    if (paperSize != other.paperSize) return false;
    if (paperWidth != other.paperWidth) return false;
    if (paperHeight != other.paperHeight) return false;
    if (scale != other.scale) return false;
    if (firstPageNumber != other.firstPageNumber) return false;
    if (fitToWidth != other.fitToWidth) return false;
    if (fitToHeight != other.fitToHeight) return false;
    if (pageOrder != other.pageOrder) return false;
    if (orientation != other.orientation) return false;
    if (usePrinterDefaults != other.usePrinterDefaults) return false;
    if (blackAndWhite != other.blackAndWhite) return false;
    if (draft != other.draft) return false;
    if (useFirstPageNumber != other.useFirstPageNumber) return false;
    if (horizontalDpi != other.horizontalDpi) return false;
    if (verticalDpi != other.verticalDpi) return false;
    if (copies != other.copies) return false;
    if (printerId != other.printerId) return false;
    if (errors != other.errors) return false;
    if (cellComments != other.cellComments) return false;
    if (printGridLines != other.printGridLines) return false;
    if (printHeadings != other.printHeadings) return false;
    if (printHorizontalCentered != other.printHorizontalCentered) return false;
    if (printVerticalCentered != other.printVerticalCentered) return false;
    return true;
}

}

#ifndef QT_NO_DEBUG_STREAM
QDebug operator<<(QDebug dbg, const QXlsx::PageSetup &f)
{
    QDebugStateSaver saver(dbg);

    dbg.nospace() << "QXlsx::PageSetup(" ;
    //TODO:
    dbg.nospace() << ")";
    return dbg;
}
#endif
