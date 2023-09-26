// xlsxabstractsheet.cpp

#include <QtGlobal>
#include <QBuffer>
#include <QFileInfo>

#include "xlsxabstractsheet.h"
#include "xlsxabstractsheet_p.h"
#include "xlsxworkbook.h"
#include "xlsxutility_p.h"
#include <QDebug>

namespace QXlsx {

AbstractSheetPrivate::AbstractSheetPrivate(AbstractSheet *p, AbstractSheet::CreateFlag flag)
    : AbstractOOXmlFilePrivate(p, flag)
{
    type = AbstractSheet::Type::Worksheet;
    sheetState = AbstractSheet::Visibility::Visible;
}

AbstractSheetPrivate::~AbstractSheetPrivate()
{
}



/*!
 * \internal
 */
AbstractSheet::AbstractSheet(const QString &name, int id, Workbook *workbook, AbstractSheetPrivate *d) :
    AbstractOOXmlFile(d)
{
    d_func()->name = name;
    d_func()->id = id;
    d_func()->workbook = workbook;
}

QString AbstractSheet::name() const
{
    Q_D(const AbstractSheet);
    return d->name;
}

bool AbstractSheet::rename(const QString &newName)
{
    if (name() == newName)
        return false;
    return workbook()->renameSheet(name(), newName);
}

/*!
 * \internal
 */
void AbstractSheet::setName(const QString &sheetName)
{
    Q_D(AbstractSheet);
    d->name = sheetName;
}

/*!
 * Returns the type of the sheet.
 */
AbstractSheet::Type AbstractSheet::type() const
{
    Q_D(const AbstractSheet);
    return d->type;
}

/*!
 * \internal
 */
void AbstractSheet::setType(Type type)
{
    Q_D(AbstractSheet);
    d->type = type;
}

AbstractSheet::Visibility AbstractSheet::visibility() const
{
    Q_D(const AbstractSheet);
    return d->sheetState;
}

void AbstractSheet::setVisibility(Visibility visibility)
{
    Q_D(AbstractSheet);
    d->sheetState = visibility;
}

bool AbstractSheet::isHidden() const
{
    Q_D(const AbstractSheet);
    return d->sheetState != Visibility::Visible;
}

bool AbstractSheet::isVisible() const
{
    return !isHidden();
}

void AbstractSheet::setHidden(bool hidden)
{
    Q_D(AbstractSheet);
    if (!hidden) d->sheetState = Visibility::Visible;
    else if (d->sheetState == Visibility::Visible) d->sheetState = Visibility::Hidden;
}

void AbstractSheet::setVisible(bool visible)
{
    Q_D(AbstractSheet);
    if (visible) d->sheetState = Visibility::Visible;
    else if (d->sheetState == Visibility::Visible) d->sheetState = Visibility::Hidden;
}

HeaderFooter AbstractSheet::headerFooter() const
{
    Q_D(const AbstractSheet);
    return d->headerFooter;
}

HeaderFooter &AbstractSheet::headerFooter()
{
    Q_D(AbstractSheet);
    return d->headerFooter;
}

bool AbstractSheet::isPublished() const
{
    Q_D(const AbstractSheet);
    return d->sheetProperties.published.value_or(true);
}

void AbstractSheet::setPublished(bool published)
{
    Q_D(AbstractSheet);
    d->sheetProperties.published = published;
}

QString AbstractSheet::codeName() const
{
    Q_D(const AbstractSheet);
    return d->sheetProperties.codeName;
}

void AbstractSheet::setCodeName(const QString &codeName)
{
    Q_D(AbstractSheet);
    d->sheetProperties.codeName = codeName;
}

Color AbstractSheet::tabColor() const
{
    Q_D(const AbstractSheet);
    return d->sheetProperties.tabColor;
}

void AbstractSheet::setTabColor(const Color &color)
{
    Q_D(AbstractSheet);
    d->sheetProperties.tabColor = color;
}

void AbstractSheet::setTabColor(const QColor &color)
{
    Q_D(AbstractSheet);
    d->sheetProperties.tabColor = Color(color);
}

QMap<PageMargins::Position, double> AbstractSheet::pageMarginsInches() const
{
    Q_D(const AbstractSheet);
    return d->pageMargins.marginsInches();
}

QMap<PageMargins::Position, double> AbstractSheet::pageMarginsMm() const
{
    Q_D(const AbstractSheet);
    return d->pageMargins.marginsMm();
}

void AbstractSheet::setPageMarginsInches(double left, double top, double right, double bottom, double header, double footer)
{
    Q_D(AbstractSheet);
    d->pageMargins.setMarginsInches(left, top, right, bottom, header, footer);
}

void AbstractSheet::setPageMarginsMm(double left, double top, double right, double bottom, double header, double footer)
{
    Q_D(AbstractSheet);
    d->pageMargins.setMarginsMm(left, top, right, bottom, header, footer);
}

PageMargins AbstractSheet::pageMargins() const
{
    Q_D(const AbstractSheet);
    return d->pageMargins;
}

PageMargins &AbstractSheet::pageMargins()
{
    Q_D(AbstractSheet);
    return d->pageMargins;
}

void AbstractSheet::setPageMargins(const PageMargins &margins)
{
    Q_D(AbstractSheet);
    d->pageMargins = margins;
}

void AbstractSheet::setDefaultPageMargins()
{
    Q_D(AbstractSheet);
    d->pageMargins = PageMargins();
}

PageSetup AbstractSheet::pageSetup() const
{
    Q_D(const AbstractSheet);
    return d->pageSetup;
}

PageSetup &AbstractSheet::pageSetup()
{
    Q_D(AbstractSheet);
    return d->pageSetup;
}

void AbstractSheet::setPageSetup(const PageSetup &pageSetup)
{
    Q_D(AbstractSheet);
    d->pageSetup = pageSetup;
}

void AbstractSheet::setPaperSize(PageSetup::PaperSize paperSize)
{
    Q_D(AbstractSheet);
    d->pageSetup.paperSize = paperSize;
}

PageSetup::PaperSize AbstractSheet::paperSize() const
{
    Q_D(const AbstractSheet);
    return d->pageSetup.paperSize.value_or(PageSetup::PaperSize::Unknown);
}

void AbstractSheet::setPaperSizeMM(double width, double height)
{
    Q_D(AbstractSheet);
    d->pageSetup.paperWidth = QString::number(width,'f')+"mm";
    d->pageSetup.paperHeight = QString::number(height,'f')+"mm";
}

void AbstractSheet::setPaperSizeInches(double width, double height)
{
    Q_D(AbstractSheet);
    d->pageSetup.paperWidth = QString::number(width,'f')+"in";
    d->pageSetup.paperHeight = QString::number(height,'f')+"in";
}

QString AbstractSheet::paperWidth() const
{
    Q_D(const AbstractSheet);
    return d->pageSetup.paperWidth;
}

void AbstractSheet::setPaperWidth(const QString &width)
{
    Q_D(AbstractSheet);
    d->pageSetup.paperWidth = width;
}

QString AbstractSheet::paperHeight() const
{
    Q_D(const AbstractSheet);
    return d->pageSetup.paperHeight;
}

void AbstractSheet::setPaperHeight(const QString &height)
{
    Q_D(AbstractSheet);
    d->pageSetup.paperHeight = height;
}

PageSetup::Orientation AbstractSheet::pageOrientation() const
{
    Q_D(const AbstractSheet);
    return d->pageSetup.orientation.value_or(PageSetup::Orientation::Default);
}

void AbstractSheet::setPageOrientation(PageSetup::Orientation orientation)
{
    Q_D(AbstractSheet);
    d->pageSetup.orientation = orientation;
}

void AbstractSheet::setBackgroundImage(const QImage &image)
{
    Q_D(AbstractSheet);

    QByteArray ba;
    QBuffer buffer(&ba);
    buffer.open(QIODevice::WriteOnly);
    image.save(&buffer, "PNG");

    d->pictureFile = QSharedPointer<MediaFile>(new MediaFile(ba, QStringLiteral("png"), QStringLiteral("image/png")));
    d->workbook->addMediaFile(d->pictureFile);
}

void AbstractSheet::setBackgroundImage(const QString &fileName)
{
    Q_D(AbstractSheet);
    QByteArray ba;
    QString suffix = QFileInfo(fileName).suffix().toLower();
    if (suffix.toLower() == QStringLiteral("png") ||
        suffix.toLower() == QStringLiteral("jpg") ||
        suffix.toLower() == QStringLiteral("jpeg")) {
        QFile file(fileName);
        if (file.open(QFile::ReadOnly))
            ba = file.readAll();
    }
    else {
        QImage image(fileName);
        QBuffer buffer(&ba);
        buffer.open(QIODevice::WriteOnly);
        image.save(&buffer, "PNG");
        suffix = QStringLiteral("png");
    }
    if (!ba.isEmpty()) {
        d->pictureFile = QSharedPointer<MediaFile>(new MediaFile(ba, suffix, QStringLiteral("image/")+suffix));
        d->workbook->addMediaFile(d->pictureFile);
    }
}

QImage AbstractSheet::backgroundImage() const
{
    Q_D(const AbstractSheet);
    QImage pic;
    pic.loadFromData(d->pictureFile->contents());
    return pic;
}

bool AbstractSheet::removeBackgroundImage()
{
    Q_D(AbstractSheet);
    if (d->pictureFile) {
        d->workbook->removeMediaFile(d->pictureFile);
        d->pictureFile.clear();
        return true;
    }
    return false;
}

bool AbstractSheet::isSelected() const
{
    Q_D(const AbstractSheet);
    if (d->sheetViews.isEmpty()) return false;
    return d->sheetViews.last().tabSelected.value_or(false);
}

void AbstractSheet::setSelected(bool selected)
{
    Q_D(AbstractSheet);
    if (d->sheetViews.isEmpty()) d->sheetViews << SheetView();
    d->sheetViews.last().tabSelected = selected;
}

int AbstractSheet::viewZoomScale() const
{
    Q_D(const AbstractSheet);
    if (d->sheetViews.isEmpty()) return 100;
    return d->sheetViews.last().zoomScale.value_or(100);
}

void AbstractSheet::setViewZoomScale(int scale)
{
    Q_D(AbstractSheet);
    if (d->sheetViews.isEmpty()) d->sheetViews << SheetView();
    d->sheetViews.last().zoomScale = scale;
}

int AbstractSheet::workbookViewId() const
{
    Q_D(const AbstractSheet);
    if (d->sheetViews.isEmpty()) return 0;
    return d->sheetViews.last().workbookViewId;
}

void AbstractSheet::setWorkbookViewId(int id)
{
    Q_D(AbstractSheet);
    if (d->sheetViews.isEmpty()) d->sheetViews << SheetView();
    d->sheetViews.last().workbookViewId = id;
}

SheetView AbstractSheet::view(int index) const
{
    Q_D(const AbstractSheet);
    return d->sheetViews.value(index);
}

SheetView &AbstractSheet::view(int index)
{
    Q_D(AbstractSheet);
    if (index < 0 || index >= d->sheetViews.size())
        throw std::out_of_range("AbstractSheet::view(): view index is out of range.");
    return d->sheetViews[index];
}

int AbstractSheet::viewsCount() const
{
    Q_D(const AbstractSheet);
    return d->sheetViews.size();
}

SheetView &AbstractSheet::addView()
{
    Q_D(AbstractSheet);
    d->sheetViews << SheetView();
    return d->sheetViews.last();
}

bool AbstractSheet::removeView(int index)
{
    Q_D(AbstractSheet);
    if (index < 0 || index >= d->sheetViews.size()) return false;
    d->sheetViews.removeAt(index);
    return true;
}

SheetProtection AbstractSheet::sheetProtection() const
{
    Q_D(const AbstractSheet);
    return d->sheetProtection.value_or(SheetProtection());
}

SheetProtection &AbstractSheet::sheetProtection()
{
    Q_D(AbstractSheet);
    if (!d->sheetProtection.has_value()) d->sheetProtection = SheetProtection();
    return d->sheetProtection.value();
}

void AbstractSheet::setSheetProtection(const SheetProtection &sheetProtection)
{
    Q_D(AbstractSheet);
    d->sheetProtection = sheetProtection;
    d->sheetProtection->protectSheet = true;
}

bool AbstractSheet::isSheetProtected() const
{
    Q_D(const AbstractSheet);
    return d->sheetProtection.has_value() && d->sheetProtection.value().isValid();
}

bool AbstractSheet::isPasswordProtectionSet() const
{
    Q_D(const AbstractSheet);
    if (d->sheetProtection.has_value()) {
        const auto &s = d->sheetProtection.value();
        return !s.algorithmName.isEmpty() and !s.hashValue.isEmpty();
    }
    return false;
}

void AbstractSheet::setPassword(const QString &algorithm, const QString &hash, const QString &salt, int spinCount)
{
    Q_D(AbstractSheet);
    if (!d->sheetProtection.has_value()) d->sheetProtection = SheetProtection();
    d->sheetProtection->algorithmName = algorithm;
    d->sheetProtection->hashValue = hash;
    d->sheetProtection->protectSheet = true;
    d->sheetProtection->saltValue = salt;
    d->sheetProtection->spinCount = spinCount;
}

void AbstractSheet::setDefaultSheetProtection()
{
    Q_D(AbstractSheet);
    d->sheetProtection = SheetProtection();
    d->sheetProtection->protectSheet = true;
}

void AbstractSheet::removeSheetProtection()
{
    Q_D(AbstractSheet);
    d->sheetProtection.reset();
}

/*!
 * \internal
 */
int AbstractSheet::id() const
{
    Q_D(const AbstractSheet);
    return d->id;
}

/*!
 * \internal
 */
Drawing *AbstractSheet::drawing() const
{
    Q_D(const AbstractSheet);
    return d->drawing.get();
}

/*!
 * Return the workbook
 */
Workbook *AbstractSheet::workbook() const
{
    Q_D(const AbstractSheet);
    return d->workbook;
}

bool HeaderFooter::isValid() const
{
    if (differentOddEven.has_value()) return true;
    if (differentFirst.has_value()) return true;
    if (scaleWithDoc.has_value()) return true;
    if (alignWithMargins.has_value()) return true;

    if (!oddHeader.isEmpty()) return true;
    if (!oddFooter.isEmpty()) return true;
    if (!evenHeader.isEmpty() && differentOddEven.value_or(false)) return true;
    if (!evenFooter.isEmpty() && differentOddEven.value_or(false)) return true;
    if (!firstHeader.isEmpty() && differentFirst.value_or(false)) return true;
    if (!firstFooter.isEmpty() && differentFirst.value_or(false)) return true;
    return false;
}

void HeaderFooter::write(QXmlStreamWriter &writer) const
{
    if (isValid()) {
        writer.writeStartElement(QStringLiteral("headerFooter"));
        writeAttribute(writer, QLatin1String("differentOddEven"), differentOddEven);
        writeAttribute(writer, QLatin1String("differentFirst"), differentFirst);
        writeAttribute(writer, QLatin1String("alignWithMargins"), alignWithMargins);
        writeAttribute(writer, QLatin1String("scaleWithDoc"), scaleWithDoc);
        if (!oddHeader.isEmpty())
            writer.writeTextElement(QStringLiteral("oddHeader"), oddHeader);
        if (!oddFooter.isEmpty())
            writer.writeTextElement(QStringLiteral("oddFooter"), oddFooter);
        if (!evenHeader.isEmpty() && differentOddEven.value_or(false))
            writer.writeTextElement(QStringLiteral("evenHeader"), evenHeader);
        if (!evenFooter.isEmpty() && differentOddEven.value_or(false))
            writer.writeTextElement(QStringLiteral("evenFooter"), evenFooter);
        if (!firstHeader.isEmpty() && differentFirst.value_or(false))
            writer.writeTextElement(QStringLiteral("firstHeader"), firstHeader);
        if (!firstFooter.isEmpty() && differentFirst.value_or(false))
            writer.writeTextElement(QStringLiteral("firstFooter"), firstFooter);
        writer.writeEndElement();
    }
}

void HeaderFooter::read(QXmlStreamReader &reader)
{
    const auto &name = reader.name();
    const auto &a = reader.attributes();
    parseAttributeBool(a, QLatin1String("differentOddEven"), differentOddEven);
    parseAttributeBool(a, QLatin1String("differentFirst"), differentFirst);
    parseAttributeBool(a, QLatin1String("scaleWithDoc"), scaleWithDoc);
    parseAttributeBool(a, QLatin1String("differentOddEven"), alignWithMargins);
    while (!reader.atEnd()) {
        auto token = reader.readNext();
        if (token == QXmlStreamReader::StartElement) {
            if (reader.name() == QLatin1String("oddHeader")) oddHeader = reader.readElementText();
            else if (reader.name() == QLatin1String("oddFooter")) oddFooter = reader.readElementText();
            else if (reader.name() == QLatin1String("evenHeader")) evenHeader = reader.readElementText();
            else if (reader.name() == QLatin1String("evenFooter")) evenFooter = reader.readElementText();
            else if (reader.name() == QLatin1String("firstHeader")) firstHeader = reader.readElementText();
            else if (reader.name() == QLatin1String("firstFooter")) firstFooter = reader.readElementText();
            else reader.skipCurrentElement();
        }
        else if (token == QXmlStreamReader::EndElement && reader.name() == name)
            break;
    }
}

void PageMargins::setMarginsInches(double left, double top, double right, double bottom, double header, double footer)
{
    mMargins.insert(Position::Left, left);
    mMargins.insert(Position::Top, top);
    mMargins.insert(Position::Right, right);
    mMargins.insert(Position::Bottom, bottom);
    mMargins.insert(Position::Header, header);
    mMargins.insert(Position::Footer, footer);
}

void PageMargins::setMarginInches(Position pos, double value)
{
    mMargins.insert(pos, value);
}
void PageMargins::setLeftMarginInches(double value)
{
    mMargins.insert(Position::Left, value);
}
void PageMargins::setRightMarginInches(double value)
{
    mMargins.insert(Position::Right, value);
}
void PageMargins::setTopMarginInches(double value)
{
    mMargins.insert(Position::Top, value);
}
void PageMargins::setBottomMarginInches(double value)
{
    mMargins.insert(Position::Bottom, value);
}
void PageMargins::setHeaderMarginInches(double value)
{
    mMargins.insert(Position::Header, value);
}
void PageMargins::setFooterMarginInches(double value)
{
    mMargins.insert(Position::Footer, value);
}

void PageMargins::setMarginsMm(double left, double top, double right, double bottom, double header, double footer)
{
    mMargins.insert(Position::Left, left / 25.4);
    mMargins.insert(Position::Top, top / 25.4);
    mMargins.insert(Position::Right, right / 25.4);
    mMargins.insert(Position::Bottom, bottom / 25.4);
    mMargins.insert(Position::Header, header / 25.4);
    mMargins.insert(Position::Footer, footer / 25.4);
}
void PageMargins::setMarginMm(Position pos, double value)
{
    mMargins.insert(pos, value / 25.4);
}
void PageMargins::setLeftMarginMm(double value)
{
    mMargins.insert(Position::Left, value / 25.4);
}
void PageMargins::setRightMarginMm(double value)
{
    mMargins.insert(Position::Right, value / 25.4);
}
void PageMargins::setTopMarginMm(double value)
{
    mMargins.insert(Position::Top, value / 25.4);
}
void PageMargins::setBottomMarginMm(double value)
{
    mMargins.insert(Position::Bottom, value / 25.4);
}
void PageMargins::setHeaderMarginMm(double value)
{
    mMargins.insert(Position::Header, value / 25.4);
}
void PageMargins::setFooterMarginMm(double value)
{
    mMargins.insert(Position::Footer, value / 25.4);
}

PageMargins PageMargins::defaultPageMargins()
{
    return {};
}

PageMargins PageMargins::pageMarginsMm(double left, double top, double right, double bottom, double header, double footer)
{
    PageMargins p;
    p.setMarginsMm(left, top, right, bottom, header, footer);
    return p;
}

PageMargins PageMargins::pageMarginsInches(double left, double top, double right, double bottom, double header, double footer)
{
    PageMargins p;
    p.setMarginsInches(left, top, right, bottom, header, footer);
    return p;
}
QMap<PageMargins::Position, double> PageMargins::marginsInches() const
{
    return mMargins;
}
QMap<PageMargins::Position, double> PageMargins::marginsMm() const
{
    auto mm = mMargins;
    for (auto &m: mm) m*=25.4;
    return mm;
}
double PageMargins::marginInches(Position pos) const
{
    return mMargins.value(pos, 0.0);
}
double PageMargins::leftMarginInches() const
{
    return mMargins.value(Position::Left, 0.0);
}
double PageMargins::rightMarginInches() const
{
    return mMargins.value(Position::Right, 0.0);
}
double PageMargins::topMarginInches() const
{
    return mMargins.value(Position::Top, 0.0);
}
double PageMargins::bottomMarginInches() const
{
    return mMargins.value(Position::Bottom, 0.0);
}
double PageMargins::headerMarginInches() const
{
    return mMargins.value(Position::Header, 0.0);
}
double PageMargins::footerMarginInches() const
{
    return mMargins.value(Position::Footer, 0.0);
}
double PageMargins::marginMm(Position pos) const
{
    return mMargins.value(pos, 0.0) * 25.4;
}
double PageMargins::leftMarginMm() const
{
    return mMargins.value(Position::Left, 0.0) * 25.4;
}
double PageMargins::rightMarginMm() const
{
    return mMargins.value(Position::Right, 0.0) * 25.4;
}
double PageMargins::topMarginMm() const
{
    return mMargins.value(Position::Top, 0.0) * 25.4;
}
double PageMargins::bottomMarginMm() const
{
    return mMargins.value(Position::Bottom, 0.0) * 25.4;
}
double PageMargins::headerMarginMm() const
{
    return mMargins.value(Position::Header, 0.0) * 25.4;
}
double PageMargins::footerMarginMm() const
{
    return mMargins.value(Position::Footer, 0.0) * 25.4;
}

bool PageMargins::hasMargin(Position pos) const
{
    return mMargins.contains(pos);
}
bool PageMargins::isValid() const
{
    return !mMargins.isEmpty();
}
void PageMargins::write(QXmlStreamWriter &writer) const
{
    if (isValid()) {
        writer.writeEmptyElement(QLatin1String("pageMargins"));
        writer.writeAttribute(QLatin1String("left"), QString::number(mMargins.value(Position::Left, 0.0), 'f'));
        writer.writeAttribute(QLatin1String("right"), QString::number(mMargins.value(Position::Right, 0.0), 'f'));
        writer.writeAttribute(QLatin1String("top"), QString::number(mMargins.value(Position::Top, 0.0), 'f'));
        writer.writeAttribute(QLatin1String("bottom"), QString::number(mMargins.value(Position::Bottom, 0.0), 'f'));
        writer.writeAttribute(QLatin1String("header"), QString::number(mMargins.value(Position::Header, 0.0), 'f'));
        writer.writeAttribute(QLatin1String("footer"), QString::number(mMargins.value(Position::Footer, 0.0), 'f'));
    }
}

void PageMargins::read(QXmlStreamReader &reader)
{
    const auto &a = reader.attributes();
    if (a.hasAttribute(QLatin1String("left"))) mMargins.insert(Position::Left, a.value(QLatin1String("left")).toDouble());
    if (a.hasAttribute(QLatin1String("right"))) mMargins.insert(Position::Right, a.value(QLatin1String("right")).toDouble());
    if (a.hasAttribute(QLatin1String("top"))) mMargins.insert(Position::Top, a.value(QLatin1String("top")).toDouble());
    if (a.hasAttribute(QLatin1String("bottom"))) mMargins.insert(Position::Bottom, a.value(QLatin1String("bottom")).toDouble());
    if (a.hasAttribute(QLatin1String("header"))) mMargins.insert(Position::Header, a.value(QLatin1String("header")).toDouble());
    if (a.hasAttribute(QLatin1String("footer"))) mMargins.insert(Position::Footer, a.value(QLatin1String("footer")).toDouble());
}

bool PageSetup::isValid() const
{
    if (paperSize.has_value()) return true;
    if (!paperWidth.isEmpty()) return true;
    if (!paperHeight.isEmpty()) return true;
    if (scale.has_value()) return true;
    if (firstPageNumber.has_value()) return true;
    if (fitToWidth.has_value()) return true;
    if (fitToHeight.has_value()) return true;
    if (pageOrder.has_value()) return true;
    if (orientation.has_value()) return true;
    if (usePrinterDefaults.has_value()) return true;
    if (blackAndWhite.has_value()) return true;
    if (draft.has_value()) return true;
    if (useFirstPageNumber.has_value()) return true;
    if (horizontalDpi.has_value()) return true;
    if (verticalDpi.has_value()) return true;
    if (copies.has_value()) return true;
    if (!printerId.isEmpty()) return true;
    if (errors.has_value()) return true;
    if (cellComments.has_value()) return true;

    return false;
}

void PageSetup::write(QXmlStreamWriter &writer, bool chartsheet) const
{
    if (!isValid()) return;
    writer.writeEmptyElement(QLatin1String("pageSetup"));
    if (paperSize.has_value()) {
        int val = static_cast<int>(paperSize.value());
        writer.writeAttribute(QLatin1String("paperSize"), QString::number(val));
    }
    writeAttribute(writer, QLatin1String("paperHeight"), paperHeight);
    writeAttribute(writer, QLatin1String("paperWidth"), paperWidth);
    if (!chartsheet)
        writeAttribute(writer, QLatin1String("scale"), scale);
    writeAttribute(writer, QLatin1String("firstPageNumber"), firstPageNumber);
    if (!chartsheet)
    writeAttribute(writer, QLatin1String("fitToWidth"), fitToWidth);
    if (!chartsheet)
    writeAttribute(writer, QLatin1String("fitToHeight"), fitToHeight);
    if (!chartsheet && pageOrder.has_value()) {
        switch (pageOrder.value()) {
            case PageOrder::DownThenOver: writeAttribute(writer, QLatin1String("pageOrder"), "downThenOver"); break;
            case PageOrder::OverThenDown: writeAttribute(writer, QLatin1String("pageOrder"), "overThenDown"); break;
        }
    }
    if (orientation.has_value()) {
        switch (orientation.value()) {
            case Orientation::Default: writeAttribute(writer, QLatin1String("orientation"), "default"); break;
            case Orientation::Portrait: writeAttribute(writer, QLatin1String("orientation"), "portrait"); break;
            case Orientation::Landscape: writeAttribute(writer, QLatin1String("orientation"), "landscape"); break;
        }
    }
    writeAttribute(writer, QLatin1String("usePrinterDefaults"), usePrinterDefaults);
    writeAttribute(writer, QLatin1String("blackAndWhite"), blackAndWhite);
    writeAttribute(writer, QLatin1String("draft"), draft);
    if (!chartsheet && cellComments.has_value()) {
        switch (cellComments.value()) {
            case CellComments::AtEnd: writeAttribute(writer, QLatin1String("cellComments"), "atEnd"); break;
            case CellComments::AsDisplayed: writeAttribute(writer, QLatin1String("cellComments"), "asDisplayed"); break;
            case CellComments::DoNotPrint: writeAttribute(writer, QLatin1String("cellComments"), "none"); break;
        }
    }
    writeAttribute(writer, QLatin1String("useFirstPageNumber"), useFirstPageNumber);
    if (!chartsheet && errors.has_value()) {
        switch (errors.value()) {
            case PrintError::Dash: writeAttribute(writer, QLatin1String("errors"), "dash"); break;
            case PrintError::Displayed: writeAttribute(writer, QLatin1String("errors"), "displayed"); break;
            case PrintError::Blank: writeAttribute(writer, QLatin1String("errors"), "blank"); break;
            case PrintError::NotAvailable: writeAttribute(writer, QLatin1String("errors"), "NA"); break;
        }
    }
    writeAttribute(writer, QLatin1String("horizontalDpi"), horizontalDpi);
    writeAttribute(writer, QLatin1String("verticalDpi"), verticalDpi);
    writeAttribute(writer, QLatin1String("copies"), copies);
    writeAttribute(writer, QLatin1String("r:id"), printerId);
}

void PageSetup::read(QXmlStreamReader &reader)
{
    const auto &a = reader.attributes();

    readPaperSize(reader);
    parseAttributeString(a, QLatin1String("paperHeight"), paperHeight);
    parseAttributeString(a, QLatin1String("paperWidth"), paperWidth);
    parseAttributeInt(a, QLatin1String("scale"), scale);
    parseAttributeInt(a, QLatin1String("firstPageNumber"), firstPageNumber);
    parseAttributeInt(a, QLatin1String("fitToWidth"), fitToWidth);
    parseAttributeInt(a, QLatin1String("fitToHeight"), fitToHeight);
    if (a.hasAttribute(QLatin1String("pageOrder"))) {
        auto s = a.value(QLatin1String("pageOrder"));
        if (s == QLatin1String("downThenOver")) pageOrder = PageOrder::DownThenOver;
        if (s == QLatin1String("overThenDown")) pageOrder = PageOrder::OverThenDown;
    }
    if (a.hasAttribute(QLatin1String("orientation"))) {
        auto s = a.value(QLatin1String("orientation"));
        if (s == QLatin1String("default")) orientation = Orientation::Default;
        if (s == QLatin1String("portrait")) orientation = Orientation::Portrait;
        if (s == QLatin1String("landscape")) orientation = Orientation::Landscape;
    }
    parseAttributeBool(a, QLatin1String("usePrinterDefaults"), usePrinterDefaults);
    parseAttributeBool(a, QLatin1String("blackAndWhite"), blackAndWhite);
    parseAttributeBool(a, QLatin1String("draft"), draft);
    if (a.hasAttribute(QLatin1String("cellComments"))) {
        auto s = a.value(QLatin1String("cellComments"));
        if (s == QLatin1String("atEnd")) cellComments = CellComments::AtEnd;
        if (s == QLatin1String("asDisplayed")) cellComments = CellComments::AsDisplayed;
        if (s == QLatin1String("none")) cellComments = CellComments::DoNotPrint;
    }
    parseAttributeBool(a, QLatin1String("useFirstPageNumber"), useFirstPageNumber);
    if (a.hasAttribute(QLatin1String("errors"))) {
        auto s = a.value(QLatin1String("errors"));
        if (s == QLatin1String("dash")) errors = PrintError::Dash;
        if (s == QLatin1String("displayed")) errors = PrintError::Displayed;
        if (s == QLatin1String("blank")) errors = PrintError::Blank;
        if (s == QLatin1String("NA")) errors = PrintError::NotAvailable;
    }
    parseAttributeInt(a, QLatin1String("horizontalDpi"), horizontalDpi);
    parseAttributeInt(a, QLatin1String("verticalDpi"), verticalDpi);
    parseAttributeInt(a, QLatin1String("copies"), copies);
    parseAttributeString(a, QLatin1String("r:id"), printerId);
}

void PageSetup::readPaperSize(QXmlStreamReader &reader)
{
    if (reader.attributes().hasAttribute(QLatin1String("paperSize"))) {
        auto val = reader.attributes().value(QLatin1String("paperSize")).toInt();
        if ((val >=1 && val <= 47) || (val >=50 && val <= 118))
            paperSize = static_cast<PaperSize>(val);
        else
            paperSize = PaperSize::Unknown;
    }
}

bool SheetProtection::isValid() const
{
    if (!algorithmName.isEmpty()) return true;
    if (!hashValue.isEmpty()) return true;
    if (!saltValue.isEmpty()) return true;
    if (spinCount.has_value()) return true;
    if (protectContent.has_value()) return true;
    if (protectObjects.has_value()) return true;
    if (protectSheet.has_value()) return true;
    if (protectScenarios.has_value()) return true;
    if (protectFormatCells.has_value()) return true;
    if (protectFormatColumns.has_value()) return true;
    if (protectFormatRows.has_value()) return true;
    if (protectInsertColumns.has_value()) return true;
    if (protectInsertRows.has_value()) return true;
    if (protectInsertHyperlinks.has_value()) return true;
    if (protectDeleteColumns.has_value()) return true;
    if (protectDeleteRows.has_value()) return true;
    if (protectSelectLockedCells.has_value()) return true;
    if (protectSort.has_value()) return true;
    if (protectAutoFilter.has_value()) return true;
    if (protectPivotTables.has_value()) return true;
    if (protectSelectUnlockedCells.has_value()) return true;
    return false;
}

void SheetProtection::write(QXmlStreamWriter &writer, bool chartsheet) const
{
    writer.writeEmptyElement(QLatin1String("sheetProtection"));
    writeAttribute(writer, QLatin1String("algorithmName"), algorithmName);
    writeAttribute(writer, QLatin1String("hashValue"), hashValue);
    writeAttribute(writer, QLatin1String("saltValue"), saltValue);
    writeAttribute(writer, QLatin1String("spinCount"), spinCount);
    if (chartsheet) writeAttribute(writer, QLatin1String("content"), protectContent);
    if (!chartsheet) writeAttribute(writer, QLatin1String("sheet"), protectSheet);
    writeAttribute(writer, QLatin1String("objects"), protectObjects);
    if (!chartsheet) writeAttribute(writer, QLatin1String("scenarios"), protectScenarios);
    if (!chartsheet) writeAttribute(writer, QLatin1String("formatCells"), protectFormatCells);
    if (!chartsheet) writeAttribute(writer, QLatin1String("formatColumns"), protectFormatColumns);
    if (!chartsheet) writeAttribute(writer, QLatin1String("formatRows"), protectFormatRows);
    if (!chartsheet) writeAttribute(writer, QLatin1String("insertColumns"), protectInsertColumns);
    if (!chartsheet) writeAttribute(writer, QLatin1String("insertRows"), protectInsertRows);
    if (!chartsheet) writeAttribute(writer, QLatin1String("insertHyperlinks"), protectInsertHyperlinks);
    if (!chartsheet) writeAttribute(writer, QLatin1String("deleteColumns"), protectDeleteColumns);
    if (!chartsheet) writeAttribute(writer, QLatin1String("deleteRows"), protectDeleteRows);
    if (!chartsheet) writeAttribute(writer, QLatin1String("selectLockedCells"), protectSelectLockedCells);
    if (!chartsheet) writeAttribute(writer, QLatin1String("sort"), protectSort);
    if (!chartsheet) writeAttribute(writer, QLatin1String("autoFilter"), protectAutoFilter);
    if (!chartsheet) writeAttribute(writer, QLatin1String("pivotTables"), protectPivotTables);
    if (!chartsheet) writeAttribute(writer, QLatin1String("selectUnlockedCells"), protectSelectUnlockedCells);
}

void SheetProtection::read(QXmlStreamReader &reader)
{
    const auto &a = reader.attributes();
    parseAttributeString(a, QLatin1String("algorithmName"), algorithmName);
    parseAttributeString(a, QLatin1String("hashValue"), hashValue);
    parseAttributeString(a, QLatin1String("saltValue"), saltValue);
    parseAttributeInt(a, QLatin1String("spinCount"), spinCount);
    parseAttributeBool(a, QLatin1String("content"), protectContent);
    parseAttributeBool(a, QLatin1String("sheet"), protectSheet);
    parseAttributeBool(a, QLatin1String("objects"), protectObjects);
    parseAttributeBool(a, QLatin1String("scenarios"), protectScenarios);
    parseAttributeBool(a, QLatin1String("formatCells"), protectFormatCells);
    parseAttributeBool(a, QLatin1String("formatColumns"), protectFormatColumns);
    parseAttributeBool(a, QLatin1String("formatRows"), protectFormatRows);
    parseAttributeBool(a, QLatin1String("insertColumns"), protectInsertColumns);
    parseAttributeBool(a, QLatin1String("insertRows"), protectInsertRows);
    parseAttributeBool(a, QLatin1String("insertHyperlinks"), protectInsertHyperlinks);
    parseAttributeBool(a, QLatin1String("deleteColumns"), protectDeleteColumns);
    parseAttributeBool(a, QLatin1String("deleteRows"), protectDeleteRows);
    parseAttributeBool(a, QLatin1String("selectLockedCells"), protectSelectLockedCells);
    parseAttributeBool(a, QLatin1String("sort"), protectSort);
    parseAttributeBool(a, QLatin1String("autoFilter"), protectAutoFilter);
    parseAttributeBool(a, QLatin1String("pivotTables"), protectPivotTables);
    parseAttributeBool(a, QLatin1String("selectUnlockedCells"), protectSelectUnlockedCells);
}

}
