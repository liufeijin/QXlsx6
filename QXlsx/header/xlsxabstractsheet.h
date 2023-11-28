// xlsxabstractsheet.h

#ifndef XLSXABSTRACTSHEET_H
#define XLSXABSTRACTSHEET_H

#include "xlsxglobal.h"
#include "xlsxabstractooxmlfile.h"
#include "xlsxmain.h"
#include "xlsxsheetview.h"
#include "xlsxsheetprotection.h"
#include "xlsxpagemargins.h"
#include "xlsxpagesetup.h"

#include <QXmlStreamWriter>
#include <QXmlStreamReader>

namespace QXlsx {

/**
 * @brief The HeaderFooter class represents the header and footer of a worksheet,
 * chartsheet or their custom views.
 *
 * Headers and footers may include formatting codes that specify the position
 * and value of their parts.
 *
 * Formatting code | Meaning | Example
 * ----|----|----
 * && | ampersand | Smith && Sons
 * &font-size | font size | &18A&36B
 * &"font-name,font-type" | A text font-name string and a text font-type string | &"Arial,Bold"
 * &A | current worksheet’s tab name |  |
 * &B | toggles bold text format on/off | "ab&Bcd&Bef&Bgh"
 * &C | starts central section of a header/footer. There may be several central sections in one header/footer | &Lleft&CD&Rright&LB&CE&RH
 * &D | current date | |
 * &E | toggles double underline text format on/off | realy &Eimportant&E text
 * &F | current workbook's file name | |
 * &G | picture as a background | not implemented yet |
 * &H | toggles shadow text format | |
 * &I | toggles italic text format | |
 * &K | text font color | An RGB Color is specified as RRGGBB: &KFF0000red text, a theme color is specified as TTSNNN where TT is the theme color Id, S is either “+” or “-“ of the tint/shade value, and NNN is the tint/shade value
 * &L | starts left section of a header/footer. There may be several left sections in one header/footer | &Lleft&CD&Rright&LB&CE&RH
 * &N | total number of pages | Page &P of &N
 * &O | toggles outline text format | |
 * &P | current page number | Page &P of &N
 * &R | starts right section of a header/footer. There may be several right sections in one header/footer | &Lleft&CD&Rright&LB&CE&RH
 * &S | toggles strikethrough text format | realy &Sunimportant&S text
 * &T | current time | |
 * &U | toggles underline text format | |
 * &X | toggles superscript text format | aa&Xbb&Xcc&Ydd&Yee&Xff&Ygg
 * &Y | toggles subscript text format | aa&Xbb&Xcc&Ydd&Yee&Xff&Ygg
 * &Z | current workbook's file path | |
 */
class QXLSX_EXPORT HeaderFooter
{
public:
    /**
     * @brief Different odd and even page headers and footers indicator.
     *
     * When `true` then oddHeader/oddFooter and evenHeader/evenFooter specify
     * the page header and footer values for odd and even pages, respectively.
     * If `false` then oddHeader/oddFooter is used, even when evenHeader/evenFooter
     * are present.
     * @note If differentOddEven is `true`, but no #evenHeader and #evenFooter
     * is set, the document is ill-formed.
     * The default value is `false`.
     */
    std::optional<bool> differentOddEven;
    /**
     * @brief Different first-page header and footer indicator.
     *
     * When `true` then firstHeader and firstFooter specify the first page header
     * and footer values, respectively. If `false` and firstHeader/firstFooter are
     * present, they are ignored.
     * @note If differentFirst is `true`, but no #firstHeader and #firstFooter is set,
     * the document is ill-formed.
     * The default value is `false`.
     */
    std::optional<bool> differentFirst;
    /**
     * @brief Scale header and footer with document scaling.
     */
    std::optional<bool> scaleWithDoc;
    /**
     * @brief Align header footer margins with page margins.
     * When `true`, as left/right  margins grow and shrink, the header and footer
     * edges stay aligned with the margins. When `false`, headers and footers are
     * aligned on the paper edges, regardless of margins.
     */
    std::optional<bool> alignWithMargins;

    QString oddHeader;
    QString oddFooter;
    QString evenHeader;
    QString evenFooter;
    QString firstHeader;
    QString firstFooter;
    bool isValid() const;
    void write(QXmlStreamWriter &writer, const QString &name) const;
    void read(QXmlStreamReader &reader);
};


class Workbook;
class Drawing;
class AbstractSheetPrivate;

/**
 * @brief The AbstractSheet class is a base class for different sheet types.
 *
 * It provides parameters common to both worksheets and chartsheets.
 *
 * ## General parameters
 *
 * Each sheet can have an editable name (see #name(), #setName(), #rename())
 * and a consistent unique code name (name that the user cannot change, see
 * #codeName(), #setCodeName()).
 *
 * You can get the sheet #type(), but you cannot change it.
 *
 * The visibility of the sheet is managed via #visibility(), #setVisibility()
 * and convenience methods #isHidden(), #isVisible(), #setHidden(),
 * #setVisible().
 *
 * The tab color of the sheet can be obtained via #tabColor() and set with
 * #setTabColor().
 *
 * A sheet can have a background image set (though a chartsheet needs a bit of
 * tweaking). See #backgroundImage(), #removeBackgroundImage() and
 * #setBackgroundImage().
 *
 * ## Sheet Views
 *
 * Each sheet can have 1 to infinity 'sheet views', that display a specific
 * portion of the sheet with specific view parameters. To ease the development
 * SheetView class represents view parameters of both worksheets and
 * chartsheets, though not all SheetView attributes are applicable to
 * chartsheets. See SheetView documentation. If a non-applicable parameter
 * appears, it will be ignored on the document write.
 *
 * There's always at least one default sheet view.
 *
 * The following methods manage sheet views:
 *
 * - #view() returns the specific view.
 * - #lastView() returns the latest added view.
 * - #viewsCount() returns the count of views in the sheet.
 * - #addView() adds a view.
 * - #removeView() removes the view.
 *
 * The following methods manage the parameters of the _latest added_ view:
 *
 * 1. In worksheets only.
 *   - Worksheet::isWindowProtected(), Worksheet::setWindowProtected() manage
 *     view protection.
 *   - Worksheet::isFormulasVisible(), Worksheet::setFormulasVisible() manage
 *     formulas visibility.
 *   - Worksheet::isGridLinesVisible(), Worksheet::setGridLinesVisible() manage
 *     visibility of gridlines between sheet cells.
 *   - Worksheet::isRowColumnHeadersVisible(),
 *     Worksheet::setRowColumnHeadersVisible() manage headers visibility.
 *   - Worksheet::isZerosVisible(), Worksheet::setZerosVisible() manage the zero
 *     values appearance.
 *   - Worksheet::isRightToLeft(), Worksheet::setRightToLeft() manage the
 *     right-to-left appearance of the view.
 *   - Worksheet::isRulerVisible(), Worksheet::setRulerVisible() manage the
 *     ruler visibility.
 *   - Worksheet::isOutlineSymbolsVisible(),
 *     Worksheet::setOutlineSymbolsVisible() manage the outline symbols
 *     visibility.
 *   - Worksheet::isPageMarginsVisible(), Worksheet::setPageMarginsVisible()
 *     manage visibility of the page layout margins.
 *   - Worksheet::isDefaultGridColorUsed(), Worksheet::setDefaultGridColorUsed()
 *     manage using the default/custom grid lines color.
 *   - Worksheet::viewType(), Worksheet::setViewType() manage the type of the
 *     sheet view.
 *   - Worksheet::viewTopLeftCell(), Worksheet::setViewTopLeftCell() manage the
 *     reference to the view's top left cell.
 *   - Worksheet::viewColorIndex(), Worksheet::setViewColorIndex() manage the
 *     sheet tab color in the current view.
 *   - Worksheet::selection(), Worksheet::setSelection(),
 *     Worksheet::selectedRanges(), Worksheet::addSelection(),
 *     Worksheet::removeSelection(), and Worksheet::clearSelection() manage the
 *     cells selection on the worksheet.
 *   - Worksheet::activeCell(), Worksheet::setActiveCell() manage the active
 *     cell on the worksheet.
 * 2. In chartsheets only:
 *   - Chartsheet::zoomToFit(), Chartsheet::setZoomToFit() manage the
 *     autofitting of the chart in the sheet.
 * 3. In worksheets and chartsheets:
 *   - #isSelected(), #setSelected() manage the selection of the sheet tab.
 *   - #viewZoomScale() and #setViewZoomScale() manage the view zoom
 *     magnification.
 *   - #workbookViewId() and #setWorkbookViewId() allow to associate the view
 *     with a specific workbook view.
 *
 * The above-mentioned methods return `std::nullopt` if the corresponding parameters
 * were not set. See SheetView and Selection documentation on the default
 * values.
 *
 * ## Sheet printing parameters and page setup parameters.
 *
 * Both chartsheets and worksheets may have page setup parameters specified. See
 * PageSetup class for the description of all parameters.
 *
 * You can access all page setup parameters via #pageSetup() or use the
 * following convenience methods:
 *
 * 1. In both chartsheets and worksheets:
 *   - #paperSize(), #setPaperSize(), #setPaperSizeMM(), #setPaperSizeInches(),
 *     #paperWidth(), #setPaperWidth(), #paperHeight(), #setPaperHeight() manage
 *     paper size.
 *   - #pageOrientation(), #setPageOrientation() manage page orientation.
 *   - #firstPageNumber(), #setFirstPageNumber(), #useFirstPageNumber(),
 *     #setUseFirstPageNumber() manage the page number for first printed page.
 *   - #printerDefaultsUsed(), #setPrinterDefaultsUsed() manage the printer
 *     default parameters.
 *   - #printBlackAndWhite(), #setPrintBlackAndWhite(), #printDraft(),
 *     #setPrintDraft(), #horizontalDpi(), #setHorizontalDpi(), #verticalDpi(),
 *     #setVerticalDpi() manage quality of printing.
 *   - #copies(), #setCopies() manage how many copies to print.
 * 2. In worksheets only:
 *   - Worksheet::printScale(), Worksheet::setPrintScale(),
 *     Worksheet::fitToWidth(), Worksheet::setFitToWidth(),
 *     Worksheet::fitToHeight(), Worksheet::setFitToHeight() manage the scale of
 *     the printed worksheet.
 *   - Worksheet::pageOrder(), Worksheet::setPageOrder() manage the order in
 *     which worksheet pages are printed.
 *   - Worksheet::printErrors(), Worksheet::setPrintErrors(),
 *     Worksheet::printCellComments(), Worksheet::setPrintCellComments() manage
 *     how to print additional cell info.
 *   - Worksheet::printGridLines(), Worksheet::setPrintGridLines() manage grid
 *     lines printing.
 *   - Worksheet::printHeadings(), Worksheet::setPrintHeadings() manage printing
 *     of row and column headings.
 *   - Worksheet::printHorizontalCentered(),
 *     Worksheet::setPrintHorizontalCentered(),
 *     Worksheet::printVerticalCentered(),
 *     Worksheet::setPrintVerticalCentered() manage the arrangement of the sheet
 *     data on paper.
 *
 * All these methods return `std::nullopt` if the parameters were not set.
 * See PageSetup class documentation on the default values.
 *
 * ## Headers and footers
 *
 * Each sheet can have a header or footer defined. Headers and footers can be
 * different for odd and even pages or for the first page.
 *
 * Headers and footers are managed through the HeaderFooter class attributes.
 * See #headerFooter(). But a number of convenience methods exists that simplify
 * the header/footer manipulation.
 *
 * ## Page margins
 *
 * You can specify page margins either in millimeters or in inches with
 * #setPageMarginsInches(), #setPageMarginsMm() or directly via #pageMargins()
 * method that returns the PageMargins object. You can also clear all values
 * with #setDefaultPageMargins().
 *
 * ## Sheet protection
 *
 * The sheet protection parameters are accessible with #sheetProtection() and
 * #setSheetProtection() methods. See the SheetProtection class documentation.
 *
 * There is a number of convenience methods that allow to quickly test and set
 * the password protection or remove the protection completely:
 * #setDefaultSheetProtection(), #removeSheetProtection(),
 * #setPasswordProtection(), #isPasswordProtectionSet(), #isSheetProtected().
 *
 * See also the Worksheet documentation on how to protect specific cell ranges,
 * as these methods are accessible only in worksheets.
 *
 * See [SheetProtection](../../examples/SheetProtection/main.cpp) example.
 *
 */
class QXLSX_EXPORT AbstractSheet : public AbstractOOXmlFile
{
    Q_DECLARE_PRIVATE(AbstractSheet)
public:
    /**
     * @brief The Type enum specifies the sheet type.
     */
    enum class Type {
        Worksheet, /**< Worksheet that contains cells with data etc.*/
        Chartsheet, /**< Chartsheet that contains a chart.*/
        Dialogsheet, /**< @note This sheet type is currently not supported. */
        Macrosheet /**< @note This sheet type is currently not supported.*/
    };
    /**
     * @brief The Visibility enum specifies the sheet visibility.
     */
    enum class Visibility {
        Visible, /**< The sheet is visible. */
        Hidden, /**< The sheet is hidden but can be shown via the UI.*/
        VeryHidden /**< The sheet is hidden and cannot be shown via the UI.
To set the sheet VeryHidden use #setVisibility() method.*/
    };

    /**
     * @brief returns pointer to this sheet's workbook.
     * @return
     */
    Workbook *workbook() const;
    /**
     * @brief returns the sheet's name.
     * @return
     */
    QString name() const;

    /**
     * @brief tries to rename the sheet
     * @param newName new sheet name.
     * @return `true` if renaming was successful.
     */
    bool setName(const QString &newName);
    /**
     * @brief tries to rename the sheet.
     *
     * This is equivalent to `rename(newName)`.
     *
     * @param sheetName new sheet name.
     * @return `true` if renaming was successful.
     *
     */
    bool rename(const QString &sheetName);

    /**
     * @brief returns the sheet's type.
     * @return
     */
    Type type() const;
    /**
     * @brief returns the sheet's visibility.
     * @return either Visibility::Visible, Visibility::Hidden or Visibility::VeryHidden.
     * Sheets are Visible by default.
     * @sa #isVisible(), #isHidden()
     */
    Visibility visibility() const;
    /**
     * @brief sets the sheet's visibility.
     * @param visibility
     * @note Use this method to set the sheet Visibility::VeryHidden.
     * @sa #setHidden(), #setVisible().
     */
    void setVisibility(Visibility visibility);
    /**
     * @brief returns whether the sheet is hidden or very hidden.
     * @return `true` if the sheet's #visibility() is Visibility::Hidden or Visibility::VeryHidden,
     * `false` otherwise.
     * This is a convenience method, equivalent to @code visibility() == Visibility::Hidden || visibility() == Visibility::VeryHidden @endcode
     * By default sheets are visible.
     * @sa #visibility(), #isVisible().
     */
    bool isHidden() const;
    /**
     * @brief returns whether the sheet is visible.
     * @return `true` if the sheet's #visibility() is Visibility::Visible, `false` otherwise.
     * This is a convenience method, equivalent to @code visibility() == Visibility::Visible @endcode
     * By default sheets are visible.
     * @sa #isHidden(), #visibility().
     */
    bool isVisible() const;
    /**
     * @brief sets the sheet's visibility.
     * @param hidden if `true`, the sheet's visibility will be set to Visibility::Hidden,
     * if `false`, the sheet will be Visibility::Visible.
     * @note If @a hidden is `true`, the very hidden sheet stays very hidden.
     * You can check the exact visibility with #visibility().
     * @sa #setVisibility(), #setVisible().
     */
    void setHidden(bool hidden);
    /**
     * @brief sets the sheet's visibility.
     * @param visible if `true`, the sheet's visibility will be set to Visibility::Visible,
     * if `false`, the sheet will be Visibility::Hidden.
     * @note If @a visible is `false`, the very hidden sheet stays very hidden.
     * You can check the exact visibility with #visibility().
     * @sa #setVisibility(), #setHidden().
     */
    void setVisible(bool visible);

    /**
     * @brief returns the header and footer parameters of a sheet.
     * @return a copy of the HeaderFooter object.
     */
    HeaderFooter headerFooter() const;
    /**
     * @brief returns the header and footer parameters of a sheet.
     * @return a reference to the HeaderFooter object.
     */
    HeaderFooter &headerFooter();

    /**
     * @brief returns whether the sheet has different header/footer for odd and
     * even pages.
     * @return When `true` then #oddHeader() / #oddFooter() and #evenHeader() /
     * #evenFooter() specify the page header and footer values for odd and even
     * pages, respectively. If `false` then #oddHeader() / #oddFooter() or
     * #header() / #footer() are used, even when #evenHeader() / #evenFooter()
     * are not empty.
     *
     * The default value is `false`.
     *
     * @sa #header(), #footer(), #oddHeader(), #oddFooter(), #evenHeader(),
     * #evenFooter(), #setHeader(), #setFooter(), #setOddHeader(),
     * #setOddFooter(), #setEvenHeader(), #setEvenFooter(), #setHeaders(),
     * #setFooters().
     */
    std::optional<bool> differentOddEvenPage() const;
    /**
     * @brief sets whether a sheet has different header/footer for odd and even pages.
     * @param different When `true` then #oddHeader() / #oddFooter() and #evenHeader() /
     * #evenFooter() specify the page header and footer values for odd and even
     * pages, respectively. If `false` then #oddHeader() / #oddFooter() or #header() / #footer()
     * are used, even when #evenHeader() / #evenFooter() are not empty.
     *
     * If not set, the default value is `false`.
     *
     * @sa #header(), #footer(), #oddHeader(), #oddFooter(), #evenHeader(), #evenFooter(),
     * #setHeader(), #setFooter(), #setOddHeader(), #setOddFooter(), #setEvenHeader(),
     * #setEvenFooter(), #setHeaders(), #setFooters().
     */
    void setDifferentOddEvenPage(bool different);
    /**
     * @brief returns whether a sheet has different header/footer for the first page.
     * @return When `true`, then #firstHeader() / #firstFooter() specify the first page
     * header and footer. If `false`, then #firstHeader() and #firstFooter() are ignored
     * even if they are not empty.
     */
    std::optional<bool> differentFirstPage() const;
    /**
     * @brief sets whether a sheet has different header/footer for the first page.
     * @param different When `true`, then #firstHeader() / #firstFooter() specify the first page
     * header and footer.
     * @sa #setHeaders(), #setFooters().
     */
    void setDifferentFirstPage(bool different);
    /**
     * @brief returns the page header.
     * @return Page header as a string that may include formatting codes. See
     * HeaderFooter class for a list of formatting codes. The return value is
     * empty if #differentOddEvenPage() and/or #differentFirstPage() is `true`.
     * In this case use #oddHeader(), #evenHeader(), #firstHeader().
     */
    QString header() const;
    /**
     * @brief sets the page header.
     * @param header a string that may include formatting codes. See
     * HeaderFooter class for a list of formatting codes.
     *
     * This method effectively sets odd/even pages and first page to `false`. If
     * needed, use #setOddHeader(), #setEvenHeader(), #setFirstHeader().
     */
    void setHeader(const QString &header);
    /**
     * @brief returns the header of odd pages.
     * @return a string that may include formatting codes. See HeaderFooter
     * class for a list of formatting codes.
     *
     * The return value is empty if #differentOddEvenPage() is `false`.
     */
    QString oddHeader() const;
    /**
     * @brief sets the header of odd pages.
     * @param header a string that may include formatting codes. See HeaderFooter class
     * for a list of formatting codes.
     *
     * This method effectively sets odd/even pages to `true`. Note that if you don't specify
     * even header and footer, the document is ill-formed.
     * @sa #setHeaders().
     */
    void setOddHeader(const QString &header);
    /**
     * @brief returns the header of even pages.
     * @return a string that may include formatting codes. See HeaderFooter class
     * for a list of formatting codes.
     *
     * The return value is empty if #differentOddEvenPage()
     * is `false`.
     */
    QString evenHeader() const;
    /**
     * @brief sets the header of even pages.
     * @param header a string that may include formatting codes. See HeaderFooter class
     * for a list of formatting codes.
     *
     * This method effectively sets odd/even pages to `true`. Note that if you don't specify
     * odd header and footer, the document is ill-formed.
     * @sa #setHeaders().
     */
    void setEvenHeader(const QString &header);
    /**
     * @brief returns the header of the first page.
     * @return a string that may include formatting codes. See HeaderFooter class
     * for a list of formatting codes.
     *
     * The return value is empty if #differentFirstPage()
     * is `false`.
     */
    QString firstHeader() const;
    /**
     * @brief sets the header of the first page.
     * @param header a string that may include formatting codes. See HeaderFooter class
     * for a list of formatting codes.
     *
     * This method effectively sets first page to `true`.
     * @sa #setHeaders().
     */
    void setFirstHeader(const QString &header);
    /**
     * @brief sets headers for page
     * @param oddHeader non-empty string that is used for headers if @a evenHeader is empty,
     * and for odd page header if @a evenHeader is not empty.
     * @param evenHeader a string that is used for even page header. If not empty, then
     * #setDifferentOddEvenPage() is set to `true`.
     * @param firstHeader a string that is used for the first page. If not empty,
     * then #setDifferentFirstPage() is set to `true`.
     */
    void setHeaders(const QString &oddHeader, const QString &evenHeader = QString(), const QString &firstHeader = QString());

    // footers
    /**
     * @brief returns the page footer.
     * @return Page footer as a string that may include formatting codes. See HeaderFooter class
     * for a list of formatting codes. The return value is empty if #differentOddEvenPage()
     * and/or #differentFirstPage() is `true`. In this case use #oddFooter(), #evenFooter(), #firstFooter().
     */
    QString footer() const;
    /**
     * @brief sets the page footer.
     * @param footer a string that may include formatting codes. See HeaderFooter class
     * for a list of formatting codes.
     *
     * This method effectively sets odd/even pages and first page to `false`. If needed, use
     * #setOddFooter(), #setEvenFooter(), #setFirstFooter().
     */
    void setFooter(const QString &footer);
    /**
     * @brief returns the footer of odd pages.
     * @return a string that may include formatting codes. See HeaderFooter class
     * for a list of formatting codes.
     *
     * The return value is empty if #differentOddEvenPage()
     * is `false`.
     */
    QString oddFooter() const;
    /**
     * @brief sets the footer of odd pages.
     * @param footer a string that may include formatting codes. See HeaderFooter class
     * for a list of formatting codes.
     *
     * This method effectively sets odd/even pages to `true`. Note that if you don't specify
     * even header and footer, the document is ill-formed.
     * @sa #setFooters().
     */
    void setOddFooter(const QString &footer);
    /**
     * @brief returns the footer of even pages.
     * @return a string that may include formatting codes. See HeaderFooter class
     * for a list of formatting codes.
     *
     * The return value is empty if #differentOddEvenPage()
     * is `false`.
     */
    QString evenFooter() const;
    /**
     * @brief sets the footer of even pages.
     * @param footer a string that may include formatting codes. See HeaderFooter class
     * for a list of formatting codes.
     *
     * This method effectively sets odd/even pages to `true`. Note that if you don't specify
     * odd header and footer, the document is ill-formed.
     * @sa #setFooters().
     */
    void setEvenFooter(const QString &footer);
    /**
     * @brief returns the footer of the first page.
     * @return a string that may include formatting codes. See HeaderFooter class
     * for a list of formatting codes.
     *
     * The return value is empty if #differentFirstPage()
     * is `false`.
     */
    QString firstFooter() const;
    /**
     * @brief sets the footer of the first page.
     * @param footer a string that may include formatting codes. See HeaderFooter class
     * for a list of formatting codes.
     *
     * This method effectively sets first page to `true`.
     * @sa #setFooters().
     */
    void setFirstFooter(const QString &footer);
    /**
     * @brief sets footers for page
     * @param oddFooter non-empty string that is used for footers if \a evenFooter is empty,
     * and for odd page footer if \a evenFooter is not empty.
     * @param evenFooter a string that is used for even page footer. If not empty, then
     * #setDifferentOddEvenPage() is set to `true`.
     * @param firstFooter a string that is used for the first page. If not empty,
     * then #setDifferentFirstPage() is set to `true`.
     */
    void setFooters(const QString &oddFooter, const QString &evenFooter = QString(), const QString &firstFooter = QString());

    // Page margins methods
    /**
     * @brief returns the sheet page margins in inches.
     * @return A map of page margins that can contain 0 to 6 values.
     */
    QMap<PageMargins::Position, double> pageMarginsInches() const;
    /**
     * @brief returns the sheet page margins in millimeters.
     * @return A map of page margins that can contain 0 to 6 values.
     */
    QMap<PageMargins::Position, double> pageMarginsMm() const;
    /**
     * @brief sets the sheet page margins in inches.
     * @param left left margin.
     * @param top top margin.
     * @param right right margin.
     * @param bottom bottom margin.
     * @param header margin afrea page header.
     * @param footer margin before page footer.
     *
     * @note If a value is 0, this method sets a zero margin. To clear margins
     * (that is setting their default values) use #setDefaultPageMargins()
     * method.
     */
    void setPageMarginsInches(double left = 0, double top = 0, double right = 0,
                              double bottom = 0, double header = 0, double footer = 0);
    /**
     * @brief sets the sheet page margins in millimetres.
     * @param left left margin.
     * @param top top margin.
     * @param right right margin.
     * @param bottom bottom margin.
     * @param header margin afrea page header.
     * @param footer margin before page footer.
     *
     * @note If a value is 0, this method sets a zero margin. To clear margins
     * (that is setting their default values) use #setDefaultPageMargins()
     * method.
     */
    void setPageMarginsMm(double left = 0, double top = 0, double right = 0,
                          double bottom = 0, double header = 0, double footer = 0);
    /**
     * @brief returns the sheet page margins.
     * @return A copy of the PageMargins object.
     */
    PageMargins pageMargins() const;
    /**
     * @brief returns the sheet page margins.
     * @return A reference to the PageMargins object.
     */
    PageMargins &pageMargins();
    /**
     * @brief sets the sheet page margins.
     * @param margins A PageMargins object.
     */
    void setPageMargins(const PageMargins &margins);
    /**
     * @brief sets default page margins, clearing all margins values.
     */
    void setDefaultPageMargins();

    /// Page setup methods

    /**
     * @brief returns the page setup and printing parameters of the sheet.
     * @return a copy of PageSetup object.
     */
    PageSetup pageSetup() const;
    /**
     * @brief returns the page setup and printing parameters of the sheet.
     * @return reference to the PageSetup object.
     */
    PageSetup &pageSetup();
    /**
     * @brief sets the sheet's page setup and printing parameters.
     * @param pageSetup page parameters as a PageSetup object.
     * @note To clear all page parameters and set them to their default values use ```sheet->setPageSetup(PageSetup());```
     */
    void setPageSetup(const PageSetup &pageSetup);
    /**
     * @brief sets the sheet's paper size as one of the predefined sizes.
     * @sa PageSetup::paperSize.
     */
    void setPaperSize(PageSetup::PaperSize paperSize);
    /**
     * @brief returns the sheet's paper size as one of the predefined sizes.
     * @return If custom paper size is set with #setPaperSizeMM(), #setPaperSizeInches(),
     * #setPaperWidth(), #setPaperHeight() or PageSetup::paperWidth, PageSetup::paperHeight,
     * then this method returns PageSetup::PaperSize::Unknown.
     * @sa PageSetup::paperSize.
     */
    std::optional<PageSetup::PaperSize> paperSize() const;
    /**
     * @brief sets the sheet's paper size in millimeters.
     * @param width width in millimeters.
     * @param height height in millimeters.
     */
    void setPaperSizeMM(double width, double height);
    /**
     * @brief sets the sheet's paper size in inches.
     * @param width width in inches.
     * @param height height in inches.
     */
    void setPaperSizeInches(double width, double height);
    /**
     * @brief returns the sheet's paper width as a univeral measure (f.e. "210mm").
     * @return A string in format "-?[0-9]+(\.[0-9]+)?(mm|cm|in|pt|pc|pi)" if paper width was set,
     * empty string otherwise.
     * @note This method does not convert #paperSize() value to string. If #paperSize() was set,
     * and paperWidth was not set, this method returns an empty string.
     */
    QString paperWidth() const;
    /**
     * @brief sets the sheet's paper width as a univeral measure (f.e. "210mm").
     * @param width A string in format "-?[0-9]+(\.[0-9]+)?(mm|cm|in|pt|pc|pi)".
     * @note This method does not check @a width for validity.
     */
    void setPaperWidth(const QString &width);
    /**
     * @brief returns the sheet's paper height as a univeral measure (f.e. "210mm").
     * @return A string in format "-?[0-9]+(\.[0-9]+)?(mm|cm|in|pt|pc|pi)" if paper height was set,
     * empty string otherwise.
     */
    QString paperHeight() const;
    /**
     * @brief sets the sheet's paper height as a univeral measure (f.e. "210mm").
     * @param height A string in format "-?[0-9]+(\.[0-9]+)?(mm|cm|in|pt|pc|pi)".
     * @note This method does not check @a height for validity.
     */
    void setPaperHeight(const QString &height);
    /**
     * @brief returns the sheet's paper orientation.
     * @return PageSetup::Orientation value. If no orientation was set, returns
     * the default value of PageSetup::Orientation::Default.
     */
    std::optional<PageSetup::Orientation> pageOrientation() const;
    /**
     * @brief sets the sheet's paper orientation.
     * @param orientation PageSetup::Orientation value.
     *
     * If not set, PageSetup::Orientation::Default is assumed.
     */
    void setPageOrientation(PageSetup::Orientation orientation);
    /**
     * @brief returns the page number for first printed page.
     * @return integer value starting from 1.
     *
     * The default value is 1.
     */
    std::optional<int> firstPageNumber() const;
    /**
     * @brief sets the page number for first printed page.
     * @param number integer value starting from 1.
     * @note This method also invokes `setUseFirstPageNumber(true);`
     *
     * If not set, 1 is assumed.
     */
    void setFirstPageNumber(int number);
    /**
     * @brief returns whether to use #firstPageNumber value for first page number,
     * and do not auto number the pages.
     * @return If `true`, then #firstPageNumber() value is used.
     *
     * The default value is `false`.
     */
    std::optional<bool> useFirstPageNumber() const;
    /**
     * @brief sets whether to use #firstPageNumber value for first page number,
     * and do not auto number the pages.
     * @param use If `true`, then #firstPageNumber() value is used.
     *
     * If not set, then `false` is assumed.
     */
    void setUseFirstPageNumber(bool use);
    /**
     * @brief returns whether to use the printer default parameters.
     *
     * Use the printer’s defaults settings for page setup values and don't use the
     * default values of the file. Example: if dpi is not present or specified in
     * the file, the application must not assume 600dpi as specified as a default
     * and instead must let the printer specify the default dpi.
     *
     * If no value is specified, then `true` is assumed.
     *
     * @return If `true`, then the sheet's default parameters are ignored and printer
     * parameters are used.
     */
    std::optional<bool> printerDefaultsUsed() const;
    /**
     * @brief sets whether to use the printer default parameters.
     *
     * Use the printer’s defaults settings for page setup values and don't use the
     * default values of the file. Example: if dpi is not present or specified in
     * the file, the application must not assume 600dpi as specified as a default
     * and instead must let the printer specify the default dpi.
     *
     * If no value is specified, then `true` is assumed.
     * @param value If `true`, then the sheet's default parameters are ignored and printer
     * parameters are used.
     */
    void setPrinterDefaultsUsed(bool value);
    /**
     * @brief returns whether to print in black and white.
     * @return If `true`, then the sheet is printed in black and white.
     *
     * The default value is `false`.
     */
    std::optional<bool> printBlackAndWhite() const;
    /**
     * @brief sets whether to print in black and white.
     * @param value If `true`, then the sheet is printed in black and white.
     *
     * If not set, then `false` is assumed.
     */
    void setPrintBlackAndWhite(bool value);
    /**
     * @brief returns whether to print the page as a draft (i.e. without graphics).
     * @return
     *
     * If no value is specified, then `false` is assumed.
     */
    std::optional<bool> printDraft() const;
    /**
     * @brief sets whether to print the page as a draft (i.e. without graphics).
     * @param draft
     *
     * If no value is specified, then `false` is assumed.
     */
    void setPrintDraft(bool draft);
    /**
     * @brief returns the horizontal print resolution of the device.
     * @return print resolution in DPI.
     *
     * The default value is 600.
     */
    std::optional<int> horizontalDpi() const;
    /**
     * @brief sets the horizontal print resolution of the device.
     * @param dpi print resolution in DPI.
     *
     * If not set, then 600 is assumed.
     */
    void setHorizontalDpi(int dpi);
    /**
     * @brief returns the vertical print resolution of the device.
     * @return int value if horizontal dpi was set, `nullopt` otherwise. If not set,
     * 600 is assumed.
     */
    std::optional<int> verticalDpi() const;
    /**
     * @brief sets the vertical print resolution of the device.
     * @param dpi print resolution in DPI.
     *
     * If not set, then 600 is assumed.
     */
    void setVerticalDpi(int dpi);
    /**
     * @brief returns how many copies to print.
     * @return int value if copies was set, `nullopt` otherwise. If not set, 1 is assumed.
     */
    std::optional<int> copies() const;
    /**
     * @brief sets how many copies to print.
     * @param count a positive integer value.
     *
     * If not set, 1 is assumed.
     */
    void setCopies(int count);



    //Sheet properties, common to both chartsheets and worksheets
    /**
     * @brief returns whether the sheet is published.
     * @return `true` if the sheet was published.
     *
     * If not set, the default value is `true`.
     */
    std::optional<bool> isPublished() const;
    /**
     * @brief sets whether the sheet is published.
     * @param published
     *
     * The default value is `true`.
     */
    void setPublished(bool published);
    /**
     * @brief returns the unique stable sheet name.
     *
     * codeName should not change over time, and does not change from user input.
     * @return Non-empty string if the codeName was set.
     */
    QString codeName() const;
    /**
     * @brief sets the unique stable sheet name.
     *
     * codeName should not change over time, and does not change from user input.
     * @param codeName a unique string.
     * @return `true` if the code name is successfully changed, `false` if this
     * code name is not unique.
     * @note Use QUuid class to create a unique codeName.
     */
    bool setCodeName(const QString &codeName);
    /**
     * @brief returns the sheet's tab color.
     * @return Valid Color if the tab color was set.
     */
    Color tabColor() const;
    /**
     * @brief sets the sheet's tab color.
     * @param color the tab color.
     */
    void setTabColor(const Color &color);
    /**
     * @overload
     * @brief sets the sheet's tab color.
     * @param color the tab color.
     */
    void setTabColor(const QColor &color);



    /**
     * @brief sets the sheet's background image
     * @param image
     */
    void setBackgroundImage(const QImage &image);
    /**
     * @brief sets the sheet's background image
     * @param fileName the name of the file where to load image from.
     * @note If the image is a PNG or JPG image, it will be written as a PNG or JPG respectively.
     * All other pictures will be converted to PNG.
     */
    void setBackgroundImage(const QString &fileName);
    /**
     * @brief returns the sheet's background image
     * @return non-null QImage if the background image was set, null QImage otherwise.
     */
    QImage backgroundImage() const;
    /**
     * @brief removes the sheet's background image
     * @return `true` if the background image was removed, `false` otherwise.
     */
    bool removeBackgroundImage();


    /**
     * @brief returns the sheet protection parameters
     * @return a copy of SheetProtection object.
     */
    SheetProtection sheetProtection() const;
    /**
     * @brief returns the sheet protection parameters.
     * @return a reference to the SheetProtection object.
     */
    SheetProtection &sheetProtection();
    /**
     * @brief sets the sheet protection parameters.
     * @param sheetProtection a SheetProtection object.
     * @note This method also invokes `SheetProtection::setProtectSheet(true)`
     */
    void setSheetProtection(const SheetProtection &sheetProtection);
    /**
     * @brief returns whether the sheet is protected.
     * @return `true` if any of the protection parameters were set.
     * @sa #isPasswordProtectionSet().
     */
    bool isSheetProtected() const;
    /**
     * @brief returns whether the sheet is protected with password.
     * @return `true` if both SheetProtection::algorithmName and SheetProtection::hashValue
     * parameters are set in #sheetProtection(), `false` otherwise.
     *
     * If password protection is set, then #isSheetProtected() also returns `true`.
     */
    bool isPasswordProtectionSet() const;
    /**
     * @brief sets the password protection to the sheet.
     * @param password a password.
     * @param algorithm a string that describes the hashing algorithm used.
     * See PasswordProtection::algorithmName for some reserved values. If no
     * algorithm is specified, 'SHA-512' is used.
     * @note This method supports only the following algorithms as they
     * are supported by QCryptographicHash: MD4, MD5, SHA-1, SHA-256, SHA-384,
     * SHA-512. If a non-supported algorithm is specified, the password
     * protection will not be set and this method will return `false`.
     * Excel uses SHA-512.
     * @param salt a salt string. If no salt is specified, either an empty
     * string or a random 16-byte string is used depending on the
     * PasswordProtection::randomizedSalt() value.
     * @param spinCount count of iterations to compute the password hash (more
     * is better, Excel uses value of 100,000). If no spinCount is specified,
     * the value of 100,000 is used.
     * @note This method computes the hashed password value and stores it in
     * the SheetProtection attributes. You can access this value (in the base-64
     * form) via SheetProtection::protection().hashValue. Moreover, the salt
     * value is also stored in the base-64 form in the
     * SheetProtection::protection().saltValue attribute. There's no way to know
     * the actual password from xlsx.
     */
    bool setPasswordProtection(const QString &password,
                               const QString &algorithm=QStringLiteral("SHA-512"),
                               const QString &salt = QString(),
                               int spinCount = 100000);
    /**
     * @brief sets the default sheet protection parameters.
     *
     * This method sets the default protection parameters to the sheet.
     * @see SheetProtection class description for the default values.
     * @note This method deletes the previous sheet protection parameters.
     */
    void setDefaultSheetProtection();
    /**
     * @brief deletes any sheet protection parameters that have been set,
     * including the default ones.
     */
    void removeSheetProtection();

    /**
     * @brief returns whether the sheet's tab is selected.
     * @return valid bool if the parameter was set, `nullopt` otherwise.
     * @note This method returns the parameter of _the last added view_.
     *
     * The default value is `false`.
     */
    std::optional<bool> isSelected() const;
    /**
     * @brief sets @a selected to the sheet's tab.
     * @param selected tab selection state.
     *
     * @note This method sets the parameter of _the last added view_.
     *
     * If not set, the default value is `false` (not selected).
     */
    void setSelected(bool selected);
    /**
     * @brief returns window zoom magnification for last added view as a percent
     * value.
     *
     * This parameter is restricted to values ranging from 10 to 400.
     *
     * @note This method returns the parameter of _the last added view_.
     *
     * The default value is 100.
     *
     * @return window zoom magnification if it was set, `nullopt` otherwise.
     * @note To get the view scales for specific view types see
     * SheetView::zoomScaleNormal, SheetView::zoomScalePageBreakView,
     * SheetView::zoomScalePageLayoutView.
     */
    std::optional<int> viewZoomScale() const;
    /**
     * @brief sets window zoom magnification for last added view as a percent value.
     * @param scale value ranging from 10 to 400.
     *
     * If not set, the default value is 100.
     * @note This method sets the parameter of _the last added view_.
     *
     * @note To set the view scales for specific view types see
     * SheetView::zoomScaleNormal, SheetView::zoomScalePageBreakView,
     * SheetView::zoomScalePageLayoutView.
     */
    void setViewZoomScale(int scale);
    /**
     * @brief returns the workbook view id that the last view of this sheet is
     * associated with.
     *
     * Usually a workbook has only one view, so this method
     * will usually return 0. But to define a complex view relations between a
     * workbook and a sheet you can use the WorkbookView and SheetView methods.
     * @return an integer id of a workbook view or -1 if there's no views
     * defined in the sheet or in the workbook.
     */
    int workbookViewId() const;
    /**
     * @brief sets the workbook view id that the last view of this sheet is
     * associated with.
     *
     * This method is useful only if you have a number of workbook views defined
     * and you want to associate this particular sheet with a specific workbook
     * view.
     *
     * @param id (usually) an index of the workbook view defined in the
     * #workbook().
     *
     * The default value of the workbook view id is 0 (the sheet view is
     * associated with the default workbook view).
     */
    void setWorkbookViewId(int id);
    /**
     * @brief returns the sheet view with the (zero-based) @a index.
     * @param index zero-based index of the sheet view.
     * @return the sheet view with the (zero-based) @a index. If no such view is
     * found, returns the default-constructed one.
     */
    SheetView view(int index) const;
    /**
     * @overload
     * @brief returns the sheet view with the (zero-based) @a index.
     * @param index non-negative index of the view. If the index is invalid,
     * the behaviour is undefined.
     * @return reference to the sheet view with the (zero-based) @a index.
     * @note Reference can be invalidated after adding or removing views.
     */
    SheetView &view(int index);
    /**
     * @brief returns the last added sheet view.
     * @return A copy of SheetView, that can be invalid if no view was defined
     * in the sheet.
     */
    SheetView lastView() const;
    /**
     * @brief returns the last added sheet view.
     * @return a reference to the SheetView object.
     *
     * In no view was defined in the sheet, this method creates and adds one.
     * @warning Reference can be invalidated after adding or removing views. Use
     * #view(int index).
     */
    SheetView &lastView();
    /**
     * @brief returns the count of sheet views defined in the worksheet.
     */
    int viewsCount() const;
    /**
     * @brief adds a new sheet view.
     * @return index of the added view.
     */
    int addView();
    /**
     * @brief removes the sheet view with @a index.
     * @param index non-negative index of the view.
     * @return `true` if the view was found and removed, `false` otherwise.
     * @note If you remove the only view, new default view will be created on
     * document write.
     */
    bool removeView(int index);

    SERIALIZE_ENUM(Visibility, {
        {Visibility::Hidden,     "hidden"},
        {Visibility::VeryHidden, "veryHidden"},
        {Visibility::Visible,    "visible"}
    });
protected:

    friend class Workbook;
    AbstractSheet(const QString &sheetName, int sheetId, Workbook *book, AbstractSheetPrivate *d);
    /**
     * @brief Copies the current sheet to a sheet called @a distName with @a distId.
     * Returns the new sheet.
     */
    virtual AbstractSheet *copy(const QString &distName, int distId) const = 0;

    void setType(Type type);
    int id() const;

    Drawing *drawing() const;
};

}
#endif // XLSXABSTRACTSHEET_H
