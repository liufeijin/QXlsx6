// xlsxabstractsheet.h

#ifndef XLSXABSTRACTSHEET_H
#define XLSXABSTRACTSHEET_H

#include "xlsxglobal.h"
#include "xlsxabstractooxmlfile.h"
#include "xlsxmain.h"
#include "xlsxsheetview.h"

#include <QXmlStreamWriter>
#include <QXmlStreamReader>

namespace QXlsx {

/**
 * @brief The HeaderFooter class represents the header and footer of a worksheet, chartsheet
 * or their custom views.
 *
 * Headers and footers may include formatting codes that specify the position and value of their
 * parts.
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
 * &G | picture as a background | |
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
     * When true then oddHeader/oddFooter and evenHeader/evenFooter specify the page header and
     * footer values for odd and even pages, respectively. If false then oddHeader/oddFooter
     * is used, even when evenHeader/evenFooter are present.
     * @note If differentOddEven is true, but no #evenHeader and #evenFooter is set,
     * the document is ill-formed.
     * The default value is false.
     */
    std::optional<bool> differentOddEven;
    /**
     * @brief Different first-page header and footer indicator.
     *
     * When true then firstHeader and firstFooter specify the first page header
     * and footer values, respectively. If false and firstHeader/firstFooter are
     * present, they are ignored.
     * @note If differentFirst is true, but no #firstHeader and #firstFooter is set,
     * the document is ill-formed.
     * The default value is false.
     */
    std::optional<bool> differentFirst;
    /**
     * @brief Scale header and footer with document scaling.
     */
    std::optional<bool> scaleWithDoc;
    /**
     * @brief Align header footer margins with page margins.
     * When true, as left/right  margins grow and shrink, the header and footer
     * edges stay aligned with the margins. When false, headers and footers are
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
    void write(QXmlStreamWriter &writer) const;
    void read(QXmlStreamReader &reader);
};

/**
 * @brief The PageMargins class represents the margins of a sheet page.
 * Margins are specified in inches or millimetres.
 * @note If you set any of the margins, you need to also set all the other margins,
 * otherwise the document will be ill-formed. Any missing margins will be written as zeroes.
 */
class QXLSX_EXPORT PageMargins
{
public:
    enum class Position
    {
        Top,
        Bottom,
        Left,
        Right,
        Header,
        Footer
    };
    /**
     * @brief creates invalid (default) page margins.
     *
     * To create valid page margins use #setMarginsInches(), #setMarginsMm() or
     * static methods #pageMarginsMm(), #pageMarginsInches().
     */
    PageMargins() {}
    /**
     * @brief sets page margins specified in inches.
     * @param left left page margin in inches.
     * @param top top page margin in inches.
     * @param right right page margin in inches.
     * @param bottom bottom page margin in inches.
     * @param header header page margin in inches.
     * @param footer footer page margin in inches.
     */
    void setMarginsInches(double left = 0, double top = 0, double right = 0, double bottom = 0,
                          double header = 0, double footer = 0);
    void setMarginInches(Position pos, double value);
    void setLeftMarginInches(double value);
    void setRightMarginInches(double value);
    void setTopMarginInches(double value);
    void setBottomMarginInches(double value);
    void setHeaderMarginInches(double value);
    void setFooterMarginInches(double value);
    /**
     * @brief sets page margins specified in millimeters.
     * @param left left page margin in millimeters.
     * @param top top page margin in millimeters.
     * @param right right page margin in millimeters.
     * @param bottom bottom page margin in millimeters.
     * @param header header page margin in millimeters.
     * @param footer footer page margin in millimeters.
     */
    void setMarginsMm(double left = 0, double top = 0, double right = 0, double bottom = 0,
                      double header = 0, double footer = 0);
    void setMarginMm(Position pos, double value);
    void setLeftMarginMm(double value);
    void setRightMarginMm(double value);
    void setTopMarginMm(double value);
    void setBottomMarginMm(double value);
    void setHeaderMarginMm(double value);
    void setFooterMarginMm(double value);

    /**
     * @brief creates default page margins.
     */
    static PageMargins defaultPageMargins();
    /**
     * @brief creates page margins from values specified in millimeters.
     * @param left left page margin in millimeters.
     * @param top top page margin in millimeters.
     * @param right right page margin in millimeters.
     * @param bottom bottom page margin in millimeters.
     * @param header header page margin in millimeters.
     * @param footer footer page margin in millimeters.
     */
    static PageMargins pageMarginsMm(double left = 0, double top = 0, double right = 0, double bottom = 0,
                                     double header = 0, double footer = 0);
    /**
     * @brief creates page margins from values specified in inches.
     * @param left left page margin in inches.
     * @param top top page margin in inches.
     * @param right right page margin in inches.
     * @param bottom bottom page margin in inches.
     * @param header header page margin in inches.
     * @param footer footer page margin in inches.
     * @note To set pageMargins in millimetres use #pageMarginsMm().
     */
    static PageMargins pageMarginsInches(double left = 0, double top = 0, double right = 0, double bottom = 0,
                                         double header = 0, double footer = 0);

    /**
     * @brief returns page margins in inches.
     * @return
     */
    QMap<Position, double> marginsInches() const;
    /**
     * @brief returns page margins in millimetres.
     * @return
     */
    QMap<Position, double> marginsMm() const;
    double marginInches(Position pos) const;
    double leftMarginInches() const;
    double rightMarginInches() const;
    double topMarginInches() const;
    double bottomMarginInches() const;
    double headerMarginInches() const;
    double footerMarginInches() const;
    double marginMm(Position pos) const;
    double leftMarginMm() const;
    double rightMarginMm() const;
    double topMarginMm() const;
    double bottomMarginMm() const;
    double headerMarginMm() const;
    double footerMarginMm() const;
    bool hasMargin(Position pos) const;

    /**
     * @brief returns whether page margins are valid, that is whether at least one margin is set.
     * @note All 6 margins must be set according to the ECMA376. Missing margins will be written
     * as zeroes.
     * @return true if at least one margin is set, false otherwise.
     */
    bool isValid() const;
    void write(QXmlStreamWriter &writer) const;
    void read(QXmlStreamReader &reader);
private:
    QMap<Position, double> mMargins;
};

/**
 * @brief The PageSetup class represents paper parameters when printing the sheet.
 *
 * Not all parameters are applicable to chartsheets. See the description of parameters.
 * If a non-applicable parameter is set for a chartsheet, it will be ignored
 * when saving the document.
 */
class QXLSX_EXPORT PageSetup
{
//TODO: setCustomPrinter via relations (printerId)
public:
    /**
     * @brief The PaperSize enum specifies possible paper size variants.
     * The application can use custom paper sizes, but they are read and subsequently
     * written as zeroes.
     */
    enum class PaperSize
    {
        Unknown = 0, /**< Unknown paper size */
        Letter = 1, /**< Letter paper (8.5 in. by 11 in.)  */
        LetterSmall, /**< Letter small paper (8.5 in. by 11 in.)  */
        Tabloid, /**< Tabloid paper (11 in. by 17 in.)  */
        Ledger, /**< Ledger paper (17 in. by 11 in.)  */
        Legal, /**< Legal paper (8.5 in. by 14 in.)  */
        Statement, /**< Statement paper (5.5 in. by 8.5 in.)  */
        Executive, /**< Executive paper (7.25 in. by 10.5 in.)  */
        A3, /**< A3 paper (297 mm by 420 mm)  */
        A4, /**< A4 paper (210 mm by 297 mm)  */
        A4Small, /**< A4 small paper (210 mm by 297 mm)  */
        A5, /**< A5 paper (148 mm by 210 mm)  */
        B4, /**< B4 paper (250 mm by 353 mm)  */
        B5, /**< B5 paper (176 mm by 250 mm)  */
        Folio, /**< olio paper (8.5 in. by 13 in.)  */
        Quarto, /**< Quarto paper (215 mm by 275 mm)  */
        Standard_10_14, /**< Standard paper (10 in. by 14 in.)  */
        Standard_11_17, /**< Standard paper (11 in. by 17 in.)  */
        Note, /**< Note paper (8.5 in. by 11 in.)  */
        Envelope9, /**< #9 envelope (3.875 in. by 8.875 in.)  */
        Envelope10, /**< #10 envelope (4.125 in. by 9.5 in.)*/
        Envelope11, /**< #11 envelope (4.5 in. by 10.375 in.)  */
        Envelope12, /**< #12 envelope (4.75 in. by 11 in.)  */
        Envelope14, /**< #14 envelope (5 in. by 11.5 in.)  */
        C, /**< C paper (17 in. by 22 in.)  */
        D, /**< D paper (22 in. by 34 in.)  */
        E, /**< E paper (34 in. by 44 in.)  */
        EnvelopeDL , /**< DL envelope (110 mm by 220 mm)  */
        EnvelopeC5 , /**< C5 envelope (162 mm by 229 mm)  */
        EnvelopeC3 , /**< C3 envelope (324 mm by 458 mm)  */
        EnvelopeC4 , /**< C4 envelope (229 mm by 324 mm)  */
        EnvelopeC6 , /**< C6 envelope (114 mm by 162 mm)  */
        EnvelopeC65 , /**< C65 envelope (114 mm by 229 mm)  */
        EnvelopeB4 , /**< B4 envelope (250 mm by 353 mm)  */
        EnvelopeB5 , /**< B5 envelope (176 mm by 250 mm)  */
        EnvelopeB6 , /**< B6 envelope (176 mm by 125 mm)  */
        EnvelopeItaly , /**< Italy envelope (110 mm by 230 mm)  */
        EnvelopeMonarch , /**< Monarch envelope (3.875 in. by 7.5 in.).  */
        Envelope634 , /**< 6 3/4 envelope (3.625 in. by 6.5 in.)  */
        USStandardFanfold, /**< US standard fanfold (14.875 in. by 11 in.)  */
        GermanStandardFanfold, /**< German standard fanfold (8.5 in. by 12 in.)  */
        GermanLegalFanfold, /**< German legal fanfold (8.5 in. by 13 in.)  */
        B4ISO, /**< ISO B4 (250 mm by 353 mm)  */
        JapaneseDoublePostcard, /**< Japanese double postcard (200 mm by 148 mm)  */
        Standard_9_11, /**< Standard paper (9 in. by 11 in.) */
        Standard_10_11, /**< Standard paper (10 in. by 11 in.)  */
        Standard_15_11, /**< Standard paper (15 in. by 11 in.)  */
        EnvelopeInvite, /**< Invite envelope (220 mm by 220 mm)  */
        LetterExtra = 50, /**< Letter extra paper (9.275 in. by 12 in.)  */
        LegalExtra, /**< Legal extra paper (9.275 in. by 15 in.)  */
        TabloidExtra, /**< Tabloid extra paper (11.69 in. by 18 in.)  */
        A4Extra, /**< A4 extra paper (236 mm by 322 mm)  */
        LetterTransverse, /**< Letter transverse paper (8.275 in. by 11 in.)  */
        A4Transverse, /**< A4 transverse paper (210 mm by 297 mm)  */
        LetterExtraTransverse, /**< Letter extra transverse paper (9.275 in. by 12 in.)  */
        SuperA, /**< SuperA/SuperA/A4 paper (227 mm by 356 mm)  */
        SuperB, /**< SuperB/SuperB/A3 paper (305 mm by 487 mm)  */
        LetterPlus, /**< Letter plus paper (8.5 in. by 12.69 in.)  */
        A4Plus, /**< A4 plus paper (210 mm by 330 mm)  */
        A5Transverse, /**< A5 transverse paper (148 mm by 210 mm)  */
        JISB5Transverse, /**< JIS B5 transverse paper (182 mm by 257 mm)  */
        A3Extra, /**< A3 extra paper (322 mm by 445 mm)  */
        A5Extra, /**< A5 extra paper (174 mm by 235 mm)  */
        ISOB5Extra, /**< ISO B5 extra paper (201 mm by 276 mm)  */
        A2, /**< A2 paper (420 mm by 594 mm)  */
        A3Transverse, /**< A3 transverse paper (297 mm by 420 mm)  */
        A3ExtraTransverse, /**< A3 extra transverse paper (322 mm by 445 mm)*/
        JapaneseDoublePostcard1, /**< Japanese Double Postcard (200 mm x 148 mm) */
        A6, /**< A6 (105 mm x 148 mm) */
        JapaneseEnvelopeKaku2, /**< Japanese Envelope Kaku #2 */
        JapaneseEnvelopeKaku3 , /**< Japanese Envelope Kaku #3 */
        JapaneseEnvelopeChou3, /**< Japanese Envelope Chou #3 */
        JapaneseEnvelopeChou4, /**< Japanese Envelope Chou #4 */
        LetterRotated, /**< Letter Rotated (11in x 8 1/2 11 in) */
        A3Rotated, /**< A3 Rotated (420 mm x 297 mm) */
        A4Rotated, /**< A4 Rotated (297 mm x 210 mm) */
        A5Rotated, /**< A5 Rotated (210 mm x 148 mm) */
        B4JISRotated, /**< B4 (JIS) Rotated (364 mm x 257 mm) */
        B5JISRotated, /**< B5 (JIS) Rotated (257 mm x 182 mm) */
        JapanesePostcardRotated, /**< Japanese Postcard Rotated (148 mm x 100 mm) */
        DoubleJapanesePostcardRotated, /**< Double Japanese Postcard Rotated (148 mm x 200 mm) */
        A6Rotated, /**< A6 Rotated (148 mm x 105 mm) */
        JapaneseEnvelopeKaku2Rotated, /**< Japanese Envelope Kaku #2 Rotated */
        JapaneseEnvelopeKaku3Rotated, /**< Japanese Envelope Kaku #3 Rotated */
        JapaneseEnvelopeChou3Rotated, /**< Japanese Envelope Chou #3 Rotated */
        JapaneseEnvelopeChou4Rotated, /**< Japanese Envelope Chou #4 Rotated */
        B6JIS, /**< B6 (JIS) (128 mm x 182 mm) */
        B6JISRotated, /**< B6 (JIS) Rotated (182 mm x 128 mm) */
        Standard_12_11, /**< (12 in x 11 in) */
        JapaneseEnvelopeYou4, /**< Japanese Envelope You #4 */
        JapaneseEnvelopeYou4Rotated, /**< Japanese Envelope You #4 Rotated */
        PRC16K, /**< PRC 16K (146 mm x 215 mm) */
        PRC32K, /**< PRC 32K (97 mm x 151 mm) */
        PRC32KBig, /**< PRC 32K(Big) (97 mm x 151 mm) */
        PRCEnvelope1, /**< PRC Envelope #1 (102 mm x 165 mm) */
        PRCEnvelope2, /**< PRC Envelope #2 (102 mm x 176 mm) */
        PRCEnvelope3, /**< PRC Envelope #3 (125 mm x 176 mm) */
        PRCEnvelope4, /**< PRC Envelope #4 (110 mm x 208 mm) */
        PRCEnvelope5, /**< PRC Envelope #5 (110 mm x 220 mm) */
        PRCEnvelope6, /**< PRC Envelope #6 (120 mm x 230 mm) */
        PRCEnvelope7, /**< PRC Envelope #7 (160 mm x 230 mm) */
        PRCEnvelope8, /**< PRC Envelope #8 (120 mm x 309 mm) */
        PRCEnvelope9, /**< PRC Envelope #9 (229 mm x 324 mm) */
        PRCEnvelope10, /**< PRC Envelope #10 (324 mm x 458 mm) */
        PRC16KRotated, /**< PRC 16K Rotated */
        PRC32KRotated, /**< PRC 32K Rotated */
        PRC32KBigRotated, /**< PRC 32K(Big) Rotated */
        PRCEnvelope1Rotated, /**< PRC Envelope #1 Rotated (165 mm x 102 mm) */
        PRCEnvelope2Rotated, /**< PRC Envelope #2 Rotated (176 mm x 102 mm) */
        PRCEnvelope3Rotated, /**< PRC Envelope #3 Rotated (176 mm x 125 mm) */
        PRCEnvelope4Rotated, /**< PRC Envelope #4 Rotated (208 mm x 110 mm) */
        PRCEnvelope5Rotated, /**< PRC Envelope #5 Rotated (220 mm x 110 mm) */
        PRCEnvelope6Rotated, /**< PRC Envelope #6 Rotated (230 mm x 120 mm) */
        PRCEnvelope7Rotated, /**< PRC Envelope #7 Rotated (230 mm x 160 mm) */
        PRCEnvelope8Rotated, /**< PRC Envelope #8 Rotated (309 mm x 120 mm) */
        PRCEnvelope9Rotated, /**< PRC Envelope #9 Rotated (324 mm x 229 mm) */
        PRCEnvelope10Rotated, /**< PRC Envelope #10 Rotated (458 mm x 324 mm)*/
    };

    /**
     * @brief The PageOrder enum specifies the order of printed pages.
     */
    enum class PageOrder
    {
        DownThenOver, /**< Order pages vertically first, then move horizontally */
        OverThenDown /**< Order pages horizontally first, then move vertically */
    };
    /**
     * @brief The Orientation enum specifies the page orientation.
     */
    enum class Orientation
    {
        Default, /**< Default orientation */
        Portrait, /**< Portrait orientation when printing */
        Landscape /**< Landscape orientation when printing */
    };
    /**
     * @brief The PrintError enum specifies how to print cell values for cells with errors.
     */
    enum class PrintError
    {
        Displayed, /**< Display cell errors as displayed on screen.*/
        Blank, /**< Display cell errors as blank.*/
        Dash, /**< Display cell errors as dashes.*/
        NotAvailable /**< Display cell errors as N/A.*/
    };
    /**
     * @brief The CellComments enum specifies how to print cell comments.
     */
    enum class CellComments
    {
        AsDisplayed, /**< Print cell comments as displayed. */
        AtEnd, /**< Print cell comments at the end of the document. */
        DoNotPrint /**< Do not print cell comments. */
    };

    /**
     * @brief specifies paper size.
     *
     * If both #paperWidth and #paperHeight are specified, #paperSize is ignored.
     *
     * If no value is specified, then PaperSize::Letter is assumed.
     * @sa AbstractSheet::setPaperSize(), AbstractSheet::paperSize().
     */
    std::optional<PaperSize> paperSize;
    /**
     * @brief specifies paper width as a universal measure (f.e. "210mm").
     *
     * The string should have the format "-?[0-9]+(\.[0-9]+)?(mm|cm|in|pt|pc|pi)".
     *
     * If both #paperWidth and #paperHeight are specified (not empty), #paperSize is ignored.
     * @sa AbstractSheet::setPaperSize(), AbstractSheet::setPaperSizeMM(), AbstractSheet::setPaperSizeInches(),
     * AbstractSheet::paperWidth(), AbstractSheet::paperHeight()
     */
    QString paperWidth;
    /**
     * @brief specifies paper height as a universal measure (f.e. "210mm").
     *
     * The string should have the format "-?[0-9]+(\.[0-9]+)?(mm|cm|in|pt|pc|pi)".
     *
     * If both #paperWidth and #paperHeight are specified (not empty), #paperSize is ignored.
     * @sa AbstractSheet::setPaperSize(), AbstractSheet::setPaperSizeMM(), AbstractSheet::setPaperSizeInches(),
     * AbstractSheet::paperWidth(), AbstractSheet::paperHeight()
     */
    QString paperHeight;
    /**
     * @brief specifies print scale in percents.
     *
     * The value of 100 equals 100% scaling. The value should be restricted ranging from 10 to 400.
     *
     * If no value is specified, then 100 is assumed.
     *
     * If #fitToWidth and/or #fitToHeight are specified, #scale is ignored.
     * @note This parameter is not applicable to chartsheets.
     * @sa Worksheet::setPrintScale(), Worksheet::printScale().
     */
    std::optional<int> scale;
    /**
     * @brief Page number for first printed page.
     *
     * If no value is specified, then 1 is assumed.
     * @sa AbstractSheet::setFirstPageNumber(), AbstractSheet::firstPageNumber().
     */
    std::optional<int> firstPageNumber;
    /**
     * @brief Number of horizontal pages to fit on.
     *
     * If no value is specified, then 1 is assumed.
     *
     * If #fitToWidth and/or #fitToHeight are specified, #scale is ignored.
     * @note This parameter is not applicable to chartsheets.
     * @sa AbstractSheet::setFitToWidth(), AbstractSheet::fitToWidth()
     */
    std::optional<int> fitToWidth;
    /**
     * @brief Number of vertical pages to fit on.
     *
     * If no value is specified, then 1 is assumed.
     *
     * If #fitToWidth and/or #fitToHeight are specified, #scale is ignored.
     * @note This parameter is not applicable to chartsheets.
     * @sa AbstractSheet::setFitToHeight(), AbstractSheet::fitToHeight()
     */
    std::optional<int> fitToHeight;
    /**
     * @brief specifies the order of printed pages.
     *
     * If no value is specified, then PageOrder::DownThenOver is assumed.
     * @note This parameter is not applicable to chartsheets.
     * @sa Worksheet::setPageOrder(), Worksheet::pageOrder().
     */
    std::optional<PageOrder> pageOrder;
    /**
     * @brief specifies the page orientation.
     *
     * If no value is specified, then Orientation::Default is assumed.
     * @sa AbstractSheet::setPageOrientation(), AbstractSheet::pageOrientation().
     */
    std::optional<Orientation> orientation;
    /**
     * @brief specifies whether to use the printer default parameters.
     *
     * Use the printer’s defaults settings for page setup values and don't use the
     * default values of the file. Example: if dpi is not present or specified in
     * the file, the application must not assume 600dpi as specified as a default
     * and instead must let the printer specify the default dpi.
     *
     * If no value is specified, then true is assumed.
     */
    std::optional<bool> usePrinterDefaults;
    /**
     * @brief specifies whether to print in black and white.
     *
     * If no value is specified, then false is assumed.
     */
    std::optional<bool> blackAndWhite;
    /**
     * @brief specifies whether to print the page as a draft (i.e. without graphics).
     *
     * If no value is specified, then false is assumed.
     */
    std::optional<bool> draft;
    /**
     * @brief specifies whether to use #firstPageNumber value for first page number,
     * and do not auto number the pages.
     *
     * If no value is specified, then false is assumed.
     */
    std::optional<bool> useFirstPageNumber;
    /**
     * @brief specifies horizontal print resolution of the device.
     *
     * If no value is specified, then 600 is assumed.
     */
    std::optional<int> horizontalDpi;
    /**
     * @brief specifies vertical print resolution of the device.
     *
     * If no value is specified, then 600 is assumed.
     */
    std::optional<int> verticalDpi;
    /**
     * @brief specifies how many copies to print.
     *
     * If no value is specified, then 1 is assumed.
     */
    std::optional<int> copies;
    /**
     * @brief Relationship Id of the devMode printer settings part.
     */
    QString printerId; //TODO: adding printer
    /**
     * @brief  specifies how to display cells with errors when printing the worksheet.
     *
     * If no value is specified, then PrintError::Displayed is assumed.
     * @note This parameter is not applicable to chartsheets.
     */
    std::optional<PrintError> errors;
    /**
     * @brief specifies how cell comments shall be printed.
     *
     * If no value is specified, then CellComments::DoNotPrint is assumed.
     */
    std::optional<CellComments> cellComments;

    bool isValid() const;
    void write(QXmlStreamWriter &writer, bool chartsheet = false) const;
    void read(QXmlStreamReader &reader);
private:
    void readPaperSize(QXmlStreamReader &reader);
};

/**
 * @brief The SheetProtection class specifies the sheet protection parameters for worksheets
 * and chartsheets.
 *
 * Not all parameters are applicable to chartsheets. See the description of parameters.
 * If a non-applicable parameter is set for a chartsheet protection, it will be ignored
 * when saving the document.
 */
class QXLSX_EXPORT SheetProtection
{
public:
    /**
     * @brief Specifies the specific cryptographic hashing algorithm which shall
     * be used along with the salt attribute and input password in order to
     * compute the hash value.
     *
     * This is an optional parameter. If algorithmName is empty, it will not be stored in the doc.
     *
     * The following names are reserved:
     *
     * Algorithm | Description
     * ----|----
     * MD2   |Specifies that the MD2 algorithm, as defined by RFC 1319, shall be used.
     * MD4 |  Specifies that the MD4 algorithm, as defined by RFC 1320, shall be used.
     * MD5 | Specifies that the MD5 algorithm, as defined by RFC 1321, shall be used.
     * RIPEMD-128 | Specifies that the RIPEMD-128 algorithm, as defined by ISO/IEC 10118-3:2004 shall be used.
     * RIPEMD-160 | Specifies that the RIPEMD-160 algorithm, as defined by ISO/IEC 10118-3:2004 shall be used.
     * SHA-1 | Specifies that the SHA-1 algorithm, as defined by ISO/IEC 10118-3:2004 shall be used.
     * SHA-256 | Specifies that the SHA-256 algorithm, as defined by ISO/IEC 10118-3:2004 shall be used.
     * SHA-384 | Specifies that the SHA-384 algorithm, as defined by ISO/IEC 10118-3:2004  shall be used.
     * SHA-512 | Specifies that the SHA-512 algorithm, as defined by ISO/IEC 10118-3:2004 shall be used.
     * WHIRLPOOL | Specifies that the WHIRLPOOL algorithm, as defined by ISO/IEC 10118-3:2004 shall be used.
     */
    QString algorithmName;
    /**
     * @brief Specifies the hash value for the password required to edit this
     * worksheet.
     *
     * This value shall be compared with the resulting hash value after hashing
     * the user-supplied password using the #algorithmName, and if the two values
     * match, the protection shall no longer be enforced.
     *
     * hashValue is specified as a base64 string.
     *
     * This is an optional parameter. If hashValue is empty, it will not be stored in the doc.
     */
    QString hashValue;
    /**
     * @brief Specifies the salt which was prepended to the user-supplied
     * password before it was hashed using #algorithmName to generate #hashValue,
     * and which shall also be prepended to the user-supplied password before
     * attempting to generate a hash value for comparison.
     *
     * A salt is a random string which is added to a user-supplied password
     * before it is hashed in order to prevent a malicious party from
     * pre-calculating all possible password/hash combinations and simply using
     * those pre-calculated values (often referred to as a "dictionary attack").
     *
     * saltValue is specified as a base64 string.
     *
     * This is an optional parameter. If saltValue is empty, it will not be stored in the doc.
     */
    QString saltValue;
    /**
     * @brief Specifies the number of times the hashing function shall be
     * iteratively run (runs using each iteration's result plus a 4 byte value
     * (0-based, little endian) containing the number of the iteration as the
     * input for the next iteration) when attempting to compare a user-supplied
     * password with the value stored in #hashValue.
     */
    std::optional<int> spinCount;
    /**
     * @brief Specifies whether editing of sheet content should not be allowed when the chartsheet is protected.
     *
     * If not set, the default value is false.
     * @note This parameter is only applicable to chartsheets.
     */
    std::optional<bool> protectContent;
    /**
     * @brief Specifies whether editing of objects should not be allowed when the sheet is protected.
     *
     * If not set, the default value is false.
     */
    std::optional<bool> protectObjects;
    /**
     * @brief This parameter dictates whether the other attributes of
     * SheetProtection should be applied.
     *
     * If true then the other parameters of SheetProtection should be applied.
     * If false then the other parameters of SheetProtection should not be applied.
     *
     * If not set, the default value is false.
     * @note This parameter is not applicable to chartsheets.
     *
     */
    std::optional<bool> protectSheet;
    /**
     * @brief Specifies whether editing of scenarios should not be allowed when
     * the sheet is protected.
     *
     * If not set, the default value is false.
     * @note This parameter is not applicable to chartsheets.
     */
    std::optional<bool> protectScenarios;
    /**
     * @brief Specifies whether editing of cells formattins should not be
     * allowed when the sheet is protected.
     *
     * If not set, the default value is true.
     * @note This parameter is not applicable to chartsheets.
     */
    std::optional<bool> protectFormatCells;
    /**
     * @brief Specifies whether editing of columns formatting should not be
     * allowed when the sheet is protected.
     *
     * If not set, the default value is true.
     * @note This parameter is not applicable to chartsheets.
     */
    std::optional<bool> protectFormatColumns;
    /**
     * @brief Specifies whether editing of rows formatting should not be allowed
     * when the sheet is protected.
     *
     * If not set, the default value is true.
     * @note This parameter is not applicable to chartsheets.
     */
    std::optional<bool> protectFormatRows;
    /**
     * @brief Specifies whether inserting of columns should not be allowed when
     * the sheet is protected.
     *
     * If not set, the default value is true.
     * @note This parameter is not applicable to chartsheets.
     */
    std::optional<bool> protectInsertColumns;
    /**
     * @brief Specifies whether inserting of rows should not be allowed when the
     * sheet is protected.
     *
     * If not set, the default value is true.
     * @note This parameter is not applicable to chartsheets.
     */
    std::optional<bool> protectInsertRows;
    /**
     * @brief Specifies whether inserting of hyperlinks should not be allowed
     * when the sheet is protected.
     *
     * If not set, the default value is true.
     * @note This parameter is not applicable to chartsheets.
     */
    std::optional<bool> protectInsertHyperlinks;
    /**
     * @brief Specifies whether deleting of columns should not be allowed when
     * the sheet is protected.
     *
     * If not set, the default value is true.
     * @note This parameter is not applicable to chartsheets.
     */
    std::optional<bool> protectDeleteColumns;
    /**
     * @brief Specifies whether deleting of rows should not be allowed when the
     * sheet is protected.
     *
     * If not set, the default value is true.
     * @note This parameter is not applicable to chartsheets.
     */
    std::optional<bool> protectDeleteRows;
    /**
     * @brief Specifies whether selection of locked cells should not be allowed
     * when the sheet is protected.
     *
     * If not set, the default value is false.
     * @note This parameter is not applicable to chartsheets.
     */
    std::optional<bool> protectSelectLockedCells;
    /**
     * @brief Specifies whether sorting should not be allowed when the sheet is protected.
     *
     * If not set, the default value is true.
     * @note This parameter is not applicable to chartsheets.
     */
    std::optional<bool> protectSort;
    /**
     * @brief Specifies whether applying autofilters should not be allowed when
     * the sheet is protected.
     *
     * If not set, the default value is true.
     * @note This parameter is not applicable to chartsheets.
     */
    std::optional<bool> protectAutoFilter;
    /**
     * @brief Specifies whether operating pivot tables should not be allowed
     * when the sheet is protected.
     *
     * If not set, the default value is true.
     * @note This parameter is not applicable to chartsheets.
     */
    std::optional<bool> protectPivotTables;
    /**
     * @brief Specifies whether selection of unlocked cells should not be
     * allowed when the sheet is protected.
     *
     * If not set, the default value is false.
     * @note This parameter is not applicable to chartsheets.
     */
    std::optional<bool> protectSelectUnlockedCells;

    /**
     * @brief returns whether any of protection parameters are set.
     * @return
     */
    bool isValid() const;
    void write(QXmlStreamWriter &writer, bool chartsheet = false) const;
    void read(QXmlStreamReader &reader);
};

class Workbook;
class Drawing;
class AbstractSheetPrivate;

/**
 * @brief The AbstractSheet class is a base class for different sheet types.
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
     * @param newName
     * @return true if renaming was successful.
     */
    bool rename(const QString &newName);

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
     * @return true if the sheet's #visibility() is Visibility::Hidden or Visibility::VeryHidden,
     * false otherwise.
     * This is a convenience method, equivalent to @code visibility() == Visibility::Hidden || visibility() == Visibility::VeryHidden @endcode
     * By default sheets are visible.
     * @sa #visibility(), #isVisible().
     */
    bool isHidden() const;
    /**
     * @brief returns whether the sheet is visible.
     * @return true if the sheet's #visibility() is Visibility::Visible, false otherwise.
     * This is a convenience method, equivalent to @code visibility() == Visibility::Visible @endcode
     * By default sheets are visible.
     * @sa #isHidden(), #visibility().
     */
    bool isVisible() const;
    /**
     * @brief sets the sheet's visibility.
     * @param hidden if true, the sheet's visibility will be set to Visibility::Hidden,
     * if false, the sheet will be Visibility::Visible.
     * @note If @a hidden is true, the very hidden sheet stays very hidden.
     * You can check the exact visibility with #visibility().
     * @sa #setVisibility(), #setVisible().
     */
    void setHidden(bool hidden);
    /**
     * @brief sets the sheet's visibility.
     * @param visible if true, the sheet's visibility will be set to Visibility::Visible,
     * if false, the sheet will be Visibility::Hidden.
     * @note If @a visible is false, the very hidden sheet stays very hidden.
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

    //TODO: add methods to fine-tune header and footer
    /**
     * @brief returns whether a sheet has different header/footer for odd and even pages.
     * @return When true then #oddHeader() / #oddFooter() and #evenHeader() /
     * #evenFooter() specify the page header and footer values for odd and even
     * pages, respectively. If false then #oddHeader() / #oddFooter() or #header() / #footer()
     * are used, even when #evenHeader() / #evenFooter() are not empty.
     *
     * The default value is false.
     *
     * @sa #header(), #footer(), #oddHeader(), #oddFooter(), #evenHeader(), #evenFooter(),
     * #setHeader(), #setFooter(), #setOddHeader(), #setOddFooter(), #setEvenHeader(),
     * #setEvenFooter(), #setHeaders(), #setFooters().
     */
    bool differentOddEvenPage() const;
    /**
     * @brief sets whether a sheet has different header/footer for odd and even pages.
     * @param different When true then #oddHeader() / #oddFooter() and #evenHeader() /
     * #evenFooter() specify the page header and footer values for odd and even
     * pages, respectively. If false then #oddHeader() / #oddFooter() or #header() / #footer()
     * are used, even when #evenHeader() / #evenFooter() are not empty.
     *
     * If not set, the default value is false.
     *
     * @sa #header(), #footer(), #oddHeader(), #oddFooter(), #evenHeader(), #evenFooter(),
     * #setHeader(), #setFooter(), #setOddHeader(), #setOddFooter(), #setEvenHeader(),
     * #setEvenFooter(), #setHeaders(), #setFooters().
     */
    void setDifferentOddEvenPage(bool different);
    /**
     * @brief returns whether a sheet has different header/footer for the first page.
     * @return When true, then #firstHeader() / #firstFooter() specify the first page
     * header and footer. If false, then #firstHeader() and #firstFooter() are ignored
     * even if they are not empty.
     */
    bool differentFirstPage() const;
    /**
     * @brief sets whether a sheet has different header/footer for the first page.
     * @param different When true, then #firstHeader() / #firstFooter() specify the first page
     * header and footer.
     * @sa #setHeaders(), #setFooters().
     */
    void setDifferentFirstPage(bool different);
    /**
     * @brief returns the page header.
     * @return Page header as a string that may include formatting codes. See HeaderFooter class
     * for a list of formatting codes. The return value is empty if #differentOddEvenPage()
     * and/or #differentFirstPage() is true. In this case use #oddHeader(), #evenHeader(), #firstHeader().
     */
    QString header() const;
    /**
     * @brief sets the page header.
     * @param header a string that may include formatting codes. See HeaderFooter class
     * for a list of formatting codes.
     *
     * This method effectively sets odd/even pages and first page to false. If needed, use
     * #setOddHeader(), #setEvenHeader(), #setFirstHeader().
     */
    void setHeader(const QString &header);
    /**
     * @brief returns the header of odd pages.
     * @return a string that may include formatting codes. See HeaderFooter class
     * for a list of formatting codes.
     *
     * The return value is empty if #differentOddEvenPage()
     * is false.
     */
    QString oddHeader() const;
    /**
     * @brief sets the header of odd pages.
     * @param header a string that may include formatting codes. See HeaderFooter class
     * for a list of formatting codes.
     *
     * This method effectively sets odd/even pages to true. Note that if you don't specify
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
     * is false.
     */
    QString evenHeader() const;
    /**
     * @brief sets the header of even pages.
     * @param header a string that may include formatting codes. See HeaderFooter class
     * for a list of formatting codes.
     *
     * This method effectively sets odd/even pages to true. Note that if you don't specify
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
     * is false.
     */
    QString firstHeader() const;
    /**
     * @brief sets the header of the first page.
     * @param header a string that may include formatting codes. See HeaderFooter class
     * for a list of formatting codes.
     *
     * This method effectively sets first page to true.
     * @sa #setHeaders().
     */
    void setFirstHeader(const QString &header);
    /**
     * @brief sets headers for page
     * @param oddHeader non-empty string that is used for headers if \a evenHeader is empty,
     * and for odd page header if \a evenHeader is not empty.
     * @param evenHeader a string that is used for even page header. If not empty, then
     * #setDifferentOddEvenPage() is set to true.
     * @param firstHeader a string that is used for the first page. If not empty,
     * then #setDifferentFirstPage() is set to true.
     */
    void setHeaders(const QString &oddHeader, const QString &evenHeader = QString(), const QString &firstHeader = QString());

    // footers
    /**
     * @brief returns the page footer.
     * @return Page footer as a string that may include formatting codes. See HeaderFooter class
     * for a list of formatting codes. The return value is empty if #differentOddEvenPage()
     * and/or #differentFirstPage() is true. In this case use #oddFooter(), #evenFooter(), #firstFooter().
     */
    QString footer() const;
    /**
     * @brief sets the page footer.
     * @param footer a string that may include formatting codes. See HeaderFooter class
     * for a list of formatting codes.
     *
     * This method effectively sets odd/even pages and first page to false. If needed, use
     * #setOddFooter(), #setEvenFooter(), #setFirstFooter().
     */
    void setFooter(const QString &footer);
    /**
     * @brief returns the footer of odd pages.
     * @return a string that may include formatting codes. See HeaderFooter class
     * for a list of formatting codes.
     *
     * The return value is empty if #differentOddEvenPage()
     * is false.
     */
    QString oddFooter() const;
    /**
     * @brief sets the footer of odd pages.
     * @param footer a string that may include formatting codes. See HeaderFooter class
     * for a list of formatting codes.
     *
     * This method effectively sets odd/even pages to true. Note that if you don't specify
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
     * is false.
     */
    QString evenFooter() const;
    /**
     * @brief sets the footer of even pages.
     * @param footer a string that may include formatting codes. See HeaderFooter class
     * for a list of formatting codes.
     *
     * This method effectively sets odd/even pages to true. Note that if you don't specify
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
     * is false.
     */
    QString firstFooter() const;
    /**
     * @brief sets the footer of the first page.
     * @param footer a string that may include formatting codes. See HeaderFooter class
     * for a list of formatting codes.
     *
     * This method effectively sets first page to true.
     * @sa #setFooters().
     */
    void setFirstFooter(const QString &footer);
    /**
     * @brief sets footers for page
     * @param oddFooter non-empty string that is used for footers if \a evenFooter is empty,
     * and for odd page footer if \a evenFooter is not empty.
     * @param evenFooter a string that is used for even page footer. If not empty, then
     * #setDifferentOddEvenPage() is set to true.
     * @param firstFooter a string that is used for the first page. If not empty,
     * then #setDifferentFirstPage() is set to true.
     */
    void setFooters(const QString &oddFooter, const QString &evenFooter = QString(), const QString &firstFooter = QString());

    // Page margins methods
    QMap<PageMargins::Position, double> pageMarginsInches() const;
    QMap<PageMargins::Position, double> pageMarginsMm() const;
    void setPageMarginsInches(double left = 0, double top = 0, double right = 0, double bottom = 0,
                              double header = 0, double footer = 0);
    void setPageMarginsMm(double left = 0, double top = 0, double right = 0, double bottom = 0,
                          double header = 0, double footer = 0);
    PageMargins pageMargins() const;
    PageMargins &pageMargins();
    void setPageMargins(const PageMargins &margins);
    /**
     * @brief sets default page margins, clearing all margins values.
     */
    void setDefaultPageMargins();

    // Page setup methods
    PageSetup pageSetup() const;
    PageSetup &pageSetup();
    /**
     * @brief sets the sheet's page parameters.
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
    PageSetup::PaperSize paperSize() const;
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
    PageSetup::Orientation pageOrientation() const;
    /**
     * @brief sets the sheet's paper orientation.
     * @param orientation PageSetup::Orientation value.
     */
    void setPageOrientation(PageSetup::Orientation orientation);


    //Sheet properties, common to both chartsheets and worksheets
    /**
     * @brief returns whether the sheet is published.
     * @return true if the sheet was published.
     *
     * If not set, the default value is true.
     */
    bool isPublished() const;
    /**
     * @brief sets whether the sheet is published.
     * @param published
     *
     * The default value is true.
     */
    void setPublished(bool published);
    /**
     * @brief returns the unique stable sheet name.
     *
     * codeName should not change over time, and does not change from user input.
     * This name should be used by code to reference a particular sheet.
     *
     * @return Non-empty string if the codeName was set.
     */
    QString codeName() const;
    /**
     * @brief sets the unique stable sheet name.
     *
     * codeName should not change over time, and does not change from user input.
     * This name should be used by code to reference a particular sheet.
     * @param a unique string.
     * @warning This method does not check the codeName to be unique throughout the workbook!
     * Use QUuid class to create a unique codeName.
     */
    void setCodeName(const QString &codeName);
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
     * @return true if the background image was removed, false otherwise.
     */
    bool removeBackgroundImage();


    /**
     * @brief returns the sheet protection parameters
     * @return
     */
    SheetProtection sheetProtection() const;
    /**
     * @brief returns the sheet protection parameters.
     * @return
     */
    SheetProtection &sheetProtection();
    /**
     * @brief sets the sheet protection parameters.
     * @param sheetProtection
     */
    void setSheetProtection(const SheetProtection &sheetProtection);
    /**
     * @brief returns whether the sheet is protected.
     * @return true if any of the protection parameters were set.
     * @sa #isPasswordProtectionSet().
     */
    bool isSheetProtected() const;
    /**
     * @brief returns whether the sheet is protected with password.
     * @return true if both SheetProtection::algorithmName and SheetProtection::hashValue
     * parameters are set in #sheetProtection(), false otherwise.
     *
     * If password protection is set, then #isSheetProtected() also returns true.
     */
    bool isPasswordProtectionSet() const;
    /**
     * @brief sets the password protection to the sheet.
     * @param algorithm a string that describes the hashing algorithm used.
     * See #SheetProtection::algorithmName for some reserved values.
     * @param hash a string that contains the hashed password in a base64 form.
     * @param salt a string that contains the salt in a base64 form.
     * @param spinCount count of iterations to compute the password hash.
     *
     * The actual hashing should be done outside this library.
     * See QCryptographicHash and QPasswordDigestor.
     */
    void setPassword(const QString &algorithm, const QString &hash, const QString &salt = QString(), int spinCount = 1);
    /**
     * @brief sets the default sheet protection parameters.
     *
     * This method sets the default protection parameters to the sheet.
     * @see SheetProtection class description for the default values.
     * @note This method deletes the previous sheet protection parameters.
     */
    void setDefaultSheetProtection();
    /**
     * @brief deletes any sheet protection parameters that has been set, including the
     * default ones.
     */
    void removeSheetProtection();

    /**
     * @brief returns whether the sheet's tab is selected.
     * @return true if the sheet's tab is selected.
     *
     * The default value is false.
     */
    bool isSelected() const;
    /**
     * @brief sets @a selected to the sheet's tab.
     * @param selected tab selection state.
     *
     * If not set, the default value is false (not selected).
     */
    void setSelected(bool selected);
    /**
     * @brief returns window zoom magnification for last added view as a percent value.
     *
     * This parameter is restricted to values ranging from 10 to 400.
     *
     * @return window zoom magnification if it was set, 100 being the default value.
     * @note To get the view scales for specific view types see SheetView::zoomScaleNormal,
     * SheetView::zoomScalePageBreakView, SheetView::zoomScalePageLayoutView.
     */
    int viewZoomScale() const;
    /**
     * @brief sets window zoom magnification for last added view as a percent value.
     * @param scale value ranging from 10 to 400.
     *
     * If not set, the default value is 100.
     * @note To set the view scales for specific view types see SheetView::zoomScaleNormal,
     * SheetView::zoomScalePageBreakView, SheetView::zoomScalePageLayoutView.
     */
    void setViewZoomScale(int scale);

    int workbookViewId() const;//TODO: doc
    void setWorkbookViewId(int id);//TODO: doc
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
     * throws std::out_of_range exception.
     * @return reference to the sheet view with the (zero-based) @a index.
     * @note The newly inserted view will point to the workbook view with
     * id 0. See SheetView::workbookViewId.
     *
     */
    SheetView &view(int index);
    /**
     * @brief returns the count of sheet views defined in the worksheet.
     * @return
     */
    int viewsCount() const;
    /**
     * @brief adds new default-constructed sheet view.
     * @return reference to the added view.
     */
    SheetView & addView();
    /**
     * @brief removes the sheet view with @a index.
     * @param index non-negative index of the view.
     * @return true if the view was found and removed, false otherwise.
     */
    bool removeView(int index);

protected:
    friend class Workbook;
    AbstractSheet(const QString &sheetName, int sheetId, Workbook *book, AbstractSheetPrivate *d);
    /**
     * @brief Copies the current sheet to a sheet called @a distName with @a distId.
     * Returns the new sheet.
     */
    virtual AbstractSheet *copy(const QString &distName, int distId) const = 0;
    void setName(const QString &sheetName);
    void setType(Type type);
    int id() const;

    Drawing *drawing() const;
};

}
#endif // XLSXABSTRACTSHEET_H
