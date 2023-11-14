#ifndef XLSXPAGESETUP_H
#define XLSXPAGESETUP_H

#include "xlsxglobal.h"

#include <QXmlStreamReader>
#include <QSharedData>
#include <QDebug>

namespace QXlsx {

class PageSetupPrivate;

/**
 * @brief The PageSetup class represents paper and printing parameters when printing the sheet.
 *
 * Not all parameters are applicable to chartsheets. See the description of parameters.
 * If a non-applicable parameter is set for a chartsheet, it will be ignored
 * when saving the document.
 *
 * The class is _implicitly shareable_.
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
     * @brief Creates an invalid (default) PageSetup.
     */
    PageSetup();
    /**
     * @brief Creates a copy of @a other.
     */
    PageSetup(const PageSetup &other);
    PageSetup &operator=(const PageSetup &other);
    ~PageSetup();
    bool operator==(const PageSetup &other) const;
    bool operator!=(const PageSetup &other) const;
    operator QVariant() const;

    bool isValid() const;

    /**
     * @brief returns paper size.
     *
     * If both #paperWidth() and #paperHeight() are specified, #paperSize() is ignored.
     *
     * If no value is specified, then PaperSize::Letter is assumed.
     * @sa AbstractSheet::setPaperSize(), AbstractSheet::paperSize().
     */
    std::optional<PaperSize> paperSize() const;
    /**
     * @brief sets paper size.
     *
     * If both #paperWidth() and #paperHeight() are specified, #paperSize() is ignored.
     *
     * If no value is specified, then PaperSize::Letter is assumed.
     * @param paperSize PaperSize enum value.
     * @sa AbstractSheet::setPaperSize(), AbstractSheet::paperSize().
     */
    void setPaperSize(PaperSize paperSize);
    /**
     * @brief returns paper width as a universal measure (f.e. "210mm").
     *
     * If both #paperWidth() and #paperHeight() are specified (not empty), #paperSize()
     * is ignored.
     * @sa AbstractSheet::setPaperSize(), AbstractSheet::setPaperSizeMM(),
     * AbstractSheet::setPaperSizeInches(), AbstractSheet::paperWidth(),
     * AbstractSheet::paperHeight().
     * @return string in the format "-?[0-9]+(\.[0-9]+)?(mm|cm|in|pt|pc|pi)".
     */
    QString paperWidth() const;
    /**
     * @brief sets paper width as a universal measure (f.e. "210mm").
     *
     * If both #paperWidth() and #paperHeight() are specified (not empty),
     * #paperSize() is ignored.
     * @sa AbstractSheet::setPaperSize(), AbstractSheet::setPaperSizeMM(), AbstractSheet::setPaperSizeInches(),
     * AbstractSheet::paperWidth(), AbstractSheet::paperHeight()
     * @param paperWidth string in the format "-?[0-9]+(\.[0-9]+)?(mm|cm|in|pt|pc|pi)".
     */
    void setPaperWidth(const QString &paperWidth);
    /**
     * @brief returns paper height as a universal measure (f.e. "210mm").
     *
     * If both #paperWidth() and #paperHeight() are specified (not empty),
     * #paperSize() is ignored.
     * @sa AbstractSheet::setPaperSize(), AbstractSheet::setPaperSizeMM(), AbstractSheet::setPaperSizeInches(),
     * AbstractSheet::paperWidth(), AbstractSheet::paperHeight().
     * @return string in the format "-?[0-9]+(\.[0-9]+)?(mm|cm|in|pt|pc|pi)".
     */
    QString paperHeight() const;
    /**
     * @brief sets paper height as a universal measure (f.e. "210mm").
     *
     * If both #paperWidth() and #paperHeight() are specified (not empty),
     * #paperSize() is ignored.
     * @sa AbstractSheet::setPaperSize(), AbstractSheet::setPaperSizeMM(), AbstractSheet::setPaperSizeInches(),
     * AbstractSheet::paperWidth(), AbstractSheet::paperHeight().
     * @param paperHeight string in the format "-?[0-9]+(\.[0-9]+)?(mm|cm|in|pt|pc|pi)".
     */
    void setPaperHeight(const QString &paperHeight);
    /**
     * @brief returns the print scale in percents.
     *
     * The value of 100 equals 100% scaling. If no value is specified, then 100 is assumed.
     *
     * If #fitToWidth() and/or #fitToHeight() are specified, #scale() is ignored.
     * @note This parameter is not applicable to chartsheets.
     * @sa Worksheet::setPrintScale(), Worksheet::printScale().
     */
    std::optional<int> scale() const;
    /**
     * @brief sets the print scale in percents.
     * The value of 100 equals 100% scaling. If no value is specified, then 100 is assumed.
     * @param scale Integer value from 10 to 400.
     *
     * If #fitToWidth() and/or #fitToHeight() are specified, #scale() is ignored.
     * @note This parameter is not applicable to chartsheets.
     * @sa Worksheet::setPrintScale(), Worksheet::printScale().
     */
    void setScale(int scale);
    /**
     * @brief returns page number for first printed page.
     *
     * If no value is specified, then 1 is assumed.
     * @sa AbstractSheet::setFirstPageNumber(), AbstractSheet::firstPageNumber().
     */
    std::optional<int> firstPageNumber() const;
    /**
     * @brief sets page number for first printed page.
     *
     * If no value is specified, then 1 is assumed.
     *
     * @note Use #setUseFirstPageNumber() to actually _switch on_ separate first page number.
     * @sa #setUseFirstPageNumber(), AbstractSheet::setFirstPageNumber(),
     * AbstractSheet::firstPageNumber().
     * @param firstPageNumber positive integer.
     */
    void setFirstPageNumber(int firstPageNumber);
    /**
     * @brief returns the number of horizontal pages to fit on.
     *
     * If no value is specified, then 1 is assumed.
     *
     * If #fitToWidth() and/or #fitToHeight() are specified, #scale() is ignored.
     * @note This parameter is not applicable to chartsheets.
     * @sa AbstractSheet::setFitToWidth(), AbstractSheet::fitToWidth()
     */
    std::optional<int> fitToWidth() const;
    /**
     * @brief sets the number of horizontal pages to fit on.
     *
     * If no value is specified, then 1 is assumed.
     *
     * If #fitToWidth() and/or #fitToHeight() are specified, #scale() is ignored.
     * @note This parameter is not applicable to chartsheets.
     * @sa AbstractSheet::setFitToWidth(), AbstractSheet::fitToWidth()
     * @param pages positive integer.
     */
    void setFitToWidth(int pages);
    /**
     * @brief returns the number of vertical pages to fit on.
     *
     * If no value is specified, then 1 is assumed.
     *
     * If #fitToWidth() and/or #fitToHeight() are specified, #scale() is ignored.
     * @note This parameter is not applicable to chartsheets.
     * @sa AbstractSheet::setFitToHeight(), AbstractSheet::fitToHeight()
     */
    std::optional<int> fitToHeight() const;
    /**
     * @brief sets the number of vertical pages to fit on.
     *
     * If no value is specified, then 1 is assumed.
     *
     * If #fitToWidth() and/or #fitToHeight() are specified, #scale() is ignored.
     * @note This parameter is not applicable to chartsheets.
     * @sa AbstractSheet::setFitToHeight(), AbstractSheet::fitToHeight()
     * @param pages positive integer.
     */
    void setFitToHeight(int pages);
    /**
     * @brief returns the order of printed pages.
     *
     * If no value is specified, then PageOrder::DownThenOver is assumed.
     * @note This parameter is not applicable to chartsheets.
     * @sa Worksheet::setPageOrder(), Worksheet::pageOrder().
     */
    std::optional<PageOrder> pageOrder() const;
    /**
     * @brief sets the order of printed pages.
     * @param pageOrder PageOrder enum value. If no value is specified, then PageOrder::DownThenOver is assumed.
     * @note This parameter is not applicable to chartsheets.
     * @sa Worksheet::setPageOrder(), Worksheet::pageOrder().
     */
    void setPageOrder(PageOrder pageOrder);
    /**
     * @brief returns the page orientation.
     *
     * If no value is specified, then Orientation::Default is assumed.
     * @sa AbstractSheet::setPageOrientation(), AbstractSheet::pageOrientation().
     */
    std::optional<Orientation> orientation() const;
    /**
     * @brief sets the page orientation.
     *
     * If no value is specified, then Orientation::Default is assumed.
     * @sa AbstractSheet::setPageOrientation(), AbstractSheet::pageOrientation().
     * @param orientation Orientation enum value.
     */
    void setOrientation(Orientation orientation);
    /**
     * @brief returns whether to use the printer default parameters.
     *
     * Use the printer’s defaults settings for page setup values and don't use the
     * default values of the file. Example: if dpi is not present or specified in
     * the file, the application must not assume 600dpi as specified as a default
     * and instead must let the printer specify the default dpi.
     *
     * If no value is specified, then true is assumed.
     */
    std::optional<bool> usePrinterDefaults() const;
    /**
     * @brief sets whether to use the printer default parameters.
     *
     * Use the printer’s defaults settings for page setup values and don't use the
     * default values of the file. Example: if dpi is not present or specified in
     * the file, the application must not assume 600dpi as specified as a default
     * and instead must let the printer specify the default dpi.
     *
     * If no value is specified, then true is assumed.
     * @param use If true, then printer defaults are used instead of PageSetup defaults.
     */
    void setUsePrinterDefaults(bool use);
    /**
     * @brief returns whether to print in black and white.
     *
     * If no value is specified, then false is assumed.
     */
    std::optional<bool> printBlackAndWhite() const;
    /**
     * @brief sets whether to print in black and white.
     * @param blackAndWhite If true, then the page will be printed in black and white.
     *
     * If no value is specified, then false is assumed.
     */
    void setPrintBlackAndWhite(bool blackAndWhite);
    /**
     * @brief returns whether to print the page as a draft (i.e. without graphics).
     *
     * If no value is specified, then false is assumed.
     */
    std::optional<bool> printDraft() const;
    /**
     * @brief sets whether to print the page as a draft (i.e. without graphics).
     * @param draft If true then no graphics is printed.
     *
     * If no value is specified, then false is assumed.
     */
    void setPrintDraft(bool draft);
    /**
     * @brief returns whether to use #firstPageNumber() value for first page number,
     * and do not auto number the pages.
     *
     * If no value is specified, then false is assumed.
     * @sa #firstPageNumber(), #setFirstPageNumber().
     */
    std::optional<bool> useFirstPageNumber() const;
    /**
     * @brief sets whether to use #firstPageNumber() value for first page number,
     * and do not auto number the pages.
     * @param use If true, then the value of #firstPageNumber() will be used
     * for first page number.
     *
     * If no value is specified, then false is assumed.
     */
    void setUseFirstPageNumber(bool use);
    /**
     * @brief returns horizontal print resolution of the device.
     *
     * If no value is specified, then 600 is assumed.
     */
    std::optional<int> horizontalDpi() const;
    /**
     * @brief sets horizontal print resolution of the device.
     * @param dpi Resolution in DPI (dots per inch). If no value is specified,
     * then 600 is assumed.
     */
    void setHorizontalDpi(int dpi);
    /**
     * @brief returns vertical print resolution of the device.
     *
     * If no value is specified, then 600 is assumed.
     */
    std::optional<int> verticalDpi() const;
    /**
     * @brief sets vertical print resolution of the device.
     * @param dpi Resolution in DPI (dots per inch). If no value is specified,
     * then 600 is assumed.
     */
    void setVerticalDpi(int dpi);
    /**
     * @brief returns how many copies to print.
     *
     * If no value is specified, then 1 is assumed.
     */
    std::optional<int> copies() const;
    /**
     * @brief sets how many copies to print.
     *
     * If no value is specified, then 1 is assumed.
     * @param copies
     */
    void setCopies(int copies);
    /**
     * @brief Relationship Id of the devMode printer settings part.
     */
    QString printerId() const; //TODO: adding printer
    /**
     * @brief returns how to display cells with errors when printing the worksheet.
     *
     * If no value is specified, then PrintError::Displayed is assumed.
     * @note This parameter is not applicable to chartsheets.
     */
    std::optional<PrintError> printErrors() const;
    /**
     * @brief sets how to display cells with errors when printing the worksheet.
     * @note This parameter is not applicable to chartsheets.
     * @param errors PrintError enum value. If no value is specified, then PrintError::Displayed is assumed.
     */
    void setPrintErrors(PrintError errors);
    /**
     * @brief returns how cell comments shall be printed.
     *
     * If no value is specified, then CellComments::DoNotPrint is assumed.
     */
    std::optional<CellComments> printCellComments() const;
    /**
     * @brief sets how cell comments shall be printed.
     * @param cellComments CellComments enum value. If no value is specified,
     * then CellComments::DoNotPrint is assumed.
     */
    void setPrintCellComments(CellComments cellComments);
    /**
     * @brief returns whether to print grid lines. The default value is false.
     * @note This parameter is not applicable to chartsheets.
     */
    std::optional<bool> printGridLines() const;
    /**
     * @brief sets whether to print grid lines.
     * @note This parameter is not applicable to chartsheets.
     * @param printGridLines If true, then cells grid lines are printed. The default value is false.
     */
    void setPrintGridLines(bool printGridLines);
    /**
     * @brief returns whether to print row and column headings. The default value is false.
     * @note This parameter is not applicable to chartsheets.
     */
    std::optional<bool> printHeadings() const;
    /**
     * @brief sets whether to print row and column headings.
     * @note This parameter is not applicable to chartsheets.
     * @param printHeadings If true, then row and column headings are printed.
     * The default value is false.
     */
    void setPrintHeadings(bool printHeadings);
    /**
     * @brief returns whether to center data horizontally when printing. The default value is false.
     * @note This parameter is not applicable to chartsheets.
     */
    std::optional<bool> printHorizontalCentered() const;
    /**
     * @brief sets whether to center data horizontally when printing. The default value is false.
     * @note This parameter is not applicable to chartsheets.
     * @param centered
     */
    void setPrintHorizontalCentered(bool centered);
    /**
     * @brief returns whether to center data vertically when printing. The default value is false.
     * @note This parameter is not applicable to chartsheets.
     */
    std::optional<bool> printVerticalCentered() const;
    /**
     * @brief sets whether to center data vertically when printing. The default value is false.
     * @note This parameter is not applicable to chartsheets.
     * @param centered
     */
    void setPrintVerticalCentered(bool centered);
private:
    QSharedDataPointer<PageSetupPrivate> d;

private:
    friend class Worksheet;
    friend class Chartsheet;
    friend class Chart;
    void readPaperSize(QXmlStreamReader &reader);
    void writeWorksheet(QXmlStreamWriter &writer, const QString &name) const;
    void writeChartsheet(QXmlStreamWriter &writer) const;
    void read(QXmlStreamReader &reader);
#ifndef QT_NO_DEBUG_STREAM
    friend QDebug operator<<(QDebug, const PageSetup &f);
#endif
};

}

#ifndef QT_NO_DEBUG_STREAM
QDebug operator<<(QDebug, const QXlsx::PageSetup &f);
#endif

Q_DECLARE_METATYPE(QXlsx::PageSetup)
Q_DECLARE_TYPEINFO(QXlsx::PageSetup, Q_MOVABLE_TYPE);

#endif // XLSXPAGESETUP_H
