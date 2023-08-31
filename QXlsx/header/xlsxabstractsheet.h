// xlsxabstractsheet.h

#ifndef XLSXABSTRACTSHEET_H
#define XLSXABSTRACTSHEET_H

#include "xlsxglobal.h"
#include "xlsxabstractooxmlfile.h"

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
class HeaderFooter
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
     * When true, as left/right*  margins grow and shrink, the header and footer
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
 * Margins are specified in inches. There are methods to set them in millimeters or in inches.
 * @note If you set any of the margins, you need to also set all the other margins,
 * otherwise the document will be ill-formed. Any missing margins will be written as zeroes.
 */
class PageMargins
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
     * @brief creates page margins from values specified in inches.
     * @param left left page margin in inches.
     * @param top top page margin in inches.
     * @param right right page margin in inches.
     * @param bottombottom page margin in inches.
     * @param header header page margin in inches.
     * @param footer footer page margin in inches.
     * @note To set pageMargins in millimetres use #setMarginsMm(), #setLeftMarginMm() etc.
     */
    PageMargins(double left = 0, double top = 0, double right = 0, double bottom = 0,
                double header = 0, double footer = 0);

    void setMarginsInches(double left = 0, double top = 0, double right = 0, double bottom = 0,
                          double header = 0, double footer = 0);
    void setMarginInches(Position pos, double value);
    void setLeftMarginInches(double value);
    void setRightMarginInches(double value);
    void setTopMarginInches(double value);
    void setBottomMarginInches(double value);
    void setHeaderMarginInches(double value);
    void setFooterMarginInches(double value);
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
     * @brief returns pointer to this sheet's workbook.
     * @return
     */
    Workbook *workbook() const;

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

public:
    /**
     * @brief returns the sheet's name.
     * @return
     */
    QString name() const;
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

    QMap<PageMargins::Position, double> pageMarginsInches() const;
    QMap<PageMargins::Position, double> pageMarginsMm() const;
    void setPageMarginsInches(double left = 0, double top = 0, double right = 0, double bottom = 0,
                              double header = 0, double footer = 0);
    void setPageMarginsMm(double left = 0, double top = 0, double right = 0, double bottom = 0,
                          double header = 0, double footer = 0);
    PageMargins pageMargins() const;
    PageMargins &pageMargins();
    void setPageMargins(const PageMargins &margins);

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
