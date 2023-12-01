#ifndef XLSXHEADERFOOTER_H
#define XLSXHEADERFOOTER_H

#include "xlsxglobal.h"

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
}

#endif // XLSXHEADERFOOTER_H
