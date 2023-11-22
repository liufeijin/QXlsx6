#ifndef XLSXPAGEMARGINS_H
#define XLSXPAGEMARGINS_H

#include "xlsxglobal.h"

#include <QXmlStreamReader>

namespace QXlsx {

/**
 * @brief The PageMargins class represents the margins of a sheet page.
 * Margins are specified in inches or millimetres.
 * @note If you set any of the margins, you need to also set all the other
 * margins, otherwise the document will be ill-formed. Any missing margins will
 * be written as zeroes.
 */
class QXLSX_EXPORT PageMargins
{
public:
    /**
     * @brief The Position enum specifies the page margin position
     */
    enum class Position
    {
        Top, /**< Top margin */
        Bottom, /**< Bottom margin */
        Left, /**< Left margin */
        Right, /**< Right margin */
        Header, /**< Header margin */
        Footer /**< Footer margin */
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
     * @return `true` if at least one margin is set, `false` otherwise.
     */
    bool isValid() const;
    void write(QXmlStreamWriter &writer) const;
    void read(QXmlStreamReader &reader);
private:
    QMap<Position, double> mMargins;
};

}

#endif // XLSXPAGEMARGINS_H
