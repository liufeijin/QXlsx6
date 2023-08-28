#ifndef XLSXTEXT_H
#define XLSXTEXT_H

#include "xlsxglobal.h"
#include "xlsxformat.h"
#include <QVariant>
#include <QStringList>
#include <QSharedDataPointer>
#include <QXmlStreamReader>
#include <QXmlStreamWriter>
#include <QTextBlock>

#include "xlsxmain.h"
#include "xlsxfont.h"
#include "xlsxshapeformat.h"


namespace QXlsx {

class Text;

/**
 * @brief The TextProperties class represents the text body layout.
 */
class QXLSX_EXPORT TextProperties
{
public:
    /**
     * @brief The OverflowType enum specifies the text overflow.
     */
    enum class OverflowType
    {
        Overflow, /**< allows overflow of text */
        Ellipsis, /**< use an ellipsis to indicate there is text which is not visible */
        Clip /**< do not indicate there is text which is not visible */
    };
    /**
     * @brief If there is vertical text, determines what kind of vertical text is going to be used.
     */
    enum class VerticalType
    {
        EastAsianVertical, /**< A special version of vertical text, where some fonts are
displayed as if rotated by 90 degrees while some fonts (mostly East Asian) are displayed vertical.*/
        Horizontal, /**< Horizontal text. */
        MongolianVertical, /**< A special version of vertical text, where some fonts are
displayed as if rotated by 90 degrees while some fonts (mostly East Asian) are displayed vertical. The
difference between this and the eastAsianVertical is the text flows top down then LEFT RIGHT, instead of
RIGHT LEFT*/
        Vertical, /**< All of the text is vertical orientation (each line is 90 degrees rotated clockwise,
so it goes from top to bottom; each next line is to the left from the previous one).*/
        Vertical270, /**< All of the text is vertical orientation (each line is 270 degrees rotated clockwise, so it goes
from bottom to top; each next line is to the right from the previous one).*/
        WordArtVertical, /**< All of the text is vertical ("one letter on top of another"). */
        WordArtVerticalRtl, /**< Vertical WordArt should be shown from right to left rather than left to right.*/
    };
    /**
     * @brief The Anchor enum specifies a list of available vertical anchoring types for text
     */
    enum class Anchor
    {
        Bottom, /**< Anchor the text at the bottom of the bounding rectangle. */
        Center, /**< Anchor the text at the middle of the bounding rectangle. */
        Distributed, /**< Distribute the lines of text vertically. If there is only one line,
                       anchor it at middle. */
        Justified, /**< Justify the lines of text vertically. If there is only one line,
                     anchor it at top.*/
        Top /**< Anchor the text at the top of the bounding rectangle. */
    };
    /**
     * @brief The TextAutofit enum specifies the autofit (scaling) type of the text body
     * in order to remain inside the text box.
     */
    enum class TextAutofit
    {
        NoAutofit, /**< The text within a text box is not scaled. This is the default value.*/
        NormalAutofit, /**<  The text within the text body should be scaled to the bounding box.
The scaling parameters are fontScale and lineSpaceReduction.*/
        ShapeAutofit /**< The shape that contains the text box should be scaled in order to
fully contain the text described within it.*/
    };

    /**
     * Specifies whether the before and after paragraph spacing defined by the user is to be respected.
     */
    std::optional<bool> spcFirstLastPara;
    /**
     * @brief Determines whether the text can flow out of the bounding box vertically.
     * This is used to determine what happens in the event that the text within a
     * shape is too large for the bounding box it is contained within. If this
     * attribute is omitted, then a value of OverflowType::Overflow is implied.
     */
    std::optional<OverflowType> verticalOverflow;
    /**
     * @brief Determines whether the text can flow out of the bounding box horizontally.
     * This is used to determine what happens in the event that the text within a
     * shape is too large for the bounding box it is contained within. If this
     * attribute is omitted, then a value of OverflowType::Overflow is implied.
     */
    std::optional<OverflowType> horizontalOverflow;
    /**
     * @brief Determines if the text within the given text body should be displayed vertically.
     * If this attribute is omitted, then a value of VerticalType::Horizontal is implied.
     */
    std::optional<VerticalType> verticalOrientation;
    /**
     * @brief  Specifies the rotation that is being applied to the text within the bounding box.
     * If it not specified, the rotation of the accompanying shape is used. If it is specified,
     * then this is applied independently from the shape.
     */
    Angle rotation;
    /**
     * @brief Specifies the wrapping options to be used for this text body. If it is specified,
     * then text is wrapped using the bounding text box.
     */
    std::optional<bool> wrap;
    /**
     * @brief Specifies the left inset of the text bounding rectangle. Insets are used just as internal
     * margins for text boxes within shapes. If leftInset is not valid, a value of 0.1 inches is implied.
     */
    Coordinate leftInset;
    /**
     * @brief Specifies the right inset of the text bounding rectangle. Insets are used just as internal
     * margins for text boxes within shapes. If rightInset is not valid, a value of 0.1 inches is implied.
     */
    Coordinate rightInset;
    /**
     * @brief Specifies the top inset of the text bounding rectangle. Insets are used just as internal
     * margins for text boxes within shapes. If topInset is not valid, a value of 0.05 inches is implied.
     */
    Coordinate topInset;
    /**
     * @brief Specifies the bottom inset of the text bounding rectangle. Insets are used just as internal
     * margins for text boxes within shapes. If bottomInset is not valid, a value of 0.05 inches is implied.
     */
    Coordinate bottomInset;
    /**
     * @brief Specifies the number of columns of text in the bounding rectangle.
     */
    std::optional<int> columnCount;
    /**
     * @brief Specifies the space between text columns in the text area when there
     * is more than 1 column present.
     */
    Coordinate columnSpace;
    /**
     * @brief Specifies whether columns are used in a right-to-left or left-to-right order.
     */
    std::optional<bool> columnsRtl;
    /**
     * @brief Specifies that text within this textbox is converted text from a WordArt object.
     */
    std::optional<bool> fromWordArt;
    /**
     * @brief  Specifies the anchoring position of the text within the shape.
     * If this attribute is omitted, then top anchoring is implied.
     */
    std::optional<Anchor> anchor;
    /**
     * @brief Specifies whether the smallest possible "bounds box" for the text
     * should be determined to center this box accordingly. The default value is false.
     */
    std::optional<bool> anchorCentering;
    /**
     * @brief Forces the text to be rendered anti-aliased regardless of the font size.
     * The default value is false.
     */
    std::optional<bool> forceAntiAlias;
    /**
     * @brief  Specifies whether text should remain upright, regardless of the transform applied to it
     * and the accompanying shape transform.
     */
    std::optional<bool> upright;
    /**
     * @brief Specifies that the line spacing for this text body is decided in a simplistic manner using
     * the font scene. The default value is false.
     */
    std::optional<bool> compatibleLineSpacing;
    /**
     * @brief  Specifies a preset geometric shape that should be used to transform (warp) a piece of text.
     */
    std::optional<PresetTextShape> textShape;

    /**
     * @brief specifies the text body scaling in order to remain inside the text box.
     */
    std::optional<TextAutofit> textAutofit;
    /**
     * @brief Specifies the percentage of the original font size to which each fragment of text
     * is scaled.
     * A value of 100.0 scales the text to 100%, while a value of 1.0 scales
     * the text to 1%. If it is not set, then a value of 100% is implied.
     * This member is ignored if textAutofit is not TextAutofit::NormalAutofit.
     */
    std::optional<double> fontScale;
    /**
     * @brief Specifies the percentage amount by which the line spacing of each paragraph
     * in the text body is reduced.
     * The reduction is applied by subtracting it from the original line spacing value.
     * A value of 100.0 reduces the line spacing by 100%, while a value of 1.0
     * reduces the line spacing by one percent. The default value is 0.0.
     */
    std::optional<double> lineSpaceReduction;

    std::optional<Scene3D> scene3D;
    // either text3D or z
    std::optional<Shape3D> text3D;
    Coordinate z;

    bool isValid() const;

    bool operator ==(const TextProperties &other) const;
    bool operator !=(const TextProperties &other) const;

    void read(QXmlStreamReader &reader);
    void write(QXmlStreamWriter &writer) const;
};

/**
 * @brief The CharacterProperties class represents the text character properties.
 */
class QXLSX_EXPORT CharacterProperties
{
public:
    enum class UnderlineType
    {
        None,
        Words,
        Single,
        Double,
        Heavy,
        Dotted,
        DottedHeavy,
        Dash,
        DashHeavy,
        DashLong,
        DashLongHeavy,
        DotDash,
        DotDashHeavy,
        DotDotDash,
        DotDotDashHeavy,
        Wavy,
        WavyHeavy,
        WavyDouble,
    };
    enum class StrikeType
    {
        None,
        Single,
        Double
    };
    enum class CapitalizationType
    {
        None,
        AllCaps,
        SmallCaps
    };

    std::optional<bool> kumimoji; /**< @brief Specifies whether the numbers contained within
                                       vertical text continue vertically with the
                                       text or whether they are to be displayed horizontally
                                       while the surrounding characters
                                       continue in a vertical fashion. */
    /**
     *  @brief Specifies the language to be used when the generating application is displaying the user
        interface controls.
     */
    QString language; // f.e. "en-US"
    /**
     * @brief Specifies the alternate language to use when the generating application is displaying the
     * user interface controls.
     *
     * If this attribute is omitted, then language is used.
     */
    QString alternateLanguage; // f.e. "en-US"
    std::optional<double> fontSize; /**<  @brief Specifies the size, in points, of text within a text fragment.*/
    std::optional<bool> bold; /**< @brief Specifies whether a fragment of text is formatted as bold text. */
    std::optional<bool> italic; /**< @brief Specifies whether a fragment of text is formatted as italic text. */
    std::optional<UnderlineType> underline; /**< @brief  Specifies whether a fragment of text is formatted as underlined text. */
    std::optional<StrikeType> strike; /**< @brief Specifies whether a fragment of text is formatted as strikethrough text. */
    std::optional<double> kerningFontSize; /**< @brief  Specifies the minimum font point size at which
                                                 character kerning occurs for this text fragment.

                                                 If this attribute is omitted, then kerning
                                                 occurs for all font sizes down to a 0 point
                                                 font. */
    std::optional<CapitalizationType> capitalization; /**< @brief  Specifies the capitalization that is to be applied to the text fragment. */
    std::optional<TextPoint> spacing; /**< @brief Specifies the spacing in points between characters within a text fragment. */
    std::optional<bool> normalizeHeights; /**< @brief Specifies the normalization of height that is
                                               to be applied to the text fragment. */
    std::optional<double> baseline;  /**< @brief Specifies the baseline for both the superscript
                                          and subscript fonts.

                                          The size is specified
                                          using a percentage where 1% is equal to
                                          1 percent of the font size and 100% is equal to
                                          100 percent font of the font size. */
    std::optional<bool> noProofing; /**< @brief Specifies that a fragment of text has been selected
                                         by the user to not be checked for mistakes.

                                         Therefore if there are spelling, grammar, etc
                                         mistakes within this text the generating
                                         application should ignore them. */
    std::optional<bool> proofingNeeded; /**< @brief  Specifies that the content of a text fragment
                                              has changed since the proofing tools have last
                                              been run.

                                              Effectively this flags text that
                                              is to be checked again by the generating
                                              application for mistakes such as spelling, grammar, etc. */
    std::optional<bool> checkForSmartTagsNeeded; /**< @brief Specifies whether or not a text fragment has been
                                                      checked for smart tags.

                                                      A value of
                                                      true here indicates to the generating application
                                                      that this text fragment should be checked for
                                                      smart tags. */
    std::optional<bool> spellingErrorFound; /**< @brief  Specifies that when this fragment of
                                                 text was checked for spelling, grammar, etc. that a
                                                 mistake was indeed found.

                                                 This allows
                                                 the generating application to effectively save the
                                                 state of the mistakes within the document instead
                                                 of having to perform a full pass check
                                                 upon opening the document. */
    std::optional<int> smartTagId; /**< @brief Specifies a smart tag identifier for a fragment of text.

                                        This ID is unique throughout the file and is used
                                        to reference corresponding auxiliary information about
                                        the smart tag. */
    QString bookmarkLinkTarget; /**< @brief Specifies the link target name that is used
                                     to reference to the proper link properties in a
                                     custom XML part within the document */
    LineFormat line; /**< @brief Specifies the line format to be applied to this text fragment. */
    FillFormat fill; /**< @brief Specifies the fill format to be applied to this text fragment. */
    Effect effect; /**< @brief Specifies the effects (blur, shadow etc.) to be applied to this text fragment. */
    Color highlightColor; /**< @brief Specifies the color to highlight this text fragment */

    /**
     * @brief Specifies that the stroke style of an underline for a fragment of text should be
     * the same as of the text fragment within which it is contained (that is #line).
     *
     * If this parameter is set, then #textUnderline is ignored.
     */
    std::optional<bool> textUnderlineFollowsText;
    /**
     * @brief Specifies a distinct stroke style of an underline if it should differ from #line.
     *
     * If this parameter is set, then #textUnderlineFollowsText should be nullopt or false.
     */
    LineFormat textUnderline;
    /**
     * @brief Specifies that the fill style of an underline for a fragment of text should be
     * the same as of the text fragment within which it is contained (that is #fill).
     *
     * If this parameter is set, then #textUnderlineFill is ignored.
     */
    std::optional<bool> textUnderlineFillFollowsText;
    /**
     * @brief Specifies a distinct fill style of an underline if it should differ from #fill.
     *
     * If this parameter is set, then #textUnderlineFillFollowsText should be nullopt or false.
     */
    FillFormat textUnderlineFill;
    /**
     * @brief Specifies a Latin Font to be used for a specific fragment of text.
     */
    Font latinFont;
    /**
     * @brief Specifies an East Asian Font to be used for a specific fragment of text.
     */
    Font eastAsianFont;
    /**
     * @brief Specifies a complex script (Arabic, Tamil etc.) Font to be used for a specific fragment of text.
     */
    Font complexScriptFont;
    /**
     * @brief Specifies a symbolic Font to be used for a specific fragment of text.
     */
    Font symbolFont;
    /**
     * @brief Specifies whether the contents of this fragment shall have right-to-left characteristics
     */
    std::optional<bool> rightToLeft;

    //TODO: hyperlinks
//          <xsd:element name="hlinkClick" type="CT_Hyperlink" minOccurs="0" maxOccurs="1"/>
//          <xsd:element name="hlinkMouseOver" type="CT_Hyperlink" minOccurs="0" maxOccurs="1"/>

    bool operator ==(const CharacterProperties &other) const;
    bool operator !=(const CharacterProperties &other) const;

    bool isValid() const;

    void read(QXmlStreamReader &reader);
    void write(QXmlStreamWriter &writer, const QString &name) const;
    static CharacterProperties from(const QTextCharFormat &format);
private:
    SERIALIZE_ENUM(UnderlineType, {
        {UnderlineType::None, "none"},
        {UnderlineType::Words, "words"},
        {UnderlineType::Single, "sng"},
        {UnderlineType::Double, "dbl"},
        {UnderlineType::Heavy, "heavy"},
        {UnderlineType::Dotted, "dotted"},
        {UnderlineType::DottedHeavy, "dottedHeavy"},
        {UnderlineType::Dash, "dash"},
        {UnderlineType::DashHeavy, "dashHeavy"},
        {UnderlineType::DashLong, "dashLong"},
        {UnderlineType::DashLongHeavy, "dashLongHeavy"},
        {UnderlineType::DotDash, "dotDash"},
        {UnderlineType::DotDashHeavy, "dotDashHeavy"},
        {UnderlineType::DotDotDash, "dotDotDash"},
        {UnderlineType::DotDotDashHeavy, "dotDotDashHeavy"},
        {UnderlineType::Wavy, "wavy"},
        {UnderlineType::WavyHeavy, "wavyHeavy"},
        {UnderlineType::WavyDouble, "wavyDbl"},
    });
    SERIALIZE_ENUM(StrikeType, {
        {StrikeType::None, "noStrike"},
        {StrikeType::Single, "sngStrike"},
        {StrikeType::Double, "dblStrike"},
    });
    SERIALIZE_ENUM(CapitalizationType, {
        {CapitalizationType::None,"none"},
        {CapitalizationType::AllCaps, "all"},
        {CapitalizationType::SmallCaps, "small"},
    });
};

/**
 * @brief The Size class
 * The class is used to set a size either in points or in percents.
 * Size can either be set as a whole positive number (int) or as a percentage (double).
 */
class QXLSX_EXPORT Size
{
public:
    Size();
    explicit Size(uint points); // [0..158400]
    explicit Size(double percents); // 34.0 = 34.0%

    int toPoints() const;
    double toPercents() const;
    QString toString() const;

    void setPoints(int points);
    void setPercents(double percents);

    void read(QXmlStreamReader &reader);
    void write(QXmlStreamWriter &writer, const QString &name) const;

    bool isValid() const;
    bool isPoints() const;
    bool isPercents() const;

    bool operator==(const Size &other) const {return val == other.val;}
    bool operator!=(const Size &other) const {return val != other.val;}
private:
    QVariant val;
};

/**
 * @brief The ParagraphProperties class
 * The class is used with the Paragraph object to set its properties: spacing, indent, alignment etc.
 */
class QXLSX_EXPORT ParagraphProperties
{
//    <xsd:complexType name="CT_TextParagraphProperties">
//      <xsd:sequence>
//        <xsd:element name="lnSpc" type="CT_TextSpacing" minOccurs="0" maxOccurs="1"/>
//        <xsd:element name="spcBef" type="CT_TextSpacing" minOccurs="0" maxOccurs="1"/>
//        <xsd:element name="spcAft" type="CT_TextSpacing" minOccurs="0" maxOccurs="1"/>
//        <xsd:group ref="EG_TextBulletColor" minOccurs="0" maxOccurs="1"/>
//        <xsd:group ref="EG_TextBulletSize" minOccurs="0" maxOccurs="1"/>
//        <xsd:group ref="EG_TextBulletTypeface" minOccurs="0" maxOccurs="1"/>
//        <xsd:group ref="EG_TextBullet" minOccurs="0" maxOccurs="1"/>
//        <xsd:element name="tabLst" type="CT_TextTabStopList" minOccurs="0" maxOccurs="1"/>
//        <xsd:element name="defRPr" type="CT_TextCharacterProperties" minOccurs="0" maxOccurs="1"/>
//        <xsd:element name="extLst" type="CT_OfficeArtExtensionList" minOccurs="0" maxOccurs="1"/>
//      </xsd:sequence>
//      <xsd:attribute name="marL" type="ST_TextMargin" use="optional"/>
//      <xsd:attribute name="marR" type="ST_TextMargin" use="optional"/>
//      <xsd:attribute name="lvl" type="ST_TextIndentLevelType" use="optional"/>
//      <xsd:attribute name="indent" type="ST_TextIndent" use="optional"/>
//      <xsd:attribute name="algn" type="ST_TextAlignType" use="optional"/>
//      <xsd:attribute name="defTabSz" type="ST_Coordinate32" use="optional"/>
//      <xsd:attribute name="rtl" type="xsd:boolean" use="optional"/>
//      <xsd:attribute name="eaLnBrk" type="xsd:boolean" use="optional"/>
//      <xsd:attribute name="fontAlgn" type="ST_TextFontAlignType" use="optional"/>
//      <xsd:attribute name="latinLnBrk" type="xsd:boolean" use="optional"/>
//      <xsd:attribute name="hangingPunct" type="xsd:boolean" use="optional"/>
//    </xsd:complexType>
public:
    enum class TextAlign
    {
        Left,
        Center,
        Right,
        Justify,
        JustifyLow,
        Distrubute,
        DistrubuteThai,
    };

    enum class FontAlign
    {
        Auto,
        Top,
        Center,
        Base,
        Bottom
    };

    enum class BulletType
    {
        None,
        AutoNumber,
        Char,
        Blip
    };

    enum class BulletAutonumberType
    {
        AlphaLowcaseParentheses, // (a), (b), (c)
        AlphaUppercaseParentheses, // (A), (B), (C)
        AlphaLowcaseRightParentheses, // a), b), c)
        AlphaUppercaseRightParentheses, // A), B), C)
        AlphaLowcasePeriod, // a., b., c.
        AlphaUppercasePeriod, // A., B., C.
        ArabicParentheses, // (1), (2), (3)
        ArabicRightParentheses, // 1), 2), 3)
        ArabicPeriod, // 1., 2., 3.
        ArabicPlain, // 1, 2, 3
        RomanLowcaseParentheses, // (i), (ii), (iii)
        RomanUppercaseParentheses, // (I), (II), (III)
        RomanLowcaseRightParentheses, // i), ii), iii)
        RomanUppercaseRightParentheses, // I), II), III)
        RomanLowcasePeriod, // i., ii., iii.
        RomanUppercasePeriod, // I., II, III.
        Circle,
        CircleWindingsBlack,
        CircleWindingsWhite,
        ArabicDoubleBytePeriod,
        ArabicDoubleByte,
        SimplifiedChinesePeriod,
        SimplifiedChinese,
        TraditionalChinesePeriod,
        TraditionalChinese,
        JapanesePeriod,
        JapaneseKorean,
        JapaneseKoreanPeriod,
        BidiArabic1Minus,
        BidiArabic2Minus,
        BidiHebrewMinus,
        ThaiAlphaPeriod,
        ThaiAlphaRightParentheses,
        ThaiAlphaParentheses,
        ThaiNumberPeriod,
        ThaiNumberRightParentheses,
        ThaiNumberParentheses,
        HindiAlphaPeriod,
        HindiNumberPeriod,
        HindiNumberParentheses,
        HindiAlpha1Period,
    };

    enum class TabAlign
    {
        Left,
        Center,
        Right,
        Decimal
    };

    Coordinate leftMargin;
    Coordinate rightMargin;
    Coordinate indent;
    std::optional<int> indentLevel; //[0..8]
    std::optional<TextAlign> align;
    Coordinate defaultTabSize;
    std::optional<bool> rtl;
    std::optional<bool> eastAsianLineBreak;
    std::optional<bool> latinLineBreak;
    std::optional<bool> hangingPunctuation;
    std::optional<FontAlign> fontAlign;

    std::optional<Size> lineSpacing;
    std::optional<Size> spacingBefore;
    std::optional<Size> spacingAfter;

    std::optional<Color> bulletColor;
    std::optional<Size> bulletSize;
    std::optional<Font> bulletFont;
    std::optional<BulletType> bulletType;
    BulletAutonumberType bulletAutonumberType; //makes sense only if bulletType == Autonumber
    std::optional<int> bulletAutonumberStart;
    QString bulletChar; //makes sense only if bulletType == Char

    std::optional<CharacterProperties> defaultTextCharacterProperties;

    QList<QPair<Coordinate, TabAlign>> tabStops;

    bool isValid() const;

    bool operator==(const ParagraphProperties &other) const;
    bool operator!=(const ParagraphProperties &other) const;

    void read(QXmlStreamReader &reader);
    void write(QXmlStreamWriter &writer, const QString &name) const;

    static ParagraphProperties from(const QTextBlock &block);
private:
    SERIALIZE_ENUM(
        BulletAutonumberType,
        {
            {BulletAutonumberType::AlphaLowcaseParentheses, "alphaLcParenBoth"},   // (a), (b), (c)
            {BulletAutonumberType::AlphaUppercaseParentheses, "alphaUcParenBoth"}, // (A), (B), (C)
            {BulletAutonumberType::AlphaLowcaseRightParentheses, "alphaLcParenR"}, // a), b), c)
            {BulletAutonumberType::AlphaUppercaseRightParentheses, "alphaUcParenR"}, // A), B), C)
            {BulletAutonumberType::AlphaLowcasePeriod, "alphaLcPeriod"},             // a., b., c.
            {BulletAutonumberType::AlphaUppercasePeriod, "alphaUcPeriod"},           // A., B., C.
            {BulletAutonumberType::ArabicParentheses, "arabicParenBoth"},        // (1), (2), (3)
            {BulletAutonumberType::ArabicRightParentheses, "arabicParenR"},      // 1), 2), 3)
            {BulletAutonumberType::ArabicPeriod, "arabicPeriod"},                // 1., 2., 3.
            {BulletAutonumberType::ArabicPlain, "arabicPlain"},                  // 1, 2, 3
            {BulletAutonumberType::RomanLowcaseParentheses, "romanLcParenBoth"}, // (i), (ii), (iii)
            {BulletAutonumberType::RomanUppercaseParentheses,
             "romanUcParenBoth"}, // (I), (II), (III)
            {BulletAutonumberType::RomanLowcaseRightParentheses, "romanLcParenR"}, // i), ii), iii)
            {BulletAutonumberType::RomanUppercaseRightParentheses, "romanUcParenR"}, // I), II), III)
            {BulletAutonumberType::RomanLowcasePeriod, "romanLcPeriod"},   // i., ii., iii.
            {BulletAutonumberType::RomanUppercasePeriod, "romanUcPeriod"}, //I., II, III.
            {BulletAutonumberType::Circle, "circleNumDbPlain"},
            {BulletAutonumberType::CircleWindingsBlack, "circleNumWdBlackPlain"},
            {BulletAutonumberType::CircleWindingsWhite, "circleNumWdWhitePlain"},
            {BulletAutonumberType::ArabicDoubleBytePeriod, "arabicDbPeriod"},
            {BulletAutonumberType::ArabicDoubleByte, "arabicDbPlain"},
            {BulletAutonumberType::SimplifiedChinesePeriod, "ea1ChsPeriod"},
            {BulletAutonumberType::SimplifiedChinese, "ea1ChsPlain"},
            {BulletAutonumberType::TraditionalChinesePeriod, "ea1ChtPeriod"},
            {BulletAutonumberType::TraditionalChinese, "ea1ChtPlain"},
            {BulletAutonumberType::JapanesePeriod, "ea1JpnChsDbPeriod"},
            {BulletAutonumberType::JapaneseKorean, "ea1JpnKorPlain"},
            {BulletAutonumberType::JapaneseKoreanPeriod, "ea1JpnKorPeriod"},
            {BulletAutonumberType::BidiArabic1Minus, "arabic1Minus"},
            {BulletAutonumberType::BidiArabic2Minus, "arabic2Minus"},
            {BulletAutonumberType::BidiHebrewMinus, "hebrew2Minus"},
            {BulletAutonumberType::ThaiAlphaPeriod, "thaiAlphaPeriod"},
            {BulletAutonumberType::ThaiAlphaRightParentheses, "thaiAlphaParenR"},
            {BulletAutonumberType::ThaiAlphaParentheses, "thaiAlphaParenBoth"},
            {BulletAutonumberType::ThaiNumberPeriod, "thaiNumPeriod"},
            {BulletAutonumberType::ThaiNumberRightParentheses, "thaiNumParenR"},
            {BulletAutonumberType::ThaiNumberParentheses, "thaiNumParenBoth"},
            {BulletAutonumberType::HindiAlphaPeriod, "hindiAlphaPeriod"},
            {BulletAutonumberType::HindiNumberPeriod, "hindiNumPeriod"},
            {BulletAutonumberType::HindiNumberParentheses, "hindiNumParenR"},
            {BulletAutonumberType::HindiAlpha1Period, "hindiAlpha1Period"},
        });
    SERIALIZE_ENUM(TabAlign,
                   {
                       {TabAlign::Left, "l"},
                       {TabAlign::Right, "r"},
                       {TabAlign::Center, "ctr"},
                       {TabAlign::Decimal, "dec"},
                   });
    SERIALIZE_ENUM(TextAlign, {
        {TextAlign::Left, "l"},
        {TextAlign::Right, "r"},
        {TextAlign::Center, "ctr"},
        {TextAlign::Justify, "just"},
        {TextAlign::JustifyLow, "justLow"},
        {TextAlign::Distrubute, "dist"},
        {TextAlign::DistrubuteThai, "thaiDist"},
    });
    SERIALIZE_ENUM(FontAlign, {
                      {FontAlign::Top, "t"},
                      {FontAlign::Bottom, "b"},
                      {FontAlign::Auto, "auto"},
                      {FontAlign::Center, "ctr"},
                      {FontAlign::Base, "base"},
                   });
    void readTabStops(QXmlStreamReader &reader);
    void writeTabStops(QXmlStreamWriter &writer, const QString &name) const;
};

class QXLSX_EXPORT ListStyleProperties
{
public:
    QVector<ParagraphProperties> vals;
    bool isEmpty() const;
    ParagraphProperties defaultParagraphProperties() const;
    ParagraphProperties &defaultParagraphProperties();
    void setDefaultParagraphProperties(const ParagraphProperties &properties);
    ParagraphProperties value(int level) const;
    void read(QXmlStreamReader &reader);
    void write(QXmlStreamWriter &writer, const QString &name) const;

    bool operator==(const ListStyleProperties &other) const;
    bool operator!=(const ListStyleProperties &other) const;
};

/**
 * @brief Represents a fragment of a text paragraph with its own character formatting.
 */
class QXLSX_EXPORT TextRun
{
public:
    /**
     * @brief specifies the text run type.
     */
    enum class Type
    {
        None, /**< @brief The text run is invalid */
        Regular, /**< @brief The text run is a plain text */
        LineBreak, /**< @brief The text run is a line break (like &lt;br&gt;) */
        TextField /**< @brief The text run is a generated text that the application should update periodically.

Each piece of text when it is generated is given a unique guid that is used to
refer to a specific field. At the time of creation the text field indicates the
kind of text that should be used to update this field. This update fieldType is
used so that all applications that did not create this text field can still know
what kind of text it should be updated with. Thus the new application can then
attach an update fieldType to the text field guid for continual updating.
*/
    };
    /**
     * @brief creates an invalid TextRun.
     */
    TextRun();
    TextRun(const TextRun &other);
    /**
     * @brief creates a new TextRun of type with text and properties.
     * @param type the type of a text run: Regular, LineBreak, TextField or None (invalid).
     * @param text the text of this fragment.
     * @param properties character properties of this fragment.
     */
    TextRun(Type type, const QString &text, const CharacterProperties &properties);
    /**
     * @brief creates a new TextRun of specified type.
     * @param type
     */
    TextRun(Type type);

    bool operator ==(const TextRun &other) const;
    bool operator !=(const TextRun &other) const;

    Type type = Type::None; /**< text run type */
    QString text; /**< a plain string */
    std::optional<CharacterProperties> characterProperties; /**< optional character properties */

    //text field properties
    std::optional<ParagraphProperties> paragraphProperties; /**< optional text field's paragraph properties */
    QString guid; /**< the unique to this document, host specified token that is
                       used to identify the field */
    QString fieldType; /**< the type of text that should be used to update this
                            text field. This is used to inform the rendering
                            application what text it should use to update this
                            text field. There are no specific syntax restrictions
                            placed on this attribute.

Reserved Values:

|Value | Description |
|----|----|
|slidenum | presentation slide number |
|datetime | default date time format for the rendering application |
|datetime1 | MM/DD/YYYY date time format [Example: 10/12/2007 ] |
|datetime2 | Day, Month DD, YYYY date time format [Example: Friday, October 12, 2007 ] |
|datetime3 | DD Month YYYY date time format [Example: 12 October 2007] |
|datetime4 | Month DD, YYYY date time format [Example: October 12, 2007] |
|datetime5 | DD-Mon-YY date time format [Example: 12-Oct-07] |
|datetime6 | Month YY date time format [Example: October 07 ] |
|datetime7 | Mon-YY date time format [Example: Oct-07 ] |
|datetime8 | MM/DD/YYYY hh:mm AM/PM date time format [Example: 10/12/2007 4:28 PM ] |
|datetime9 | MM/DD/YYYY hh:mm:ss AM/PM date time format [Example: 10/12/2007 4:28:34 PM] |
|datetime10| hh:mm date time format [Example: 16:28] |
|datetime11|  hh:mm:ss date time format [Example: 16:28:34] |
|datetime12|  hh:mm AM/PM date time format [Example: 4:28 PM] |
|datetime13 | hh:mm:ss: AM/PM date time format [Example: 4:28:34 PM] |

*/

    void read(QXmlStreamReader &reader);
    void write(QXmlStreamWriter &writer) const;
};

/**
 * @brief The class is used to represent a paragraph in the text.
 * Each paragraph has ParagraphProperties and zero to
 * infinity TextRun, i.e. fragments of text with the specific formatting.
 */
class QXLSX_EXPORT Paragraph
{
public:
    std::optional<ParagraphProperties> paragraphProperties;
    QList<TextRun> textRuns;
    /**
     * @brief the text run properties that are to be used if another run is inserted after the last run
specified.

This effectively saves the run property state so that it can be applied when the user enters additional
text. If this element is omitted, then the application can determine which default properties to apply. It is
recommended that this element be specified at the end of the list of text runs within the paragraph so that an
orderly list is maintained.
     */
    std::optional<CharacterProperties> endParagraphDefaultCharacterProperties;

    bool isValid() const;
    QString toPlainString() const;

    bool operator ==(const Paragraph &other) const;
    bool operator !=(const Paragraph &other) const;

    void read(QXmlStreamReader &reader);
    void write(QXmlStreamWriter &writer, const QString &name) const;
};

class TextFormatPrivate;

/**
 * @brief This is a helper class. Do not use it directly. Instead use Text and Title methods.
 */
class TextFormat
{
public:
    TextFormat();
    TextFormat(const TextFormat &other);
    ~TextFormat();
    TextFormat &operator=(const TextFormat &other);
    bool operator==(const TextFormat &other) const;
    bool operator!=(const TextFormat &other) const;

    bool isValid() const;
    bool isEmpty() const;
    /**
     * @brief returns this TextFormat as a plain string, keeping paragraphs and line breaks separated by "\n".
     * @return
     */
    QString toPlainString() const;

    void clear();
    void addParagraph(const Paragraph &p);
    void insertParagraph(int index, const Paragraph &p);
    int paragraphsCount() const;
    Paragraph paragraph(int index) const;
    Paragraph &paragraph(int index);
    Paragraph &lastParagraph();
    void setParagraph(int index, const Paragraph &paragraph);
    void removeParagraph(int index);
    void addFragment(const QString &text, const CharacterProperties &format);
    void addFragment(int paragraphIndex, const QString &text, const CharacterProperties &format);
    void insertFragment(int paragraphIndex, int fragmentIndex, const QString &text, const CharacterProperties &format);
    void addLineBreak(const CharacterProperties &format);
    void addLineBreak(int paragraphIndex, const CharacterProperties &format);
    void insertLineBreak(int paragraphIndex, int fragmentIndex, const CharacterProperties &format);
    void addTextField(const QString &guid, const QString &text, const QString &type);
    QString fragmentText(int paragraphIndex, int fragmentIndex) const;
    CharacterProperties fragmentFormat(int index) const;
    TextRun fragment(int paragraphIndex, int fragmentIndex) const;
    TextRun &fragment(int paragraphIndex, int fragmentIndex);
    TextRun &lastFragment();

    TextProperties textProperties() const;
    TextProperties &textProperties();
    void setTextProperties(const TextProperties &textProperties);
    ParagraphProperties defaultParagraphProperties() const;
    ParagraphProperties &defaultParagraphProperties();
    void setDefaultParagraphProperties(const ParagraphProperties &defaultParagraphProperties);
    CharacterProperties defaultCharacterProperties() const;
    CharacterProperties &defaultCharacterProperties();
    void setDefaultCharacterProperties(const CharacterProperties &defaultCharacterProperties);

    void read(QXmlStreamReader &reader);
    void write(QXmlStreamWriter &writer, const QString &name) const;
private:
    void readParagraph(QXmlStreamReader &reader);
    void writeParagraphs(QXmlStreamWriter &writer) const;

    QSharedDataPointer<TextFormatPrivate> d;
    friend QDebug operator<<(QDebug dbg, const TextFormat &f);
};

class TextPrivate;

/**
 * @brief The Text class represents text placed on charts, either as a part of Title (chart title, axis title),
 * or as a standalone entity (series name).
 * Text can be either plain string, string reference or rich-formatted string.
 */
class QXLSX_EXPORT Text
{
public:
    /**
     * @brief Specifies the Text type
     */
    enum class Type {
        None, /**< invalid Text */
        StringRef, /**< String reference */
        RichText, /**< Rich-formatted text */
        PlainText /**< Plain text */
    };

    /**
     * @brief creates invalid Text
     */
    Text();
    /**
     * @brief creates a plain Text with no rich formatting
     * @param text
     */
    explicit Text(const QString& text);
    /**
     * @brief creates empty Text of specified type
     * @param type
     */
    explicit Text(Type type); //null text of specific type

    Text(const Text &other);
    ~Text();
    Text &operator=(const Text &other);

    void setType(Type type);
    Type type() const;

    /**
     * @brief sets the text as a plain string (and changes type to be PlainText)
     * @param s
     */
    void setPlainString(const QString &s);
    QString toPlainString() const;
    /**
     * @brief sets text as a rich-formatted string (and changes type to be RichText).
     * @param text
     * @param detectHtml if set to false, the Text type will be RichText, but no formatting
     * will be set.
     */
    void setHtml(const QString &text, bool detectHtml = true);
    /**
     * @brief returns the rich-formatted text as an html-formatted string.
     * @warning not yet implemented. Returns an empty string.
     * @return
     */
    QString toHtml() const;

    void setStringReference(const QString &ref);
    void setStringCashe(const QStringList &cashe);
    QString stringReference() const;
    QStringList stringCashe() const;

    /**
     * @brief checks the Text type and returns whether the type is RichText
     * @return
     */
    bool isRichText() const;
    bool isPlainText() const;
    bool isStringReference() const;
    bool isEmpty() const;
    bool isValid() const;

    /**
     * @brief textProperties text body properties
     * @return
     */
    TextProperties textProperties() const;
    TextProperties &textProperties();
    void setTextProperties(const TextProperties &textProperties);

    ParagraphProperties defaultParagraphProperties() const;
    ParagraphProperties &defaultParagraphProperties();
    void setDefaultParagraphProperties(const ParagraphProperties &defaultParagraphProperties);

    CharacterProperties defaultCharacterProperties() const;
    CharacterProperties &defaultCharacterProperties();
    void setDefaultCharacterProperties(const CharacterProperties &defaultCharacterProperties);

    /**
     * @brief adds a new paragraph and sets the first text fragment.
     * @param text text of the first fragment.
     * @param characterProperties properties of the first fragment.
     * @param paragraphProperties properties of a new paragraph.
     */
    void addParagraph(const QString &text = QString(),
                      const CharacterProperties &characterProperties = CharacterProperties(),
                      const ParagraphProperties &paragraphProperties = ParagraphProperties());
    /**
     * @brief adds a new paragraph
     * @param paragraph
     */
    void addParagraph(const Paragraph &paragraph);
    /**
     * @brief inserts a new paragraph
     * @param index valid paragraph index [0..paragraphsCount]
     * @param text text of the first fragment.
     * @param characterProperties properties of the first fragment.
     * @param paragraphProperties properties of a new paragraph.
     */
    void insertParagraph(int index, const QString &text = QString(),
                         const CharacterProperties &characterProperties = CharacterProperties(),
                         const ParagraphProperties &paragraphProperties = ParagraphProperties());
    void insertParagraph(int index, const Paragraph &paragraph);
    /**
     * @brief returns a copy of the paragraph.
     * @param index paragraph index
     * @return a copy of the paragraph or invalid Paragraph if the index is out of bounds.
     */
    Paragraph paragraph(int index) const;
    /**
     * @brief returns a reference to the paragraph.
     * @param index paragraph index (must be valid).
     * @return a reference to the paragraph or fails if the the index is out of bounds.
     */
    Paragraph &paragraph(int index);
    /**
     * @brief returns the reference to the last paragraph.
     * @return a reference to the paragraph. If there's no paragraphs a new one will be created.
     */
    Paragraph &lastParagraph();
    /**
     * @brief sets the new value of a paragraph.
     * @param index paragraph index (must be valid).
     * @param paragraph new value of a paragraph.
     */
    void setParagraph(int index, const Paragraph &paragraph);
    /**
     * @brief returns paragraphs count.
     * @return
     */
    int paragraphsCount() const;
    /**
     * @brief removes paragraph from the text
     * @param index the paragraph index
     */
    void removeParagraph(int index);

    /**
     * @brief adds new text fragment to the last paragraph.
     * @param text plain text.
     * @param format text characters format.
     *
     * If there is no paragraphs, a new one will be created.
     */
    void addFragment(const QString &text,
                     const CharacterProperties &format = CharacterProperties());
    /**
     * @brief adds new text fragment to the paragraph with index paragraphIndex.
     * @param paragraphIndex the paragraph index.
     * @param text plain text.
     * @param format text characters format.
     */
    void addFragment(int paragraphIndex, const QString &text,
                     const CharacterProperties &format = CharacterProperties());
    void insertFragment(int paragraphIndex, int fragmentIndex, const QString &text,
                     const CharacterProperties &format = CharacterProperties());

    void addLineBreak(const CharacterProperties &format = CharacterProperties());
    void addLineBreak(int paragraphIndex, const CharacterProperties &format = CharacterProperties());
    void insertLineBreak(int paragraphIndex, int fragmentIndex, const CharacterProperties &format = CharacterProperties());

    /**
     * @brief adds a new text field to the last paragraph.
     * @param guid A non-empty text field guid.
     * @param text An initial text for the field.
     * @param type optional text field type.
     */
    void addTextField(const QString &guid, const QString &text = QString(), const QString &type = QString());
    /**
     * @brief fragmentText returns the text of the fragment from the paragraph.
     * @param paragraphIndex index of the paragraph.
     * @param fragmentIndex index of the text fragment (a.k.a. text run).
     * @return valid string if such paragraph and fragment exist, empty string otherwise.
     */
    QString fragmentText(int paragraphIndex, int fragmentIndex) const;
    /**
     * @brief fragmentFormat format of the fragment from the
     * last paragraph.
     * @param index index of the text fragment (a.k.a. text run)
     * @return
     */
    CharacterProperties fragmentFormat(int index) const;

    TextRun fragment(int paragraphIndex, int fragmentIndex) const;
    /**
     * @brief returns the reference to the fragment.
     * @param paragraphIndex valid paragraph index.
     * @param fragmentIndex valid fragment index.
     * @return the reference to the fragment or fails if no such fragment exists.
     */
    TextRun &fragment(int paragraphIndex, int fragmentIndex);

    /**
     * @brief returns the last fragment of the last paragraph.
     * @return reference to the fragment.
     * @note if the last paragraph has no fragments or there's no paragraphs
     * the method will create the default fragment of Type::Regular.
     */
    TextRun &lastFragment();

    void read(QXmlStreamReader &reader);
    void write(QXmlStreamWriter &writer, const QString &name) const;

    operator QVariant() const;

    bool operator ==(const Text &other) const;
    bool operator !=(const Text &other) const;
private:
    void readStringReference(QXmlStreamReader &reader);
    void readRichString(QXmlStreamReader &reader);
    void readPlainString(QXmlStreamReader &reader);

    void writeStringReference(QXmlStreamWriter &writer, const QString &name) const;
    void writePlainString(QXmlStreamWriter &writer, const QString &name) const;

    //TODO: all comparison operators
//    friend bool operator==(const Text &rs1, const Text &rs2);
//    friend bool operator!=(const Text &rs1, const Text &rs2);
//    friend bool operator<(const Text &rs1, const Text &rs2);
//    friend bool operator==(const Text &rs1, const QString &rs2);
//    friend bool operator==(const QString &rs1, const Text &rs2);
//    friend bool operator!=(const Text &rs1, const QString &rs2);
//    friend bool operator!=(const QString &rs1, const Text &rs2);
    friend QDebug operator<<(QDebug dbg, const Text &rs);

    QSharedDataPointer<TextPrivate> d;
};

QDebug operator<<(QDebug dbg, const Text &t);
QDebug operator<<(QDebug dbg, const TextProperties &t);
QDebug operator<<(QDebug dbg, const ListStyleProperties &t);
QDebug operator<<(QDebug dbg, const Paragraph &t);
QDebug operator<<(QDebug dbg, const ParagraphProperties &t);
QDebug operator<<(QDebug dbg, const CharacterProperties &t);
QDebug operator<<(QDebug dbg, const TextRun &t);
QDebug operator<<(QDebug dbg, const ParagraphProperties::TabAlign &t);
QDebug operator<<(QDebug dbg, const TextFormat &f);

}

Q_DECLARE_METATYPE(QXlsx::Text)
Q_DECLARE_METATYPE(QXlsx::TextFormat)

#endif // XLSXTEXT_H
