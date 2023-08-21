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
 * @brief The TextProperties class represents the text body layout
 */
class QXLSX_EXPORT TextProperties
{
    //TODO: convert to shared data
public:
    enum class OverflowType
    {
        Overflow,
        Ellipsis,
        Clip
    };
    enum class VerticalType
    {
        EastAsianVertical,
        Horizontal,
        MongolianVertical,
        Vertical,
        Vertical270,
        WordArtVertical,
        WordArtVerticalRtl,
    };
    enum class Anchor
    {
        Bottom,
        Center,
        Distributed,
        Justified,
        Top
    };
    enum class TextAutofit
    {
        NoAutofit,
        NormalAutofit,
        ShapeAutofit
    };

    //Specifies whether the before and after paragraph spacing defined by the user is to be respected
    std::optional<bool> spcFirstLastPara;
    std::optional<OverflowType> verticalOverflow;
    std::optional<OverflowType> horizontalOverflow;
    std::optional<VerticalType> verticalOrientation;
    std::optional<Angle> rotation;
    std::optional<bool> wrap;
    Coordinate leftInset;
    Coordinate rightInset;
    Coordinate topInset;
    Coordinate bottomInset;
    std::optional<int> columnCount;
    Coordinate columnSpace;
    std::optional<bool> columnsRtl;
    std::optional<bool> fromWordArt;
    std::optional<Anchor> anchor;
    std::optional<bool> anchorCentering;
    std::optional<bool> forceAntiAlias;
    std::optional<bool> upright;
    std::optional<bool> compatibleLineSpacing;
    std::optional<PresetTextShape> textShape;

    // textScale and lineSpaceReduction are only valid if textAutofit == NormalAutofit
    std::optional<TextAutofit> textAutofit;
    std::optional<double> fontScale; // in %
    std::optional<double> lineSpaceReduction; // in %

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
 * @brief The CharacterProperties class
 * is used to set the text character properties
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
    std::optional<double> fontSize; /**<  @brief Specifies the size, in points, of text within a text run.*/
    std::optional<bool> bold; /**< @brief Specifies whether a run of text is formatted as bold text. */
    std::optional<bool> italic; /**< @brief Specifies whether a run of text is formatted as italic text. */
    std::optional<UnderlineType> underline; /**< @brief  Specifies whether a run of text is formatted as underlined text. */
    std::optional<StrikeType> strike; /**< @brief Specifies whether a run of text is formatted as strikethrough text. */
    std::optional<double> kerningFontSize; /**< @brief  Specifies the minimum font point size at which
                                                 character kerning occurs for this text run.

                                                 If this attribute is omitted, then kerning
                                                 occurs for all font sizes down to a 0 point
                                                 font. */
    std::optional<CapitalizationType> capitalization; /**< @brief  Specifies the capitalization that is to be applied to the text run. */
    std::optional<TextPoint> spacing; /**< @brief Specifies the spacing in points between characters within a text run. */
    std::optional<bool> normalizeHeights; /**< @brief Specifies the normalization of height that is
                                               to be applied to the text run. */
    std::optional<double> baseline;  /**< @brief Specifies the baseline for both the superscript
                                          and subscript fonts.

                                          The size is specified
                                          using a percentage where 1% is equal to
                                          1 percent of the font size and 100% is equal to
                                          100 percent font of the font size. */
    std::optional<bool> noProofing; /**< @brief Specifies that a run of text has been selected
                                         by the user to not be checked for mistakes.

                                         Therefore if there are spelling, grammar, etc
                                         mistakes within this text the generating
                                         application should ignore them. */
    std::optional<bool> proofingNeeded; /**< @brief  Specifies that the content of a text run
                                              has changed since the proofing tools have last
                                              been run.

                                              Effectively this flags text that
                                              is to be checked again by the generating
                                              application for mistakes such as spelling, grammar, etc. */
    std::optional<bool> checkForSmartTagsNeeded; /**< @brief Specifies whether or not a text run has been
                                                      checked for smart tags.

                                                      A value of
                                                      true here indicates to the generating application
                                                      that this text run should be checked for
                                                      smart tags. */
    std::optional<bool> spellingErrorFound; /**< @brief  Specifies that when this run of
                                                 text was checked for spelling, grammar, etc. that a
                                                 mistake was indeed found.

                                                 This allows
                                                 the generating application to effectively save the
                                                 state of the mistakes within the document instead
                                                 of having to perform a full pass check
                                                 upon opening the document. */
    std::optional<int> smartTagId; /**< @brief Specifies a smart tag identifier for a run of text.

                                        This ID is unique throughout the file and is used
                                        to reference corresponding auxiliary information about
                                        the smart tag. */
    QString bookmarkLinkTarget; /**< @brief Specifies the link target name that is used
                                     to reference to the proper link properties in a
                                     custom XML part within the document */
    LineFormat line; /**< @brief Specifies the line format to be applied to this text run. */
    FillFormat fill; /**< @brief Specifies the fill format to be applied to this text run. */
    Effect effect; /**< @brief Specifies the effects (blur, shadow etc.) to be applied to this text run. */
    Color highlightColor; /**< @brief Specifies the color to highlight this text run */

    /**
     * @brief Specifies that the stroke style of an underline for a run of text should be
     * the same as of the text run within which it is contained (that is #line).
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
     * @brief Specifies that the fill style of an underline for a run of text should be
     * the same as of the text run within which it is contained (that is #fill).
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
     * @brief Specifies a Latin Font to be used for a specific run of text.
     */
    Font latinFont;
    /**
     * @brief Specifies an East Asian Font to be used for a specific run of text.
     */
    Font eastAsianFont;
    /**
     * @brief Specifies a complex script (Arabic, Tamil etc.) Font to be used for a specific run of text.
     */
    Font complexScriptFont;
    /**
     * @brief Specifies a symbolic Font to be used for a specific run of text.
     */
    Font symbolFont;
    /**
     * @brief Specifies whether the contents of this run shall have right-to-left characteristics
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
 * @brief Represents a fragment of the text paragraph with its own character formatting.
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

    bool operator ==(const Paragraph &other) const;
    bool operator !=(const Paragraph &other) const;

    void read(QXmlStreamReader &reader);
    void write(QXmlStreamWriter &writer, const QString &name) const;
};

class TextPrivate;

class QXLSX_EXPORT Text
{
    //<xsd:complexType name="CT_Tx">
    //  <xsd:sequence>
    //    <xsd:choice minOccurs="1" maxOccurs="1">
    //      <xsd:element name="strRef" type="CT_StrRef" minOccurs="1" maxOccurs="1"/>
    //      <xsd:element name="rich" type="a:CT_TextBody" minOccurs="1" maxOccurs="1"/>
    //    </xsd:choice>
    //  </xsd:sequence>
    //</xsd:complexType>
public:
    enum class Type {
        None,
        StringRef,
        RichText,
        PlainText
    };

    Text(); // null text of type None
    explicit Text(const QString& text); // text of type PlainText with no formatting
    explicit Text(Type type); //null text of specific type

    Text(const Text &other);
    ~Text();

    void setType(Type type);
    Type type() const;

    void setPlainString(const QString &s);
    QString toPlainString() const;

    void setHtml(const QString &text, bool detectHtml = true);
    QString toHtml() const;

    void setStringReference(const QString &ref);
    void setStringCashe(const QStringList &cashe);
    QString stringReference() const;
    QStringList stringCashe() const;

    bool isRichString() const; //only checks type, ignores actual formatting
    bool isPlainString() const;
    bool isStringReference() const;
//    bool isNull() const;
//    bool isEmpty() const;
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

    void read(QXmlStreamReader &reader, bool diveInto = true);
    void write(QXmlStreamWriter &writer, const QString &name, bool diveInto = true) const;

    operator QVariant() const;

    Text &operator=(const Text &other);

    bool operator ==(const Text &other) const;
    bool operator !=(const Text &other) const;
private:
    void readStringReference(QXmlStreamReader &reader);
    void readRichString(QXmlStreamReader &reader);
    void readPlainString(QXmlStreamReader &reader);
    void readParagraph(QXmlStreamReader &reader);

    void writeStringReference(QXmlStreamWriter &writer, const QString &name) const;
    void writeRichString(QXmlStreamWriter &writer, const QString &name) const;
    void writePlainString(QXmlStreamWriter &writer, const QString &name) const;
    void writeParagraphs(QXmlStreamWriter &writer) const;

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

}

Q_DECLARE_METATYPE(QXlsx::Text)

#endif // XLSXTEXT_H
