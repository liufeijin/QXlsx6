// xlsxcolor_p.h

#ifndef QXLSX_XLSXCOLOR_H
#define QXLSX_XLSXCOLOR_H

#include <QtGlobal>
#include <QVariant>
#include <QColor>
#include <QXmlStreamWriter>
#include <QXmlStreamReader>

#include "xlsxglobal.h"
#include "xlsxutility_p.h"

namespace QXlsx {

//<xsd:group name="EG_ColorChoice">
//    <xsd:choice>
//      <xsd:element name="scrgbClr" type="CT_ScRgbColor" minOccurs="1" maxOccurs="1
//      <xsd:element name="srgbClr" type="CT_SRgbColor" minOccurs="1" maxOccurs="1
//      <xsd:element name="hslClr" type="CT_HslColor" minOccurs="1" maxOccurs="1
//      <xsd:element name="sysClr" type="CT_SystemColor" minOccurs="1" maxOccurs="1
//      <xsd:element name="schemeClr" type="CT_SchemeColor" minOccurs="1" maxOccurs="1
//      <xsd:element name="prstClr" type="CT_PresetColor" minOccurs="1" maxOccurs="1
//    </xsd:choice>
//  </xsd:group>

/**
 * @brief This class represents all color transformations for the specified color.
 * You can make the color lighter or darker, increase or decrease its components,
 * add alpha component etc.
 */
class ColorTransform
{
public:
    /**
     * @brief The Type enum represents all the available transforms for color.
     */
    enum class Type
    {
        Tint, /**< the color produced is a lighter version of the input color.
Specified as a positive percentage value (a 10.0 tint is 10% of the input color combined with 90% white.)*/
        Shade, /**< the color produced is a darker version of the color.
Specified as a positive percentage value (a 10.0 shade is 10% of the input color combined with 90% black.)*/
        Complement, /**< the color produced is the complement of the input color.
Specified as a bool value. */
        Inverse, /**< the color produced is the inverse of the input color.
Specified as a bool value. */
        Grayscale, /**< the color produced is the grayscale of the input color.
Specified as a bool value. */
        Alpha, /**< specifies (replaces) the opacity of the input color.
Specified as a positive percentage value (a 50.0 alpha gives 50% opacity to the color.) */
        AlphaOffset, /**< increases or decreases the input color opacity by the specified precentage offset.
A 10.0 alpha offset increases a 50% opacity to 60%. A -10.0 alpha offset decreases a 50% opacity to 40%.*/
        AlphaModulation, /**< specifies a more or less opaque version of its input color.
An alpha modulate never increases the alpha beyond 100%. A 200.0 alpha modulate makes
an input color twice as opaque as before. A 50.0 alpha modulate makes an input color
half as opaque as before.*/
        Hue, /**< specifies the input color with the specified hue, but with its saturation and luminance unchanged.
Specified as a degrees value (a 0.0 or 360.0 hue gives red, a 120.0 hue gives green)*/
        HueOffset, /**< specifies the input color with its hue shifted by a given angular offset,
but with its saturation and luminance unchanged.*/
        HueModulation, /**< specifies the input color with its hue modulated by the given percentage.
A 50.0 hue modulate decreases the angular hue value by half. A 200.0 hue modulate doubles the angular hue value.*/
        Saturation, /**< specifies the input color with the specified saturation, but with its luminance and hue unchanged.
Specified as a positive percentage value (0.0 to 100.0)*/
        SaturationOffset, /**< specifies the input color with its saturation shifted, but with its hue and luminance unchanged. A
10.0 offset to 20% saturation yields 30% saturation. */
        SaturationModulation, /**< specifies the input color with its saturation modulated by the given percentage.
A 50.0 saturation modulate reduces the saturation by half. A 200.0 saturation modulate doubles the saturation.*/
        Luminance, /**< specifies the input color with the specified luminance, but with its saturation and hue unchanged.
Specified as a positive percentage value (0.0 to 100.0)*/
        LuminanceOffset, /**< specifies the input color with its luminance shifted, but with its hue and saturation unchanged. A
10.0 offset to 20% luminance yields 30% luminance. */
        LuminanceModulation, /**< specifies the input color with its luminance modulated by the given percentage.
A 50.0 luminance modulate reduces the luminance by half. A 200.0 luminance modulate doubles the luminance.*/
        Red, /**< specifies the red component of the input color.
A value of 0.0 is minimal red, a value of 100.0 is maximal red.*/
        RedOffset, /**< specifies the red component as expressed by a percentage offset increase or decrease
        to the input color component.
        Increases never increase the red component beyond 100%, decreases never decrease the red component below 0%.*/
        RedModulation, /**< specifies the input color with its red component modulated by the given percentage.
A 50.0 red modulate reduces the red component by half. A 200.0 red modulate doubles the red component.*/
        Green, /**< specifies the green component of the input color.
A value of 0.0 is minimal green, a value of 100.0 is maximal green.*/
        GreenOffset, /**< specifies the green component as expressed by a percentage offset increase or decrease
        to the input color component.
        Increases never increase the green component beyond 100%, decreases never decrease the green component below 0%.*/
        GreenModulation, /**< specifies the input color with its green component modulated by the given percentage.
A 50.0 green modulate reduces the green component by half. A 200.0 green modulate doubles the green component.*/
        Blue, /**< specifies the blue component of the input color.
A value of 0.0 is minimal blue, a value of 100.0 is maximal blue.*/
        BlueOffset, /**< specifies the blue component as expressed by a percentage offset increase or decrease
        to the input color component.
        Increases never increase the blue component beyond 100%, decreases never decrease the blue component below 0%.*/
        BlueModulation, /**< specifies the input color with its blue component modulated by the given percentage.
A 50.0 blue modulate reduces the blue component by half. A 200.0 blue modulate doubles the blue component.*/
        Gamma, /**< specifies that the output color should be the sRGB gamma shift of the input color.
                    Specified as a bool value.
@note the following formula is applied to the r,g,b components of the input color:
\f$ v_gamma = v*12.92 if v<=0.0031308, 1.055*v^{1/2.4}-0.055 otherwise \f$


*/
        InverseGamma, /**< specifies that the output color should be the inverse sRGB gamma shift of the input color.
                    Specified as a bool value.
@note the following formula is applied to the r,g,b components of the input color:
\f$ v_degamma = v/12.92 if v<=0.04045, (\frac{v+0.055}{1.055})^2.4 otherwise \f$
*/
    };

    QMap<Type, QVariant> vals;

    operator QVariant() const;

    QColor transformed(const QColor &inputColor) const;
    bool hasTransform(Type type) const;
    QVariant transform(Type type) const;
    /**
     * @brief applies transform and returns a reference to the new color
     * @param type transform type
     * @param val transform parameter, see ColorTransform::Type
     * @param color input color
     * @return output color transformed.
     */
    static QColor &applyTransform(Type type, const QVariant &val, QColor &color);

    void read(QXmlStreamReader &reader);
    void write(QXmlStreamWriter &writer) const;
private:

    SERIALIZE_ENUM(Type, {
                       {Type::Tint, "tint"},
                       {Type::Shade, "shade"},
                       {Type::Complement, "comp"},
                       {Type::Inverse, "inv"},
                       {Type::Grayscale, "gray"},
                       {Type::Alpha, "alpha"},
                       {Type::AlphaOffset, "alphaOff"},
                       {Type::AlphaModulation, "alphaMod"},
                       {Type::Hue, "hue"},
                       {Type::HueOffset, "hueOff"},
                       {Type::HueModulation, "hueMod"},
                       {Type::Saturation, "sat"},
                       {Type::SaturationOffset, "satOff"},
                       {Type::SaturationModulation, "satMod"},
                       {Type::Luminance, "lum"},
                       {Type::LuminanceOffset, "lumOff"},
                       {Type::LuminanceModulation, "lumMod"},
                       {Type::Red, "red"},
                       {Type::RedOffset, "redOff"},
                       {Type::RedModulation, "redMod"},
                       {Type::Green, "green"},
                       {Type::GreenOffset, "greenOff"},
                       {Type::GreenModulation, "greenMod"},
                       {Type::Blue, "blue"},
                       {Type::BlueOffset, "blueOff"},
                       {Type::BlueModulation, "blueMod"},
                       {Type::Gamma, "gamma"},
                       {Type::InverseGamma, "invGamma"},
    });
};

#if !defined(QT_NO_DATASTREAM)
    QDataStream &operator<<(QDataStream &s, const ColorTransform &color);
    QDataStream &operator>>(QDataStream &s, ColorTransform &color);
#endif


class Color
{
public:
    /**
     * @brief The ColorType enum specifies the color type
     */
    enum class Type
    {
        Invalid, /**< @brief invalid color */
        Auto, /**< @brief auto color, that is system color dependent */
        Indexed, /**< @brief color specified by its index in a palette */
        RGB, /**< @brief color specified by RGB (red, green, blue) components*/
        HSL, /**< @brief color specified by HSL (hue, saturation, luminance) components*/
        System, /**< @brief color specified by a SystemColor enum*/
        Scheme, /**< @brief color specified by a SchemeColor enum*/
        Preset /**< @brief color specified by its string name*/
    };
    enum class SchemeColor
    {
        Background1,
        Text1,
        Background2,
        Text2,
        Accent1,
        Accent2,
        Accent3,
        Accent4,
        Accent5,
        Accent6,
        Hlink,
        FollowedHlink,
        Style,
        Dark1,
        Light1,
        Dark2,
        Light2,
    };

    enum class SystemColor
    {
        None,
        ScrollBar,
        Background,
        ActiveCaption,
        InactiveCaption,
        Menu,
        Window,
        WindowFrame,
        MenuText,
        WindowText,
        CaptionText,
        ActiveBorder,
        InactiveBorder,
        ApplicationWorkspace,
        Highlight,
        HighlightText,
        ButtonFace,
        ButtonShadow,
        GrayText,
        ButtonText,
        InactiveCaptionText,
        ButtonHighlight,
        DarkShadow3d,
        Light3d,
        InfoText,
        InfoBackground,
        HotLight,
        GradientActiveCaption,
        GradientInactiveCaption,
        MenuHighlight,
        MenuBar,
    };
    /**
     * @brief creates an invalid color.
     */
    Color();
    /**
     * @brief creates a named color.
     * @param colorName valid svg color name, f.e. "red" or "magenta" or "yellow".
     */
    Color(const QString &colorName);
    Color(const char *colorName);
    /**
     * @brief creates an RGB color.
     * @param color
     */
    Color(QColor color);
    /**
     * @brief creates a scheme color.
     * @param color SchemeColor enum.
     */
    Color(SchemeColor color);
    /**
     * @brief creates a system color.
     * @param color
     */
    Color(SystemColor color);
    Color(const Color &other);
    Color &operator=(const Color &other);
    ~Color();

    bool isValid() const;

    /**
     * @brief sets color as an (A)RGB value from QColor.
     * @param color
     */
    void setRgb(const QColor &color);
    void setHsl(const QColor &color);
    void setIndexedColor(int index);
    void setPresetColor(const QString &colorName);
    void setSchemeColor(SchemeColor color);
    void setSystemColor(SystemColor color);
    void setAutoColor();
    /**
     * @brief adds color transform to the list of transforms.
     * @param transform transform type
     * @param val transform parameter, see ColorTransform::Type.
     * @note only one transform of each type is supported, you cannot add two transforms of the same type.
     */
    void addTransform(ColorTransform::Type transform, QVariant val);
    ColorTransform transforms() const;
    /**
     * @brief applies all added color transforms and returns the result as QColor.
     * @return transformed color.
     */
    QColor transformed() const;
    /**
     * @brief returns whether the transforms list has the specified transform.
     * @param type the specified transform type.
     * @return
     */
    bool hasTransform(ColorTransform::Type type) const;
    /**
     * @brief returns the transform parameter if the transform with this type was added
     * @param type
     * @return valid QVariant if the transform was set, invalid one otherwise.
     */
    QVariant transform(ColorTransform::Type type) const;
    /**
     * @brief returns the color type
     * @return
     */
    Type type() const;
    /**
     * @brief returns whether the color has type Type::Auto
     */
    bool isAutoColor() const;

    /**
     * @brief returns QColor if the color type is Color::Type::RGB
     * @return valid QColor if the color type is Color::Type::RGB,
     * invalid QColor otherwise.
     */
    QColor rgb() const;
    /**
     * @brief returns QColor if the color type is Color::Type::HSL
     * @return valid QColor if the color type is Color::Type::HSL,
     * invalid QColor otherwise.
     */
    QColor hsl() const;
    /**
     * @brief attempts to get a valid QColor and returns the result.
     * @return valid QColor if the color type is Color::Type::RGB or Color::Type::HSL or Color::Type::Preset
     */
    QColor toQColor() const;
    /**
     * @brief returns an index of the indexed color.
     * @return color index starting from 0 or -1 if the color is not indexed.
     */
    int indexedColor() const;
    /**
     * @brief returns the preset color as a QColor
     * @return valid QColor if the color name is valid.
     */
    QColor presetColor() const;
    /**
     * @brief returns the preset color name
     * @return non-empty QString if the color is a preset color.
     */
    QString presetColorName() const;

    SchemeColor schemeColor() const;
    QString schemeColorName() const;

    SystemColor systemColor() const;
    QString systemColorName() const;

    operator QVariant() const;

    static QColor fromARGBString(const QString &c);
    static QString toARGBString(const QColor &c);
    static QString toRGBString(const QColor &c);

    bool write(QXmlStreamWriter &writer, const QString &node=QString()) const;
    bool read(QXmlStreamReader &reader);

    bool operator == (const Color &other) const;
    bool operator != (const Color &other) const;
private:
    bool readSimple(QXmlStreamReader &reader);
    bool readComplex(QXmlStreamReader &reader);
    SERIALIZE_ENUM(SchemeColor, {
        {SchemeColor::Background1, "bg1"},
        {SchemeColor::Text1, "tx1"},
        {SchemeColor::Background2, "bg2"},
        {SchemeColor::Text2, "bg1"},
        {SchemeColor::Accent1, "accent1"},
        {SchemeColor::Accent2, "accent2"},
        {SchemeColor::Accent3, "accent3"},
        {SchemeColor::Accent4, "accent4"},
        {SchemeColor::Accent5, "accent5"},
        {SchemeColor::Accent6, "accent6"},
        {SchemeColor::Hlink, "hlink"},
        {SchemeColor::FollowedHlink, "folHlink"},
        {SchemeColor::Style, "phClr"},
        {SchemeColor::Dark1, "dk1"},
        {SchemeColor::Light1, "lt1"},
        {SchemeColor::Dark2, "dk2"},
        {SchemeColor::Light2, "lt2"},
    });
    SERIALIZE_ENUM(SystemColor, {
        {SystemColor::ScrollBar, "scrollBar"},
        {SystemColor::Background, "background"},
        {SystemColor::ActiveCaption, "activeCaption"},
        {SystemColor::InactiveCaption, "inactiveCaption"},
        {SystemColor::Menu, "menu"},
        {SystemColor::Window, "window"},
        {SystemColor::WindowFrame, "windowFrame"},
        {SystemColor::MenuText, "menuText"},
        {SystemColor::WindowText, "windowText"},
        {SystemColor::CaptionText, "captionText"},
        {SystemColor::ActiveBorder, "activeBorder"},
        {SystemColor::InactiveBorder, "inactiveBorder"},
        {SystemColor::ApplicationWorkspace, "appWorkspace"},
        {SystemColor::Highlight, "highlight"},
        {SystemColor::HighlightText, "highlightText"},
        {SystemColor::ButtonFace, "btnFace"},
        {SystemColor::ButtonShadow, "btnShadow"},
        {SystemColor::GrayText, "grayText"},
        {SystemColor::ButtonText, "btnText"},
        {SystemColor::InactiveCaptionText, "inactiveCaptionText"},
        {SystemColor::ButtonHighlight, "btnHighlight"},
        {SystemColor::DarkShadow3d, "3dDkShadow"},
        {SystemColor::Light3d, "3dLight"},
        {SystemColor::InfoText, "infoText"},
        {SystemColor::InfoBackground, "infoBk"},
        {SystemColor::HotLight, "hotLight"},
        {SystemColor::GradientActiveCaption, "gradientActiveCaption"},
        {SystemColor::GradientInactiveCaption, "gradientInactiveCaption"},
        {SystemColor::MenuHighlight, "menuHighlight"},
        {SystemColor::MenuBar, "menuBar"},
    });
    SERIALIZE_ENUM(Type, {
        {Type::Invalid,     "invalid"},
        {Type::Auto,     "auto"},
        {Type::RGB,    "rgb"},
        {Type::HSL,    "hsl"},
        {Type::System, "system"},
        {Type::Scheme, "scheme"},
        {Type::Indexed, "simple"},
        {Type::Preset, "preset"},
    });
    QByteArray idKey() const;

    Type type_ = Type::Invalid;
    QVariant val;

    ColorTransform tr;

    // Only used if type == ColorType::SystemColor
    QColor lastColor;

    bool isCRGB = false;

    mutable bool isDirty = true;
    mutable QByteArray m_key;
#if !defined(QT_NO_DATASTREAM)
    friend QDataStream &operator<<(QDataStream &s, const Color &color);
    friend QDataStream &operator>>(QDataStream &s, Color &color);
#endif
    friend QDebug operator<<(QDebug dbg, const Color &c);
    friend uint qHash(const Color &c, uint seed) Q_DECL_NOTHROW;
};

uint qHash(const Color &c, uint seed = 0) Q_DECL_NOTHROW;

#if !defined(QT_NO_DATASTREAM)
QDataStream &operator<<(QDataStream &s, const Color &color);
QDataStream &operator>>(QDataStream &s, Color &color);
#endif

QDebug operator<<(QDebug dbg, const Color &c);

}

Q_DECLARE_METATYPE(QXlsx::ColorTransform)
Q_DECLARE_METATYPE(QXlsx::Color)


#endif // QXLSX_XLSXCOLOR_P_H
