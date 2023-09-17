#ifndef XLSXMAIN_H
#define XLSXMAIN_H

#include <QFont>
#include <QColor>
#include <QByteArray>
#include <QList>
#include <QVariant>
#include <QXmlStreamWriter>
#include <QXmlStreamReader>
#include <QSharedData>
#include <QVector3D>
#include <QSize>
#include <optional>

#include "xlsxglobal.h"
#include "xlsxcolor.h"

namespace QXlsx {

Q_NAMESPACE

enum class ShapeType {
    line,
    lineInv,
    triangle,
    rtTriangle,
    rect,
    diamond,
    parallelogram,
    trapezoid,
    nonIsoscelesTrapezoid,
    pentagon,
    hexagon,
    heptagon,
    octagon,
    decagon,
    dodecagon,
    star4,
    star5,
    star6,
    star7,
    star8,
    star10,
    star12,
    star16,
    star24,
    star32,
    roundRect,
    round1Rect,
    round2SameRect,
    round2DiagRect,
    snipRoundRect,
    snip1Rect,
    snip2SameRect,
    snip2DiagRect,
    plaque,
    ellipse,
    teardrop,
    homePlate,
    chevron,
    pieWedge,
    pie,
    blockArc,
    donut,
    noSmoking,
    rightArrow,
    leftArrow,
    upArrow,
    downArrow,
    stripedRightArrow,
    notchedRightArrow,
    bentUpArrow,
    leftRightArrow,
    upDownArrow,
    leftUpArrow,
    leftRightUpArrow,
    quadArrow,
    leftArrowCallout,
    rightArrowCallout,
    upArrowCallout,
    downArrowCallout,
    leftRightArrowCallout,
    upDownArrowCallout,
    quadArrowCallout,
    bentArrow,
    uturnArrow,
    circularArrow,
    leftCircularArrow,
    leftRightCircularArrow,
    curvedRightArrow,
    curvedLeftArrow,
    curvedUpArrow,
    curvedDownArrow,
    swooshArrow,
    cube,
    can,
    lightningBolt,
    heart,
    sun,
    moon,
    smileyFace,
    irregularSeal1,
    irregularSeal2,
    foldedCorner,
    bevel,
    frame,
    halfFrame,
    corner,
    diagStripe,
    chord,
    arc,
    leftBracket,
    rightBracket,
    leftBrace,
    rightBrace,
    bracketPair,
    bracePair,
    straightConnector1,
    bentConnector2,
    bentConnector3,
    bentConnector4,
    bentConnector5,
    curvedConnector2,
    curvedConnector3,
    curvedConnector4,
    curvedConnector5,
    callout1,
    callout2,
    callout3,
    accentCallout1,
    accentCallout2,
    accentCallout3,
    borderCallout1,
    borderCallout2,
    borderCallout3,
    accentBorderCallout1,
    accentBorderCallout2,
    accentBorderCallout3,
    wedgeRectCallout,
    wedgeRoundRectCallout,
    wedgeEllipseCallout,
    cloudCallout,
    cloud,
    ribbon,
    ribbon2,
    ellipseRibbon,
    ellipseRibbon2,
    leftRightRibbon,
    verticalScroll,
    horizontalScroll,
    wave,
    doubleWave,
    plus,
    flowChartProcess,
    flowChartDecision,
    flowChartInputOutput,
    flowChartPredefinedProcess,
    flowChartInternalStorage,
    flowChartDocument,
    flowChartMultidocument,
    flowChartTerminator,
    flowChartPreparation,
    flowChartManualInput,
    flowChartManualOperation,
    flowChartConnector,
    flowChartPunchedCard,
    flowChartPunchedTape,
    flowChartSummingJunction,
    flowChartOr,
    flowChartCollate,
    flowChartSort,
    flowChartExtract,
    flowChartMerge,
    flowChartOfflineStorage,
    flowChartOnlineStorage,
    flowChartMagneticTape,
    flowChartMagneticDisk,
    flowChartMagneticDrum,
    flowChartDisplay,
    flowChartDelay,
    flowChartAlternateProcess,
    flowChartOffpageConnector,
    actionButtonBlank,
    actionButtonHome,
    actionButtonHelp,
    actionButtonInformation,
    actionButtonForwardNext,
    actionButtonBackPrevious,
    actionButtonEnd,
    actionButtonBeginning,
    actionButtonReturn,
    actionButtonDocument,
    actionButtonSound,
    actionButtonMovie,
    gear6,
    gear9,
    funnel,
    mathPlus,
    mathMinus,
    mathMultiply,
    mathDivide,
    mathEqual,
    mathNotEqual,
    cornerTabs,
    squareTabs,
    plaqueTabs,
    chartX,
    chartStar,
    chartPlus,
};

Q_ENUM_NS(ShapeType)

enum class TextShapeType
{
    textNoShape,
    textPlain,
    textStop,
    textTriangle,
    textTriangleInverted,
    textChevron,
    textChevronInverted,
    textRingInside,
    textRingOutside,
    textArchUp,
    textArchDown,
    textCircle,
    textButton,
    textArchUpPour,
    textArchDownPour,
    textCirclePour,
    textButtonPour,
    textCurveUp,
    textCurveDown,
    textCanUp,
    textCanDown,
    textWave1,
    textWave2,
    textDoubleWave1,
    textWave4,
    textInflate,
    textDeflate,
    textInflateBottom,
    textDeflateBottom,
    textInflateTop,
    textDeflateTop,
    textDeflateInflate,
    textDeflateInflateDeflate,
    textFadeRight,
    textFadeLeft,
    textFadeUp,
    textFadeDown,
    textSlantUp,
    textSlantDown,
    textCascadeUp,
    textCascadeDown,
};

Q_ENUM_NS(TextShapeType)

enum class CameraType
{
    legacyObliqueTopLeft,
    legacyObliqueTop,
    legacyObliqueTopRight,
    legacyObliqueLeft,
    legacyObliqueFront,
    legacyObliqueRight,
    legacyObliqueBottomLeft,
    legacyObliqueBottom,
    legacyObliqueBottomRight,
    legacyPerspectiveTopLeft,
    legacyPerspectiveTop,
    legacyPerspectiveTopRight,
    legacyPerspectiveLeft,
    legacyPerspectiveFront,
    legacyPerspectiveRight,
    legacyPerspectiveBottomLeft,
    legacyPerspectiveBottom,
    legacyPerspectiveBottomRight,
    orthographicFront,
    isometricTopUp,
    isometricTopDown,
    isometricBottomUp,
    isometricBottomDown,
    isometricLeftUp,
    isometricLeftDown,
    isometricRightUp,
    isometricRightDown,
    isometricOffAxis1Left,
    isometricOffAxis1Right,
    isometricOffAxis1Top,
    isometricOffAxis2Left,
    isometricOffAxis2Right,
    isometricOffAxis2Top,
    isometricOffAxis3Left,
    isometricOffAxis3Right,
    isometricOffAxis3Bottom,
    isometricOffAxis4Left,
    isometricOffAxis4Right,
    isometricOffAxis4Bottom,
    obliqueTopLeft,
    obliqueTop,
    obliqueTopRight,
    obliqueLeft,
    obliqueRight,
    obliqueBottomLeft,
    obliqueBottom,
    obliqueBottomRight,
    perspectiveFront,
    perspectiveLeft,
    perspectiveRight,
    perspectiveAbove,
    perspectiveBelow,
    perspectiveAboveLeftFacing,
    perspectiveAboveRightFacing,
    perspectiveContrastingLeftFacing,
    perspectiveContrastingRightFacing,
    perspectiveHeroicLeftFacing,
    perspectiveHeroicRightFacing,
    perspectiveHeroicExtremeLeftFacing,
    perspectiveHeroicExtremeRightFacing,
    perspectiveRelaxed,
    perspectiveRelaxedModerately,
};

Q_ENUM_NS(CameraType)

enum class LightRigType
{
    legacyFlat1,
    legacyFlat2,
    legacyFlat3,
    legacyFlat4,
    legacyNormal1,
    legacyNormal2,
    legacyNormal3,
    legacyNormal4,
    legacyHarsh1,
    legacyHarsh2,
    legacyHarsh3,
    legacyHarsh4,
    threePt,
    balanced,
    soft,
    harsh,
    flood,
    contrasting,
    morning,
    sunrise,
    sunset,
    chilly,
    freezing,
    flat,
    twoPt,
    glow,
    brightRoom,
};

Q_ENUM_NS(LightRigType)

enum class LightRigDirection
{
    TopLeft,
    Top,
    TopRight,
    Left,
    Right,
    BottomLeft,
    Bottom,
    BottomRight
};

class Extension
{
public:

};

class QXLSX_EXPORT ExtensionList
{
public:
    void write(QXmlStreamWriter &writer, const QString &name) const;
    void read(QXmlStreamReader &reader);
    bool isValid() const;
    bool operator==(const ExtensionList &other) const;
    bool operator!=(const ExtensionList &other) const;
private:
    QByteArray vals;
};

/**
 * @brief The RelativeRect class specifies the rectangle with edges relative to
 * the parent's rectangle.
 * Each edge of the rectangle is defined by a percentage offset from the
 * corresponding edge of the parent's bounding box. A positive percentage specifies
 * an inset, while a negative percentage specifies an outset.
 */
class RelativeRect
{
public:
    /**
     * @brief creates an invalid RelativeRect.
     */
    RelativeRect();
    /**
     * @brief creates a RelativeRect with fully specified edges.
     * @param left a percentage offset from the left edge of the parent's bounding box
     * @param top a percentage offset from the top edge of the parent's bounding box
     * @param right a percentage offset from the right edge of the parent's bounding box
     * @param bottom a percentage offset from the bottom edge of the parent's bounding box
     *
     * Values of 100 equal 100%. Values can be positive or negative.
     */
    RelativeRect(double left, double top, double right, double bottom);
    std::optional<double> left; /**< @brief a percentage offset from the left edge of the parent's bounding box */
    std::optional<double> top; /**< @brief a percentage offset from the top edge of the parent's bounding box */
    std::optional<double> right; /**< @brief right a percentage offset from the right edge of the parent's bounding box */
    std::optional<double> bottom; /**< @brief bottom a percentage offset from the bottom edge of the parent's bounding box */

    void write(QXmlStreamWriter &writer, const QString &name) const;
    void read(QXmlStreamReader &reader);
    bool isValid() const;
    bool operator==(const RelativeRect &other) const;
    bool operator!=(const RelativeRect &other) const;
};

QDebug operator<<(QDebug dbg, const RelativeRect &r);

/**
 * @brief The NumberFormat class represents the number formatting for
 * the parent element.
 */
class QXLSX_EXPORT NumberFormat
{
public:
    NumberFormat();
    NumberFormat(const QString &format);
    NumberFormat(const NumberFormat &other);

    bool operator==(const NumberFormat &other) const;
    bool operator!=(const NumberFormat &other) const;

    /**
     * @brief format a string representing the format code to apply.
     *
     *
     */
    QString format = QLatin1String("General");
    /**
     * @brief sourceLinked
     */
    bool sourceLinked = true;

    void write(QXmlStreamWriter &writer, const QString &name) const;
    void read(QXmlStreamReader &reader);
};

QDebug operator<<(QDebug dbg, const NumberFormat &f);

/**
 * @brief The Coordinate class represents a coordinate, either as a whole number in EMU,
 * as a number of points (i.e. EMU / 12700)
 * or as a qualified measure in mm, pt, in etc.
 */
class QXLSX_EXPORT Coordinate
{
public:
    /**
     * @brief Coordinate creates invalid coordinate
     */
    Coordinate() {}

    /**
     * @brief Coordinate creates Coordinate as a number of EMU
     * @param val coordinate value [-27273042329600 .. -27273042329600]
     */
    explicit Coordinate(qint64 emu);

    /**
     * @brief Coordinate creates universal coordinate measure
     * @param val string in the format "-?[0-9]+(\.[0-9]+)?(mm|cm|in|pt|pc|pi)",
     * f.e. "-12.3pt" or "200mm"
     */
    explicit Coordinate(const QString &val);

    /**
     * @brief Coordinate creates Coordinate in points
     * @param points coordinate value
     */
    explicit Coordinate(double points);

    /**
     * @brief returns coordinate as EMU (i.e. points * 12700).
     * @return EMU integer value in EMUs (1 pt is 12700 EMU).
     */
    qint64 toEMU() const;

    /**
     * @brief returns string representation of the coordinate.
     * @return
     */
    QString toString() const;

    /**
     * @brief returns coordinate in points (i.e. EMU / 12700)
     * @return
     */
    double toPoints() const;

    /**
     * @brief returns coordinate in pixels.
     * @return pixels at 96 DPI (default).
     */
    double toPixels(int dpi = 96) const;

    void setEMU(qint64 val);
    void setString(const QString &val);
    void setPoints(double points);
    /**
     * @brief sets the coordinate value in pixels (at 96 DPI by default)
     * and stores it in EMUs.
     * @param pixels new value in px.
     */
    void setPixels(double pixels, int dpi = 96);

    /**
     * @brief create parses val and creates a valid Coordinate
     * @param val string, either EMU (381000) or measure (30pt)
     * @return
     */
    static Coordinate create(const QStringView &val);

    bool isValid() const;

    bool operator == (const Coordinate &other) const;
    bool operator != (const Coordinate &other) const;
private:
    QVariant val;
};

/**
 * @brief The TextPoint class represents character spacing, either in points,
 * or as a qualified measure in mm, pt, in etc.
 */
class QXLSX_EXPORT TextPoint
{
public:
    /**
     * @brief TextPoint creates invalid spacing
     */
    TextPoint() {}

    /**
     * @brief TextPoint creates TextPoint in points
     * @param points spacing value [-4000.0 .. 4000.0]
     */
    explicit TextPoint(double points);

    /**
     * @brief TextPoint creates spacing measure
     * @param val string in the format "-?[0-9]+(\.[0-9]+)?(mm|cm|in|pt|pc|pi)",
     * f.e. "-12.3pt" or "200mm"
     */
    explicit TextPoint(const QString &val);

    /**
     * @brief toString returns string representation of spacing
     * @return
     */
    QString toString() const;

    /**
     * @brief toPoints return spacing in points
     * @return
     */
    double toPoints() const;

    void setString(const QString &val);
    void setPoints(double points);

    bool operator==(const TextPoint &other) const {return val == other.val;}
    bool operator!=(const TextPoint &other) const {return val != other.val;}

    /**
     * @brief create parses val and creates a valid TextPoint
     * @param val string, either pt*100 ([-400000..400000]) or measure (30pt)
     * @return
     */
    static TextPoint create(const QString &val);
private:
    QVariant val;
};

void parseAttribute(const QXmlStreamAttributes &a, const QLatin1String &name, Coordinate &target);

/**
 * @brief The Angle class represents an angle in 60,000ths of a degree. Positive
 * angles are clockwise (i.e., towards the positive y axis); negative angles are
 * counter-clockwise (i.e., towards the negative y axis).
 */
class QXLSX_EXPORT Angle
{
public:
    /**
     * @brief Angle creates an invalid angle
     */
    Angle() {}

    /**
     * @brief creates an angle in 60,000ths of a degree
     * @param angleInFrac the angle value in 60,000ths of a degree
     */
    Angle(qint64 angleInFrac);

    /**
     * @brief creates an angle in degrees
     * @param angleInDegrees the angle value in degrees
     */
    Angle(double angleInDegrees);

    qint64 toFrac() const;
    double toDegrees() const;
    QString toString() const;

    void setFrac(qint64 frac);
    void setDegrees(double degrees);
    bool isValid() const;

    /**
     * @brief parses s and creates an Angle
     * @param s string representation of an angle
     * @return
     */
    static Angle create(const QString &s);

    bool operator == (const Angle &other) const;
    bool operator != (const Angle &other) const;
private:
    std::optional<qint64> val;
};

void parseAttribute(const QXmlStreamAttributes &a, const QLatin1String &name, Angle &target);

class QXLSX_EXPORT Transform2D
{
public:
    std::optional<QPoint> offset; //element ext, optional
    std::optional<QSize> extension; //element off, optional
    Angle rotation; //attribute rot, optional
    std::optional<bool> flipHorizontal; //attribute flipH, optional
    std::optional<bool> flipVertical; //attribute flipV, optional

    bool isValid() const;

    void write(QXmlStreamWriter &writer, const QString &name) const;
    void read(QXmlStreamReader &reader);

    bool operator == (const Transform2D &other) const;
    bool operator != (const Transform2D &other) const;
};

struct QXLSX_EXPORT GeometryGuide
{
    GeometryGuide() {}
    GeometryGuide(const QString &name, const QString &formula) :
        name(name), formula(formula) {}
    QString name;
    QString formula;

    bool isValid() const {return !name.isEmpty() || !formula.isEmpty();}

    bool operator == (const GeometryGuide &other) const
    {
        return name == other.name && formula == other.formula;
    }
    bool operator != (const GeometryGuide &other) const
    {
        return name != other.name || formula != other.formula;
    }
};

class QXLSX_EXPORT PresetGeometry2D
{
public:
    PresetGeometry2D();
    PresetGeometry2D(ShapeType presetShape);
    QList<GeometryGuide> avLst; //element, optional
    ShapeType presetShape; // attribute, required

    void write(QXmlStreamWriter &writer, const QString &name) const;
    void read(QXmlStreamReader &reader);

    bool operator == (const PresetGeometry2D &other) const;
    bool operator != (const PresetGeometry2D &other) const;
};

class QXLSX_EXPORT PresetTextShape
{
public:
    PresetTextShape() {}
    PresetTextShape(TextShapeType textShapeType) : prst{textShapeType} {}

    QList<GeometryGuide> avLst; //element, optional
    TextShapeType prst; // attribute, required

    bool operator ==(const PresetTextShape &other) const;
    bool operator !=(const PresetTextShape &other) const;

    void write(QXmlStreamWriter &writer, const QString &name) const;
    void read(QXmlStreamReader &reader);
};

class QXLSX_EXPORT Scene3D
{
//    <xsd:complexType name="CT_Scene3D">
//      <xsd:sequence>
//        <xsd:element name="camera" type="CT_Camera" minOccurs="1" maxOccurs="1"/>
//        <xsd:element name="lightRig" type="CT_LightRig" minOccurs="1" maxOccurs="1"/>
//        <xsd:element name="backdrop" type="CT_Backdrop" minOccurs="0" maxOccurs="1"/>
//        <xsd:element name="extLst" type="CT_OfficeArtExtensionList" minOccurs="0" maxOccurs="1"/>
//      </xsd:sequence>
//    </xsd:complexType>
public:
    class Camera
    {
    public:
        QVector<Angle> rotation; //element, optional
        CameraType type = CameraType::obliqueTop; //atribute, required
        Angle fovAngle; //attribute, optional
        std::optional<double> zoom; /**< @brief camera zoom, in %. Must be positive. Defaults to 100. */
        bool operator ==(const Camera &other) const;
        bool operator !=(const Camera &other) const;
    };
    class LightRig
    {
    public:
        QVector<Angle> rotation; //element, optional
        LightRigType type = LightRigType::balanced; //attribute, required
        LightRigDirection lightDirection = LightRigDirection::Top; //attribute, required
        bool operator ==(const LightRig &other) const;
        bool operator !=(const LightRig &other) const;
    };
    class BackdropPlane
    {
    public:
        QVector<Coordinate> anchor;
        QVector<Coordinate> norm;
        QVector<Coordinate> up;
        bool operator ==(const BackdropPlane &other) const;
        bool operator !=(const BackdropPlane &other) const;
    };

    Camera camera; //element, required
    LightRig lightRig; //element, required
    std::optional<BackdropPlane> backdropPlane; //element, optional

    bool isValid() const;

    bool operator ==(const Scene3D &other) const;
    bool operator !=(const Scene3D &other) const;

    void write(QXmlStreamWriter &writer, const QString &name) const;
    void read(QXmlStreamReader &reader);
private:
    void readCamera(QXmlStreamReader &reader);
    void writeCamera(QXmlStreamWriter &writer, const QString &name) const;
    void readLightRig(QXmlStreamReader &reader);
    void writeLightRig(QXmlStreamWriter &writer, const QString &name) const;
    void readBackdrop(QXmlStreamReader &reader);
    void writeBackdrop(QXmlStreamWriter &writer, const QString &name) const;
};

enum class BevelType
{
    relaxedInset,
    circle,
    slope,
    cross,
    angle,
    softRound,
    convex,
    coolSlant,
    divot,
    riblet,
    hardEdge,
    artDeco,
};

Q_ENUM_NS(BevelType)


class QXLSX_EXPORT Bevel
{
public:
//    <xsd:complexType name="CT_Bevel">
//        <xsd:attribute name="w" type="ST_PositiveCoordinate" use="optional" default="76200"/>
//        <xsd:attribute name="h" type="ST_PositiveCoordinate" use="optional" default="76200"/>
//        <xsd:attribute name="prst" type="ST_BevelPresetType" use="optional" default="circle"/>
//      </xsd:complexType>
    Coordinate width;
    Coordinate height;
    std::optional<BevelType> type;

    bool operator ==(const Bevel &other) const;
    bool operator !=(const Bevel &other) const;

    void write(QXmlStreamWriter &writer, const QString &name) const;
    void read(QXmlStreamReader &reader);
};

enum class MaterialType
{
    legacyMatte,
    legacyPlastic,
    legacyMetal,
    legacyWireframe,
    matte,
    plastic,
    metal,
    warmMatte,
    translucentPowder,
    powder,
    dkEdge,
    softEdge,
    clear,
    flat,
    softmetal,
};

Q_ENUM_NS(MaterialType)

class QXLSX_EXPORT Shape3D
{
//    <xsd:complexType name="CT_Shape3D">
//      <xsd:sequence>
//        <xsd:element name="bevelT" type="CT_Bevel" minOccurs="0" maxOccurs="1"/>
//        <xsd:element name="bevelB" type="CT_Bevel" minOccurs="0" maxOccurs="1"/>
//        <xsd:element name="extrusionClr" type="CT_Color" minOccurs="0" maxOccurs="1"/>
//        <xsd:element name="contourClr" type="CT_Color" minOccurs="0" maxOccurs="1"/>
//        <xsd:element name="extLst" type="CT_OfficeArtExtensionList" minOccurs="0" maxOccurs="1"/>
//      </xsd:sequence>
//      <xsd:attribute name="z" type="ST_Coordinate" use="optional" default="0"/>
//      <xsd:attribute name="extrusionH" type="ST_PositiveCoordinate" use="optional" default="0"/>
//      <xsd:attribute name="contourW" type="ST_PositiveCoordinate" use="optional" default="0"/>
//      <xsd:attribute name="prstMaterial" type="ST_PresetMaterialType" use="optional"
//        default="warmMatte"/>
//    </xsd:complexType>
public:
    std::optional<Bevel> bevelTop;
    std::optional<Bevel> bevelBottom;
    Color extrusionColor;
    Color contourColor;
    Coordinate z;
    Coordinate extrusionHeight;
    Coordinate contourWidth;
    std::optional<MaterialType> material;

    bool operator ==(const Shape3D &other) const;
    bool operator !=(const Shape3D &other) const;

    void write(QXmlStreamWriter &writer, const QString &name) const;
    void read(QXmlStreamReader &reader);
};

QDebug operator<<(QDebug dbg, const Transform2D &f);
QDebug operator<<(QDebug dbg, const PresetGeometry2D &f);
QDebug operator<<(QDebug dbg, const PresetTextShape &f);
QDebug operator<<(QDebug dbg, const Scene3D &f);
QDebug operator<<(QDebug dbg, const Shape3D &f);
QDebug operator<<(QDebug dbg, const Coordinate &f);
QDebug operator<<(QDebug dbg, const GeometryGuide &f);
QDebug operator<<(QDebug dbg, const Bevel &f);
QDebug operator<<(QDebug dbg, const Scene3D::Camera &f);
QDebug operator<<(QDebug dbg, const Scene3D::LightRig &f);
QDebug operator<<(QDebug dbg, const Scene3D::BackdropPlane &f);
QDebug operator<<(QDebug dbg, const Angle &f);

}

#endif // XLSXMAIN_H
