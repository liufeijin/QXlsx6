#ifndef XLSXFILLFORMAT_H
#define XLSXFILLFORMAT_H

#include <QFont>
#include <QColor>
#include <QByteArray>
#include <QList>
#include <QVariant>
#include <QBrush>
#include <QXmlStreamWriter>
#include <QXmlStreamReader>
#include <QSharedData>

#include "xlsxglobal.h"
#include "xlsxcolor.h"
#include "xlsxmain.h"


namespace QXlsx {

class FillFormatPrivate;
class Workbook;
class Relationships;
class Effect;

/**
 * @brief Specifies fill properties for lines, shapes etc.
 *
 *  There are 6 types of filling. Each type has its own set of parameters.
 *
 *  ### 1. FillType::NoFill
 *
 *  A line or a shape is not filled, that is invisible.
 *
 *  ### 2. FillType::SolidFill
 *
 *  A line or a shape is filled with a solid color, that can be specified with
 *  #setColor() methods, either as an RGB color or as a system, theme or HSL
 *  color. See QColor and QXlsx::Color documentation.
 *
 *  ### 3. FillType::PatternFill
 *
 *  A line or a shape is filled with some pattern (See FillFormat::PatternType
 *  enum). A pattern is drawn with two colors: #foregroundColor() and
 *  #backgroundColor(). See #setForegroundColor(), #setBackgroundColor()
 *  methods.
 *
 *  ### 4. FillType::PictureFill
 *
 *  This type of fill is applicable to only shapes. A shape is filled with a
 *  specified picture that will be written in xlsx in PNG format (see
 *  #setPicture()).
 *
 *  You can also specify additional parameters of the picture fill: picture
 *  resolution with #setPictureDpi(), picture region with
 *  #setPictureSourceRect(), picture behavior with #setRotateWithShape(), some
 *  additional picture info (#setPictureCompression()) etc.
 *
 *  The picture can either be stretched or tiled. The stretch rectangle is set
 *  via #setPictureStretchRect(), the tile parameters via
 *  #setPictureHorizontalOffset(), #setPictureVerticalOffset(),
 *  #setPictureHorizontalScale(), #setPictureVerticalScale(), #setTileAlignment(),
 *  #setTileFlipMode().
 *
 *  ### 5. FillType::GradientFill
 *
 *  A line or a shape is filled with a gradient. There are two types of
 *  gradient fills: linear gradient and path gradient. Each has its own set of
 *  parameters.
 *
 *  ### 6. FillType::GroupFill
 *
 *  This type of fill is applicable to only shapes. A shape inherits the fill
 *  properties of a group it belongs.
 *
 *  Below is a table of methods and fill types:
 *
 *  Method | Fill type
 *  ----|----
 *  #color(), #setColor() | solid fill
 *  #gradientList(), #setGradientList(), addGradientStop(), #tileRect(), #setTileRect(), #rotateWithShape(), #setRotateWithShape() | gradient fills
 *  #linearShadeAngle(), #setLinearShadeAngle(), #linearShadeScaled(), #setLinearShadeScaled() | linear gradient fill
 *  #pathType(), #setPathType(), #pathRect(), #setPathRect(), #tileFlipMode(), #setTileFlipMode() | path gradient fill
 *  #foregroundColor(), #setForegroundColor(), #backgroundColor(), #setBackgroundColor(), #patternType(), #setPatternType() | pattern fill
 *  #picture(), #setPicture(), #pictureFillMode(), #setPictureFillMode(), #pictureSourceRect(), #setPictureSourceRect(), #pictureDpi(), #setPictureDpi(), #pictureStretchRect(), #setPictureStretchRect(), #pictureHorizontalOffset(), #setPictureHorizontalOffset(), #pictureVerticalOffset(), #setPictureVerticalOffset(), #pictureAlpha(), #setPictureAlpha(), #pictureCompression(), #setPictureCompression(), #tileAlignment(), #setTileAlignment() | picture fill
 *
 */
class QXLSX_EXPORT FillFormat
{
public:
    /**
     * @brief The FillType enum specifies the fill type for lines and shapes.
     *
     */
    enum class FillType {
        NoFill, /**< @brief line or shape is not filled (not drawn) */
        SolidFill, /**< @brief line or shape has one color */
        GradientFill, /**< @brief line or shape is drawn with gradient */
        PictureFill, /**< @brief shape is filled with a picture (not applicable to lines) */
        PatternFill, /**< @brief line or shape is filled with a predefined pattern */
        GroupFill  /**< @brief shape inherits the fill properties of a group (not applicable to lines)*/
    };
    /**
     * @brief The PathType enum specifies the shape of path to follow for a gradient.
     */
    enum class PathType
    {
        Shape, /**< Gradient follows the shape */
        Circle, /**< Gradient follows a circular path */
        Rectangle /**< Gradient follows a rectangular path */
    };
    /**
     * @brief The TileFlipMode enum specifies whether/how to flip the contents
     * of a tile region when using it to fill a larger fill region.
     */
    enum class TileFlipMode
    {
        None, /**< Tiles are not flipped */
        X, /**< Tiles are flipped horizontally*/
        Y, /**< Tiles are flipped vertically*/
        XY /**< Tiles are flipped horizontally and vertically*/
    };
    /**
     * @brief The PatternType enum specifies the pattern fill type.
     */
    enum class PatternType
    {
        Percent5,
        Percent10,
        Percent20,
        Percent25,
        Percent30,
        Percent40,
        Percent50,
        Percent60,
        Percent70,
        Percent75,
        Percent80,
        Percent90,
        Horizontal,
        Vertical,
        LightHorizontal,
        LightVertical,
        DarkHorizontal,
        DarkVertical,
        NarrowHorizontal,
        NarrowVertical,
        DashedHorizontal,
        DashedVertical,
        Cross,
        DownwardDiagonal,
        UpwardDiagonal,
        LightDownwardDiagonal,
        LightUpwardDiagonal,
        DarkDownwardDiagonal,
        DarkUpwardDiagonal,
        WideDownwardDiagonal,
        WideUpwardDiagonal,
        DashedDownwardDiagonal,
        DashedUpwardDiagonal,
        DiagonalCross,
        SmallCheckerBoard,
        LargeCheckerBoard,
        SmallGrid,
        LargeGrid,
        DottedGrid,
        SmallConfetti,
        LargeConfetti,
        HorizontalBrick,
        DiagonalBrick,
        SolidDiamond,
        OpenDiamond,
        DottedDiamond,
        Plaid,
        Sphere,
        Weave,
        Divot,
        Shingle,
        Wave,
        Trellis,
        ZigZag,
    };
    /**
     * @brief The Alignment enum specifies where to align the first tile with respect to the shape.
     */
    enum class Alignment
    {
        AlignTopLeft, /**< Tiling will begin from the top left corner of the shape's bounding rectangle. */
        AlignTop, /**< Tiling will begin from the top edge of the shape's bounding rectangle. */
        AlignTopRight, /**< Tiling will begin from the top right corner of the shape's bounding rectangle. */
        AlignRight, /**< Tiling will begin from the right edge of the shape's bounding rectangle. */
        AlignCenter, /**< Tiling will begin from the center of the shape's bounding rectangle. */
        AlignLeft, /**< Tiling will begin from the left edge of the shape's bounding rectangle. */
        AlignBottomLeft, /**< Tiling will begin from the bottom left corner of the shape's bounding rectangle. */
        AlignBottom, /**< Tiling will begin from the bottom edge of the shape's bounding rectangle. */
        AlignBottomRight, /**< Tiling will begin from the bottom right corner of the shape's bounding rectangle. */
    };

    /**
     * @brief The PictureCompression enum specifies the amount of compression that
     * has been used for a fill image or picture.
     */
    enum class PictureCompression
    {
        None, /**< No compression was used (default) */
        Email, /**< Compression size suitable for inclusion with email */
        Screen, /**< Compression size suitable for viewing on screen */
        Print, /**< Compression size suitable for printing */
        HqPrint /**< Compression size suitable for high quality printing */
    };
    /**
     * @brief The PictureFillMode enum specifies the mode of applying picture fill.
     */
    enum class PictureFillMode
    {
        Stretch, /**< The picture will be stretched to fill the shape. */
        Tile /**< The picture will be tiled to fill the shape */
    };

    /**
     * @brief creates an invalid fill. @see isValid().
     */
    FillFormat();
    /**
     * @brief creates a fill of specified type.
     * @param type a fill type.
     * @note There is a distinction between a default fill (when the FillFormat object
     * is not valid) and a transparent fill (the FillFormat object has FillType::NoFill).
     * Compare:
     *
     * @code
     * shape1.setFill(FillFormat::NoFill); // sets a transparent fill
     * shape2.setFill({}); //sets default fill, that is a solid fill with white color.
     * @endcode
     *
     */
    explicit FillFormat(FillType type);
    /**
     * @brief creates a fill from a QBrush object.
     * @param brush
     */
    FillFormat(const QBrush &brush);
    /**
     * @brief creates solid fill.
     * @param color the color to be used for a fill.
     */
    FillFormat(const QColor &color);
    /**
     * @brief creates solid fill.
     * @param color the color to be used for a fill.
     */
    FillFormat(const Color &color);
    /**
     * @brief creates pattern fill
     * @param foreground the foreground color
     * @param background the background color
     * @param pattern the pattern type
     */
    FillFormat(const QColor &foreground, const QColor &background, PatternType pattern);
    /**
     * @brief creates gradient fill with minimal parameters
     * @param gradientList a map of color stops.
     * @param pathType type of the gradient path.
     */
    FillFormat(const QMap<double, Color> &gradientList, PathType pathType);
    /**
     * @brief creates linear gradient fill with minimal parameters.
     * @param gradientList a map of color stops.
     */
    FillFormat(const QMap<double, Color> &gradientList);

    /**
     * @brief creates picture fill
     * @param picture a picture to be used for a fill.
     */
    FillFormat(const QImage &picture);

    FillFormat(const FillFormat &other);
    FillFormat &operator=(const FillFormat &rhs);
    ~FillFormat();

    /**
     * @brief returns the fill type.
     */
    FillType type() const;
    /**
     * @brief sets the fill type.
     * @param type A fill type.
     */
    void setType(FillType type);
    /**
     * @brief tries to get the @a brush attributes and sets the fill type and
     * parameters from @a brush.
     * @param brush QBrush object.
     *
     */
    void setBrush(const QBrush &brush);
    /**
     * @brief tries to convert the fill parameters to a QBrush object and
     * returns the result.
     */
    QBrush toBrush() const;

    /* Solid fill properties */
    /**
     * @brief returns the solid fill color
     * @return
     */
    Color color() const;
    /**
     * @brief sets the color for solid fill
     * @param color
     */
    void setColor(const Color &color);
    /**
     * @brief sets an rgb color for solid fill
     * @param color
     */
    void setColor(const QColor &color);

    /* Gradient fill properties */

    /**
     * @brief returns the gradients list for the gradient fill
     * @return QMap, where keys are color stops in percents, values are gradient
     * colors.
     * @note This method does not check the fill type to be FillType::GradientFill.
     */
    QMap<double, Color> gradientList() const;
    /**
     * @brief sets the gradient list for the gradient fill.
     * @param list a map with keys being the color stops (in percents), values
     * being the colors.
     *
     * The simplest gradient list:
     *
     * @code
     * setGradientList({{0, "red"}, {100, "blue"}});
     * @endcode
     *
     * @note This method does not check the fill type to be
     * FillType::GradientFill.
     */
    void setGradientList(const QMap<double, Color> &list);
    /**
     * @brief adds a gradient stop to the gradient list.
     * @param stop percentage value (100.0 is 100%).
     * @param color Color.
     * @note This method does not check the fill type to be
     * FillType::GradientFill.
     */
    void addGradientStop(double stop, const Color &color);
    /**
     * @brief returns the direction of color change for the linear gradient.
     * @return valid Angle if the parameter is set.
     *
     * The direction of color change is measured clockwise starting from the
     * horizontal.
     * @note This method does not check the fill type to be
     * FillType::GradientFill.
     */
    Angle linearShadeAngle() const;
    /**
     * @brief sets the direction of color change for the linear gradient.
     * @param val the direction measured clockwise starting from the horizontal.
     * Values should be in the range [0..360]
     * @warning This parameter is applicable to only linear gradient. Invoking
     * this method will clear the path gradient parameters (#pathType(),
     * #pathRect()).
     * @note This method does not check the fill type to be
     * FillType::GradientFill.
     */
    void setLinearShadeAngle(Angle val);
    /**
     * @brief returns whether the linear gradient angle scales with the fill
     * region.
     * If set to `true`, the linearShadeAngle() is scaled to the shape's fill
     * region. For example, gradient with an angle of 45° (i.e. vector (1, -1))
     * in a shape with width of 300 and height of 200 will be scaled to
     * vector(300,-200), that is an angle of 33.69°.
     * @note This method does not check the fill type to be
     * FillType::GradientFill.
     */
    std::optional<bool> linearShadeScaled() const;
    /**
     * @brief sets whether the linear gradient angle scales with the fill
     * region.
     * @param scaled If `true`, then the linearShadeAngle() will be scaled to
     * the shape's fill region.
     *
     * For example, gradient with an angle of 45° (i.e. vector (1, -1)) in a
     * shape with width of 300 and height of 200 will be scaled to
     * vector(300,-200), that is an angle of 33.69°.
     *
     * If `false`, the gradient angle is independent of the shape's fill region.
     * @warning This parameter is applicable to only linear gradient. Invoking
     * this method will clear the path gradient parameters (#pathType(),
     * #pathRect()).
     * @note This method does not check the fill type to be
     * FillType::GradientFill.
     */
    void setLinearShadeScaled(bool scaled);

    /**
     * @brief returns the type of the path gradient.
     * @note This method does not check the fill type to be
     * FillType::GradientFill.
     */
    std::optional<PathType> pathType() const;
    /**
     * @brief sets the type of the path gradient.
     * @param pathType type of the path gradient.
     * @warning This parameter is applicable to only path gradient. Invoking
     * this method will clear the linear gradient parameters
     * (#linearShadeAngle(), #linearShadeScaled()).
     * @note This method does not check the fill type to be
     * FillType::GradientFill.
     */
    void setPathType(PathType pathType);

    /**
     * @brief returns the "focus" rectangle for the first color in the gradient
     * list, specified relative to the fill tile rectangle.
     *
     * The first color in the gradient list fills the entire tile except for the
     * margins specified by pathRect.
     * @note This method does not check the fill type to be
     * FillType::GradientFill.
     */
    std::optional<RelativeRect> pathRect() const;
    /**
     * @brief sets the "fill" rectangle for the first color in the gradient
     * list, specified relative to the fill tile rectangle.
     *
     * The first color in the gradient list will fill the entire tile except for
     * the margins specified by #pathRect().
     *
     * For example, if you specify path rect as (50,50,50,50), this means that
     * the first color will fill the point in the center of the tile rect, that
     * is the default behavior.
     *
     * If you need to _expand_ the area filled with the first color, set path
     * rect with margins less than 50. If you set path rect as (0,0,0,0), this
     * means that the first color in the gradient list will fill the entire tile
     * rect. This is actually equivalent to filling the tile rect with solid
     * color.
     *
     * To specify the "focus point" (starting point) of the gradient make sure
     * the sum of the opposite margins in pathRect() is 100.0. For example, to
     * move the focus point to the center of the bottom edge use
     * `setPathRect(50,100,50,0);`
     *
     * @param rect
     * @warning This parameter is applicable to only path gradient. Invoking
     * this method will clear the linear gradient parameters
     * (#linearShadeAngle(), #linearShadeScaled()).
     * @note This method does not check the fill type to be
     * FillType::GradientFill.
     */
    void setPathRect(RelativeRect rect);

    /**
     * @brief returns a rectangular region of the shape to which a gradient fill
     * is applied.
     *
     * This region is then tiled across the remaining area of the shape to
     * complete the fill. The tile rectangle is defined by percentage offsets
     * from the sides of the shape's bounding box.
     * @return valid RelativeRect if the parameter is set, `nullopt` otherwise.
     * @note This method does not check the fill type to be
     * FillType::GradientFill.
     */
    std::optional<RelativeRect> tileRect() const;
    /**
     * @brief sets a rectangular region of the shape to which a gradient fill is
     * applied.
     *
     * This region is then tiled across the remaining area of the shape to
     * complete the fill.
     *
     * The tile rectangle is defined by percentage offsets from the sides of the
     * shape's bounding box.
     *
     * This parameter is applicable to both linear and path gradients.
     * @param rect Tile rectangle defined by percentage offsets from the sides
     * of the shape's bounding box.
     * @note This method does not check the fill type to be
     * FillType::GradientFill.
     */
    void setTileRect(RelativeRect rect);
    /**
     * @brief returns the flipping of a #tileRect() in the #pathRect().
     * @return TileFlipMode enum value if flipping was set, `nullopt` otherwise.
     */
    std::optional<FillFormat::TileFlipMode> tileFlipMode() const;
    /**
     * @brief sets the flipping of a #tileRect() in the #pathRect().
     * @param tileFlipMode Tiles can be flipped horizontally, vertically,
     * horizontally and vertically or not flipped at all.
     */
    void setTileFlipMode(TileFlipMode tileFlipMode);
    /**
     * @brief returns whether the gradient or picture should be rotated with the
     * shape rotation.
     *
     * Value of `true` means that when the shape that has been filled with a
     * picture or a gradient is transformed with a rotation then the fill is
     * transformed with the same rotation.
     *
     * @return valid optional value if the parameter is set, `nullopt`
     * otherwise.
     */
    std::optional<bool> rotateWithShape() const;
    /**
     * @brief sets whether the gradient or picture should be rotated with the
     * shape rotation.
     * @param val `true` means that when the shape that has been filled
     * with a picture or a gradient is transformed with a rotation then the fill
     * is transformed with the same rotation.
     *
     * If not set, the default value is `false`.
     */
    void setRotateWithShape(bool val);

    /* Pattern fill properties */
    /**
     * @brief returns the foreground color of the pattern fill.
     * @return Valid Color if foregroundColor was set.
     * @note This method does not check the fill type to be
     * FillType::PatternFill.
     */
    Color foregroundColor() const;
    /**
     * @brief sets the foreground color of the pattern fill.
     * @param color the foreground color of the pattern fill.
     * @note This method does not check the fill type to be
     * FillType::PatternFill.
     */
    void setForegroundColor(const Color &color);
    /**
     * @brief returns the background color of the pattern fill.
     * @return Valid Color if backgroundColor was set.
     * @note This method does not check the fill type to be
     * FillType::PatternFill.
     */
    Color backgroundColor() const;
    /**
     * @brief sets the background color of the pattern fill.
     * @param color the background color of the pattern fill.
     * @note This method does not check the fill type to be
     * FillType::PatternFill.
     */
    void setBackgroundColor(const Color &color);
    /**
     * @brief returns the type of the pattern fill.
     * @return pattern type if it was set, `nullopt` otherwise.
     * @note This method does not check the fill type to be
     * FillType::PatternFill.
     */
    std::optional<PatternType> patternType();
    /**
     * @brief sets the type of the pattern fill.
     * @param patternType pattern type.
     * @note This method does not check the fill type to be
     * FillType::PatternFill.
     */
    void setPatternType(PatternType patternType);

    /* Blip fill properties */
    /**
     * @brief sets the picture to be used for the fill.
     *
     * This method can be used to set the picture for the fill. The picture will
     * be written in xlsx file as PNG.
     *
     * You can also specify additional parameters of the picture fill:
     * #setPictureDpi(), #setRotateWithShape(), #setPictureSourceRect(). The
     * picture can either be stretched or tiled. The stretch rectangle is set
     * via #setPictureStretchRect(), the tile parameters via
     * #setPictureHorizontalOffset(), #setPictureVerticalOffset(),
     * #setPictureHorizontalScale(), #setPictureVerticalScale(),
     * #setTileAlignment(), #setTileFlipMode().
     *
     * @param picture If `null`, the previous picture will be removed.
     * @note This method does not check the fill type to be
     * FillType::PictureFill.
     */
    void setPicture(const QImage &picture);
    /**
     * @brief returns the picture that is used for filling.
     * @return Not-null QImage if the picture was set, null QImage otherwise.
     * @note This method does not check the fill type to be
     * FillType::PictureFill.
     */
    QImage picture() const;
    /**
     * @brief returns the mode of using the picture to fill the shape.
     * @return Valid optional value if the parameter is set, `nullopt`
     * otherwise.
     * @note If this parameter is not set, the picture will simply be truncated
     * to the shape's bounding box.
     *
     * @note This method does not check the fill type to be
     * FillType::PictureFill.
     */
    std::optional<PictureFillMode> pictureFillMode() const;
    /**
     * @brief sets the mode of using the picture to fill the shape.
     * @param mode If set to Stretch, the picture will be stretched to fill the
     * pictureStretchRect(). If set to Tile, the picture will be tiled.
     * @note If pictureFillMode() is not set, the picture will simply be
     * truncated to the shape's bounding box.
     *
     * You can use this method to clear the picture fill
     * mode: `setPictureFillMode(std::nullopt);`
     * @note This method does not check the fill type to be
     * FillType::PictureFill.
     */
    void setPictureFillMode(std::optional<PictureFillMode> mode);
    /**
     * @brief returns the rectangle portion of the picture used for the fill.
     *
     * Each edge of the source rectangle is defined by a percentage offset from
     * the corresponding edge of the bounding box. A positive percentage
     * specifies an inset, while a negative percentage specifies an outset. For
     * example, a left offset of 25% specifies that the left edge of the source
     * rectangle is located to the right of the bounding box's left edge by an
     * amount equal to 25% of the bounding box's width.
     *
     * @return valid optional if the parameter is set, `nullopt` otherwise.
     * @note This method does not check the fill type to be
     * FillType::PictureFill.
     */
    std::optional<RelativeRect> pictureSourceRect() const;
    /**
     * @brief sets the rectangle portion of the picture that will be used for
     * the fill.
     * @param rect RelativeRect that defines percentage offsets from the edges
     * of the shape's bounding box. A positive percentage specifies an inset,
     * while a negative percentage specifies an outset. For example, a left
     * offset of 25% specifies that the left edge of the source rectangle is
     * located to the right of the bounding box's left edge by an amount equal
     * to 25% of the bounding box's width.
     * @note This method does not check the fill type to be
     * FillType::PictureFill.
     */
    void setPictureSourceRect(const RelativeRect &rect);
    /**
     * @brief returns the DPI (dots per inch) used to calculate the size of the
     * picture.
     * @return valid optional if the parameter was set, `nullopt` otherwise.
     * @note This method does not check the fill type to be
     * FillType::PictureFill.
     */
    std::optional<int> pictureDpi() const;
    /**
     * @brief sets the DPI (dots per inch) used to calculate the size of the
     * picture.
     *
     * If not present or zero, the DPI in the picture is used.
     * @param dpi the DPI (dots per inch).
     * @note This method does not check the fill type to be
     * FillType::PictureFill.
     */
    void setPictureDpi(int dpi);
    /**
     * @brief returns the rectangle used to stretch the picture fill.
     * @return valid optional if the parameter is set, `nullopt` otherwise.
     * @note This method does not check the fill type to be
     * FillType::PictureFill.
     */
    std::optional<RelativeRect> pictureStretchRect() const;
    /**
     * @brief sets the rectangle used to stretch the picture fill.
     * @note This method also sets pictureFillMode to PictureFillMode::Stretch.
     * @param rect RelativeRect that defines percentage offsets from the edges
     * of the shape's bounding box. A positive percentage specifies an inset,
     * while a negative percentage specifies an outset. For example, a left
     * offset of 25% specifies that the left edge of the fill rectangle is
     * located to the right of the bounding box's left edge by an amount equal
     * to 25% of the bounding box's width.
     * @note This method does not check the fill type to be
     * FillType::PictureFill.
     */
    void setPictureStretchRect(const RelativeRect &rect);
    /**
     * @brief returns the horizontal offset of the picture to tile the shape's
     * bounding rectangle.
     * @return valid Coordinate if the parameter was set.
     * @note This method does not check the fill type to be
     * FillType::PictureFill.
     */
    Coordinate pictureHorizontalOffset() const;
    /**
     * @brief sets the horizontal offset of the picture to tile the shape's
     * bounding rectangle.
     * @param offset horizontal picture offset.
     * @note This method also sets pictureFillMode to PictureFillMode::Tile.
     * @note This method does not check the fill type to be
     * FillType::PictureFill.
     */
    void setPictureHorizontalOffset(const Coordinate &offset);
    /**
     * @brief returns the vertical offset of the picture to tile the shape's
     * bounding rectangle.
     * @return valid Coordinate if the parameter was set.
     * @note This method does not check the fill type to be
     * FillType::PictureFill.
     */
    Coordinate pictureVerticalOffset() const;
    /**
     * @brief sets the vertical offset of the picture to tile the shape's
     * bounding rectangle.
     * @param offset vertical picture offset.
     * @note This method also sets pictureFillMode to PictureFillMode::Tile.
     * @note This method does not check the fill type to be
     * FillType::PictureFill.
     */
    void setPictureVerticalOffset(const Coordinate &offset);
    /**
     * @brief returns the horizontal scaling of the picture.
     * @return Percentage of the scaling. Value of 100.0 means 100% scaling.
     * @note This method does not check the fill type to be
     * FillType::PictureFill.
     */
    std::optional<double> pictureHorizontalScale() const;
    /**
     * @brief sets the horizontal scaling of the picture.
     * @param scale Percentage of the scaling. Value of 100.0 means 100%
     * scaling.
     * @note This method also sets pictureFillMode to PictureFillMode::Tile.
     * @note This method does not check the fill type to be
     * FillType::PictureFill.
     */
    void setPictureHorizontalScale(double scale);
    /**
     * @brief returns the vertical scaling of the picture.
     * @return Percentage of the scaling. Value of 100.0 means 100% scaling.
     * @note This method does not check the fill type to be
     * FillType::PictureFill.
     */
    std::optional<double> pictureVerticalScale() const;
    /**
     * @brief sets the vertical scaling of the picture.
     * @param scale Percentage of the scaling. Value of 100.0 means 100%
     * scaling.
     * @note This method also sets pictureFillMode to PictureFillMode::Tile.
     * @note This method does not check the fill type to be
     * FillType::PictureFill.
     */
    void setPictureVerticalScale(double scale);
    /**
     * @brief returns where to align the first tile with respect to the shape.
     *
     * Alignment happens after the scaling, but before the additional offset.
     * @return valid optional if the parameter was set, `nullopt` otherwise.
     * @note This method does not check the fill type to be
     * FillType::PictureFill.
     */
    std::optional<Alignment> tileAlignment();
    /**
     * @brief sets where to align the first tile with respect to the
     * shape.
     *
     * Alignment happens after the scaling, but before the additional offset.
     * @param alignment picture alignment.
     * @note This method also sets pictureFillMode to PictureFillMode::Tile.
     * @note This method does not check the fill type to be
     * FillType::PictureFill.
     */
    void setTileAlignment(Alignment alignment);
    /**
     * @brief returns the compression quality that was used for a picture.
     * @return valid optional if the parameter was set, `nullopt` otherwise.
     * @note This parameter serves as an additional info.
     * @note This method does not check the fill type to be
     * FillType::PictureFill.
     */
    std::optional<FillFormat::PictureCompression> pictureCompression() const;
    /**
     * @brief sets the compression quality that was used for a picture.
     * @param compression picture compression.
     * @note This parameter serves as an additional info.
     * @note This method does not check the fill type to be
     * FillType::PictureFill.
     */
    void setPictureCompression(FillFormat::PictureCompression compression);
    /**
     * @brief returns the picture's alpha (opacity) (0 to 100.0 percent.)
     *
     * If pictureAlpha() is 30.0, the picture has 30% opacity.
     *
     * @return valid optional if the parameter was set, `nullopt` otherwise.
     * @note This method does not check the fill type to be
     * FillType::PictureFill.
     */
    std::optional<double> pictureAlpha() const;
    /**
     * @brief sets the picture's alpha (0 to 100.0 percent.)
     * @param alpha the picture opacity (0 means transparent, 100.0 means full
     * opacity.)
     * @note This method does not check the fill type to be
     * FillType::PictureFill.
     */
    void setPictureAlpha(double alpha);

    bool isValid() const;

    void write(QXmlStreamWriter &writer) const;
    void read(QXmlStreamReader &reader);

    bool operator==(const FillFormat &other) const;
    bool operator!=(const FillFormat &other) const;

    operator QVariant() const;

private:
    SERIALIZE_ENUM(FillType, {
        {FillType::NoFill, "noFill"},
        {FillType::SolidFill, "solidFill"},
        {FillType::GradientFill, "gradFill"},
        {FillType::PictureFill, "blipFill"},
        {FillType::PatternFill, "pattFill"},
        {FillType::GroupFill, "grpFill"},
    });
    SERIALIZE_ENUM(PathType, {
        {PathType::Shape, "shape"},
        {PathType::Circle, "circle"},
        {PathType::Rectangle, "rect"},
    });
    SERIALIZE_ENUM(TileFlipMode, {
        {TileFlipMode::None, "none"},
        {TileFlipMode::X, "x"},
        {TileFlipMode::Y, "y"},
        {TileFlipMode::XY, "xy"},
    });
    SERIALIZE_ENUM(PatternType, {
        {PatternType::Percent5, "pct5"},
        {PatternType::Percent10, "pct10"},
        {PatternType::Percent20, "pct20"},
        {PatternType::Percent25, "pct25"},
        {PatternType::Percent30, "pct30"},
        {PatternType::Percent40, "pct40"},
        {PatternType::Percent50, "pct50"},
        {PatternType::Percent60, "pct60"},
        {PatternType::Percent70, "pct70"},
        {PatternType::Percent75, "pct75"},
        {PatternType::Percent80, "pct80"},
        {PatternType::Percent90, "pct90"},
        {PatternType::Horizontal, "horz"},
        {PatternType::Vertical, "vert"},
        {PatternType::LightHorizontal, "ltHorz"},
        {PatternType::LightVertical, "ltVert"},
        {PatternType::DarkHorizontal, "dkHorz"},
        {PatternType::DarkVertical, "dkVert"},
        {PatternType::NarrowHorizontal, "narHorz"},
        {PatternType::NarrowVertical, "narVert"},
        {PatternType::DashedHorizontal, "dashHorz"},
        {PatternType::DashedVertical, "dashVert"},
        {PatternType::Cross, "cross"},
        {PatternType::DownwardDiagonal, "dnDiag"},
        {PatternType::UpwardDiagonal, "upDiag"},
        {PatternType::LightDownwardDiagonal, "ltDnDiag"},
        {PatternType::LightUpwardDiagonal, "ltUpDiag"},
        {PatternType::DarkDownwardDiagonal, "dkDnDiag"},
        {PatternType::DarkUpwardDiagonal, "dkUpDiag"},
        {PatternType::WideDownwardDiagonal, "wdDnDiag"},
        {PatternType::WideUpwardDiagonal, "wdUpDiag"},
        {PatternType::DashedDownwardDiagonal, "dashDnDiag"},
        {PatternType::DashedUpwardDiagonal, "dashUpDiag"},
        {PatternType::DiagonalCross, "diagCross"},
        {PatternType::SmallCheckerBoard, "smCheck"},
        {PatternType::LargeCheckerBoard, "lgCheck"},
        {PatternType::SmallGrid, "smGrid"},
        {PatternType::LargeGrid, "lgGrid"},
        {PatternType::DottedGrid, "dotGrid"},
        {PatternType::SmallConfetti, "smConfetti"},
        {PatternType::LargeConfetti, "lgConfetti"},
        {PatternType::HorizontalBrick, "horzBrick"},
        {PatternType::DiagonalBrick, "diagBrick"},
        {PatternType::SolidDiamond, "solidDmnd"},
        {PatternType::OpenDiamond, "openDmnd"},
        {PatternType::DottedDiamond, "dotDmnd"},
        {PatternType::Plaid, "plaid"},
        {PatternType::Sphere, "sphere"},
        {PatternType::Weave, "weave"},
        {PatternType::Divot, "divot"},
        {PatternType::Shingle, "shingle"},
        {PatternType::Wave, "wave"},
        {PatternType::Trellis, "trellis"},
        {PatternType::ZigZag, "zigZag"},
    });
    SERIALIZE_ENUM(PictureCompression, {
        {PictureCompression::None, "none"},
        {PictureCompression::Email, "email"},
        {PictureCompression::Screen, "screen"},
        {PictureCompression::Print, "print"},
        {PictureCompression::HqPrint, "hqprint"},
    });
    SERIALIZE_ENUM(Alignment,
    {
        {Alignment::AlignTopLeft, "tl"},
        {Alignment::AlignTop, "t"},
        {Alignment::AlignTopRight, "tr"},
        {Alignment::AlignRight, "r"},
        {Alignment::AlignCenter, "ctr"},
        {Alignment::AlignLeft, "l"},
        {Alignment::AlignBottomLeft, "bl"},
        {Alignment::AlignBottom, "b"},
        {Alignment::AlignBottomRight, "br"},
    });
    friend QDebug operator<<(QDebug, const FillFormat &f);
    QSharedDataPointer<FillFormatPrivate> d;
    friend class Effect;
    friend class Chart; //for the following 3 methods
    void setPictureID(int id);
    int registerBlip(Workbook *workbook);
    void loadBlip(Workbook *workbook, Relationships *relationships);

    void readNoFill(QXmlStreamReader &reader);
    void readSolidFill(QXmlStreamReader &reader);
    void readGradientFill(QXmlStreamReader &reader);
    void readPatternFill(QXmlStreamReader &reader);
    void readGroupFill(QXmlStreamReader &reader);
    void readPictureFill(QXmlStreamReader &reader);

    void writeNoFill(QXmlStreamWriter &writer) const;
    void writeSolidFill(QXmlStreamWriter &writer) const;
    void writeGradientFill(QXmlStreamWriter &writer) const;
    void writePatternFill(QXmlStreamWriter &writer) const;
    void writeGroupFill(QXmlStreamWriter &writer) const;
    void writePictureFill(QXmlStreamWriter &writer) const;

    void readGradientList(QXmlStreamReader &reader);
    void writeGradientList(QXmlStreamWriter &writer) const;
    void readBlip(QXmlStreamReader &reader);
};

QDebug operator<<(QDebug dbg, const FillFormat &f);

}

Q_DECLARE_TYPEINFO(QXlsx::FillFormat, Q_MOVABLE_TYPE);
Q_DECLARE_METATYPE(QXlsx::FillFormat)

#endif // XLSXFillformat_H
