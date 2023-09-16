#include "xlsxfillformat.h"
#include "xlsxutility_p.h"
#include "xlsxmediafile_p.h"
#include "xlsxworkbook.h"
#include <QDebug>
#include <QtMath>
#include <QBuffer>
#include <QDir>
#include "xlsxrelationships_p.h"
#include "xlsxchart.h"

namespace QXlsx {

class FillFormatPrivate: public QSharedData
{
public:
    FillFormatPrivate();
    FillFormatPrivate(const FillFormatPrivate &other);
    ~FillFormatPrivate();

    void parse(const QBrush &brush);
    QBrush toBrush() const;

    FillFormat::FillType type;
    Color color;

    //Gradient fill properties
    QMap<double, Color> gradientList;
    std::optional<RelativeRect> tileRect;
    std::optional<FillFormat::TileFlipMode> tileFlipMode;
    std::optional<bool> rotWithShape;
    //linear gradient
    Angle linearShadeAngle; //0..360
    std::optional<bool> linearShadeScaled;
    //path gradient
    std::optional<FillFormat::PathType> pathShadeType;
    std::optional<RelativeRect> pathShadeRect;



    //Pattern fill properties
    Color foregroundColor;
    Color backgroundColor;
    std::optional<FillFormat::PatternType> patternType;

    //Blip fill - limited support
    std::optional<int> dpi;
    QSharedPointer<MediaFile> blipImage;
    QString blipID; // internal
    std::optional<double> alpha;
    std::optional<RelativeRect> blipSrcRect;
    std::optional<RelativeRect> blipFillRect;
    Coordinate tx;
    Coordinate ty;
    std::optional<double> sx;
    std::optional<double> sy;
    std::optional<FillFormat::Alignment> tileAlignment;
    std::optional<FillFormat::PictureFillMode> fillMode;
    std::optional<FillFormat::PictureCompression> blipCompression;
    //TODO: all other blip effects (16 more)
    //TODO: create a separate class Blip

    bool operator==(const FillFormatPrivate &other) const;
};

FillFormatPrivate::FillFormatPrivate()
{

}

FillFormatPrivate::FillFormatPrivate(const FillFormatPrivate &other)
    : QSharedData(other), type(other.type), color(other.color),
      gradientList(other.gradientList), tileRect(other.tileRect),
      tileFlipMode(other.tileFlipMode), rotWithShape(other.rotWithShape),
      linearShadeAngle(other.linearShadeAngle), linearShadeScaled(other.linearShadeScaled),
      pathShadeType(other.pathShadeType), pathShadeRect(other.pathShadeRect),
      foregroundColor(other.foregroundColor), backgroundColor(other.backgroundColor),
      patternType(other.patternType), dpi{other.dpi},
      blipImage{other.blipImage}, //TODO: check copying of fills with picture fill
      blipID{other.blipID}, alpha{other.alpha}, blipSrcRect{other.blipSrcRect},
      blipFillRect{other.blipFillRect}, tx{other.tx}, ty{other.ty},
      sx{other.sx}, sy{other.sy}, tileAlignment{other.tileAlignment},
      fillMode{other.fillMode}, blipCompression{other.blipCompression}
{

}

FillFormatPrivate::~FillFormatPrivate()
{

}

bool FillFormatPrivate::operator==(const FillFormatPrivate &other) const
{
    if (type != other.type) return false;
    if (color != other.color) return false;
    if (gradientList != other.gradientList) return false;
    if (linearShadeAngle != other.linearShadeAngle) return false;
    if (linearShadeScaled != other.linearShadeScaled) return false;
    if (pathShadeType != other.pathShadeType) return false;
    if (pathShadeRect != other.pathShadeRect) return false;
    if (tileRect != other.tileRect) return false;
    if (tileFlipMode != other.tileFlipMode) return false;
    if (rotWithShape != other.rotWithShape) return false;
    if (foregroundColor != other.foregroundColor) return false;
    if (backgroundColor != other.backgroundColor) return false;
    if (patternType != other.patternType) return false;
    if (dpi != other.dpi) return false;
    if (blipImage != other.blipImage) return false;
    if (blipID != other.blipID) return false;
    if (alpha != other.alpha) return false;
    if (blipSrcRect != other.blipSrcRect) return false;
    if (blipFillRect != other.blipFillRect) return false;
    if (tx != other.tx) return false;
    if (ty != other.ty) return false;
    if (sx != other.sx) return false;
    if (sy != other.sy) return false;
    if (tileAlignment != other.tileAlignment) return false;
    if (fillMode != other.fillMode) return false;
    if (blipCompression != other.blipCompression) return false;

    return true;
}

void FillFormatPrivate::parse(const QBrush &brush)
{
    switch (brush.style()) {
        case Qt::NoBrush: type = FillFormat::FillType::NoFill; break;
        case Qt::SolidPattern: {
            type = FillFormat::FillType::SolidFill;
            color = brush.color(); break;
        }
        case Qt::Dense1Pattern: {
            type = FillFormat::FillType::PatternFill;
            patternType = FillFormat::PatternType::Percent80;
            foregroundColor = brush.color();
            break;
        }
        case Qt::Dense2Pattern: {
            type = FillFormat::FillType::PatternFill;
            patternType = FillFormat::PatternType::Percent75;
            foregroundColor = brush.color();
            break;
        }
        case Qt::Dense3Pattern: {
            type = FillFormat::FillType::PatternFill;
            patternType = FillFormat::PatternType::Percent60;
            foregroundColor = brush.color();
            break;
        }
        case Qt::Dense4Pattern: {
            type = FillFormat::FillType::PatternFill;
            patternType = FillFormat::PatternType::Percent50;
            foregroundColor = brush.color();
            break;
        }
        case Qt::Dense5Pattern: {
            type = FillFormat::FillType::PatternFill;
            patternType = FillFormat::PatternType::Percent30;
            foregroundColor = brush.color();
            break;
        }
        case Qt::Dense6Pattern: {
            type = FillFormat::FillType::PatternFill;
            patternType = FillFormat::PatternType::Percent20;
            foregroundColor = brush.color();
            break;
        }
        case Qt::Dense7Pattern: {
            type = FillFormat::FillType::PatternFill;
            patternType = FillFormat::PatternType::Percent5;
            foregroundColor = brush.color();
            break;
        }
        case Qt::HorPattern: {
            type = FillFormat::FillType::PatternFill;
            patternType = FillFormat::PatternType::Horizontal;
            foregroundColor = brush.color();
            break;
        }
        case Qt::VerPattern: {
            type = FillFormat::FillType::PatternFill;
            patternType = FillFormat::PatternType::Vertical;
            foregroundColor = brush.color();
            break;
        }
        case Qt::CrossPattern: {
            type = FillFormat::FillType::PatternFill;
            patternType = FillFormat::PatternType::Cross;
            foregroundColor = brush.color();
            break;
        }
        case Qt::BDiagPattern: {
            type = FillFormat::FillType::PatternFill;
            patternType = FillFormat::PatternType::UpwardDiagonal;
            foregroundColor = brush.color();
            break;
        }
        case Qt::FDiagPattern: {
            type = FillFormat::FillType::PatternFill;
            patternType = FillFormat::PatternType::DownwardDiagonal;
            foregroundColor = brush.color();
            break;
        }
        case Qt::DiagCrossPattern: {
            type = FillFormat::FillType::PatternFill;
            patternType = FillFormat::PatternType::OpenDiamond;
            foregroundColor = brush.color();
            break;
        }
        case Qt::LinearGradientPattern: {
            type = FillFormat::FillType::GradientFill;
            const auto gradient = brush.gradient();
            if (gradient->type() == QGradient::LinearGradient) {
                auto mode = gradient->coordinateMode();
                if (mode == QGradient::LogicalMode)
                    linearShadeScaled = false;
                if (mode == QGradient::ObjectMode)
                    linearShadeScaled = true;
                const auto stops = gradient->stops();
                for (auto st: stops)
                    gradientList.insert(st.first * 100.0, st.second);
                auto spread = gradient->spread();
                switch (spread) {
                    case QGradient::PadSpread: tileFlipMode = FillFormat::TileFlipMode::None; break;
                    case QGradient::RepeatSpread: tileFlipMode = FillFormat::TileFlipMode::XY; break;
                    case QGradient::ReflectSpread: tileFlipMode = FillFormat::TileFlipMode::XY; break;
                }

                if (auto linear = static_cast<const QLinearGradient*>(gradient)) {
                    auto start = linear->start();
                    auto end = linear->finalStop();

                    //vector from start to end will give us the gradient angle
                    QLineF l(start, end);
                    auto angle = - l.angle();
                    if (angle < 0) angle += 360;
                    linearShadeAngle = Angle(angle);

                    //there is no way to find out the size of a shape if mode is LogicalMode
                    //or the size of a paint device if mode is StretchToDeviceMode,
                    //so we unfortunately ignore the gradient coords in these modes
                    if (mode == QGradient::ObjectMode) {
                        auto left = 0.0, right = 0.0, top = 0.0, bottom = 0.0;
                        if (start.x() != end.x()) {
                            left = qMin(start.x(), end.x());
                            right = qMin(1.0 - end.x(), 1.0 - start.x());
                        }
                        if (start.y() != end.y()) {
                            top = qMin(start.y(), end.y());
                            bottom = qMin(1.0 - end.y(), 1.0 - start.y());
                        }
                        tileRect = RelativeRect(left * 100, top * 100, right * 100, bottom * 100);
                    }
                }
            }
            break;
        }
        case Qt::RadialGradientPattern: {
            type = FillFormat::FillType::GradientFill;
            const auto gradient = brush.gradient();
            if (gradient->type() == QGradient::RadialGradient) {
                pathShadeType = FillFormat::PathType::Circle;
                //properties common for all gradient types
                const auto stops = gradient->stops();
                for (auto st: stops)
                    gradientList.insert(st.first * 100.0, st.second);

                const auto mode = gradient->coordinateMode();

                const auto spread = gradient->spread();
                switch (spread) {
                    case QGradient::PadSpread: tileFlipMode = FillFormat::TileFlipMode::None; break;
                    case QGradient::RepeatSpread: tileFlipMode = FillFormat::TileFlipMode::XY; break;
                    case QGradient::ReflectSpread: tileFlipMode = FillFormat::TileFlipMode::XY; break;
                }
                //properties of radial gradient
                auto radial = static_cast<const QRadialGradient*>(gradient);
                if (radial) {
                    const auto center = radial->center();
                    const auto centerRadius = radial->centerRadius();
                    const auto focal = radial->focalPoint();
                    const auto focalRadius = radial->focalRadius();
                    //there is no way to find out the size of a shape if mode is LogicalMode
                    //or the size of a paint device if mode is StretchToDeviceMode,
                    //so we unfortunately ignore the radial gradient coords in these modes
                    if (mode == QGradient::ObjectMode) {
                        //center and centerRadius of gradient will define the tileRect
                        auto left = center.x() - centerRadius;
                        auto right = center.x() + centerRadius;
                        auto top = center.y() - centerRadius;
                        auto bottom = center.y() + centerRadius;
                        tileRect = RelativeRect(left * 100, top * 100, right * 100, bottom * 100);
                        //focal will define the pathShadeRect
                        //focalRadius is ignored
                        left = focal.x();
                        right = 1.0 - focal.x();
                        top = focal.y();
                        bottom = 1.0 - focal.y();
                        pathShadeRect = RelativeRect(left * 100, top * 100, right * 100, bottom * 100);
                        Q_UNUSED(focalRadius);
                    }
                }
            }
            break;
        }
        case Qt::ConicalGradientPattern: {
            //There is no way to implement conical gradient in xlsx, so use solid fill with 1st color
            const auto gradient = brush.gradient();
            type = FillFormat::FillType::SolidFill;
            const auto stops = gradient->stops();
            if (!stops.isEmpty())
                color = stops.first().second;
            break;
        }
        case Qt::TexturePattern:  {
            type = FillFormat::FillType::PictureFill;
            auto picture = brush.textureImage();
            if (picture.isNull()) {
                blipImage.reset();
                blipID.clear();
            }
            else {
                QByteArray ba;
                QBuffer buffer(&ba);
                buffer.open(QIODevice::WriteOnly);
                picture.save(&buffer, "PNG");
                blipImage = QSharedPointer<MediaFile>(new MediaFile(ba, QStringLiteral("png"), QStringLiteral("image/png")));
                //no possibility to register the new MediaFile in workbook, this will be done later.
            }
            break;
        }

    }
}

QBrush FillFormatPrivate::toBrush() const
{
    QBrush brush;
    //TODO:
    return brush;
}

FillFormat::FillFormat()
{

}

FillFormat::FillFormat(FillFormat::FillType type)
{
    d = new FillFormatPrivate;
    d->type = type;
}

FillFormat::FillFormat(const QBrush &brush)
{
    d = new FillFormatPrivate;
    d->parse(brush);
}

FillFormat::FillFormat(const QColor &color)
{
    d = new FillFormatPrivate;
    d->type = FillType::SolidFill;
    d->color = color;
}

FillFormat::FillFormat(const Color &color)
{
    d = new FillFormatPrivate;
    d->type = FillType::SolidFill;
    d->color = color;
}

FillFormat::FillFormat(const QColor &foreground, const QColor &background, PatternType pattern)
{
    d = new FillFormatPrivate;
    d->type = FillType::PatternFill;
    d->foregroundColor = foreground;
    d->backgroundColor = background;
    d->patternType = pattern;
}

FillFormat::FillFormat(const QMap<double, Color> &gradientList, PathType pathType)
{
    d = new FillFormatPrivate;
    d->type = FillType::GradientFill;
    setGradientList(gradientList);
    d->pathShadeType = pathType;
}

FillFormat::FillFormat(const QMap<double, Color> &gradientList)
{
    d = new FillFormatPrivate;
    d->type = FillType::GradientFill;
    setGradientList(gradientList);
}

FillFormat::FillFormat(const QImage &picture)
{
    d = new FillFormatPrivate;
    d->type = FillType::PictureFill;
    setPicture(picture);
}

FillFormat::FillFormat(const FillFormat &other) : d(other.d)
{

}

FillFormat &FillFormat::operator=(const FillFormat &rhs)
{
    if (*this != rhs)
        d = rhs.d;
    return *this;
}

FillFormat::~FillFormat()
{

}

void FillFormat::setBrush(const QBrush &brush)
{
    d = new FillFormatPrivate;
    d->parse(brush);
}

QBrush FillFormat::toBrush() const
{
    if (d) return d->toBrush();
    return {};
}

FillFormat::FillType FillFormat::type() const
{
    if (d) return d->type;
    return FillType::NoFill;
}

void FillFormat::setType(FillFormat::FillType type)
{
    if (!d) d = new FillFormatPrivate;
    d->type = type;
}

Color FillFormat::color() const
{
    if (d) return d->color;
    return Color();
}

void FillFormat::setColor(const Color &color)
{
    if (!d) {
        d = new FillFormatPrivate;
        d->type = FillType::SolidFill;
    }
    d->color = color;
}

void FillFormat::setColor(const QColor &color)
{
    if (!d) {
        d = new FillFormatPrivate;
        d->type = FillType::SolidFill;
    }
    d->color = color;
}

QMap<double, Color> FillFormat::gradientList() const
{
    if (d) return d->gradientList;
    return {};
}

void FillFormat::setGradientList(const QMap<double, Color> &list)
{
    if (!d) d = new FillFormatPrivate;
    d->gradientList = list;
}

void FillFormat::addGradientStop(double stop, const Color &color)
{
    if (!d) d = new FillFormatPrivate;
    d->gradientList.insert(stop, color);
}

Angle FillFormat::linearShadeAngle() const
{
    if (d)  return d->linearShadeAngle;
    return {};
}

void FillFormat::setLinearShadeAngle(Angle val)
{
    if (!d) d = new FillFormatPrivate;
    d->linearShadeAngle = val;
    d->pathShadeType.reset();
    d->pathShadeRect.reset();
}

std::optional<bool> FillFormat::linearShadeScaled() const
{
    if (d)  return d->linearShadeScaled;
    return {};
}

void FillFormat::setLinearShadeScaled(bool scaled)
{
    if (!d) d = new FillFormatPrivate;
    d->linearShadeScaled = scaled;
    d->pathShadeType.reset();
    d->pathShadeRect.reset();
}

std::optional<FillFormat::PathType> FillFormat::pathType() const
{
    if (d)  return d->pathShadeType;
    return {};
}

void FillFormat::setPathType(FillFormat::PathType pathType)
{
    if (!d) d = new FillFormatPrivate;
    d->pathShadeType = pathType;
    d->linearShadeAngle = {};
    d->linearShadeScaled.reset();
}

std::optional<RelativeRect> FillFormat::pathRect() const
{
    if (d)  return d->pathShadeRect;
    return {};
}

void FillFormat::setPathRect(RelativeRect rect)
{
    if (!d) d = new FillFormatPrivate;
    d->pathShadeRect = rect;
    d->linearShadeAngle = {};
    d->linearShadeScaled.reset();
}

std::optional<RelativeRect> FillFormat::tileRect() const
{
    if (d)  return d->tileRect;
    return {};
}

void FillFormat::setTileRect(RelativeRect rect)
{
    if (!d) d = new FillFormatPrivate;
    d->tileRect = rect;
}

std::optional<FillFormat::TileFlipMode> FillFormat::tileFlipMode() const
{
    if (d)  return d->tileFlipMode;
    return {};
}

void FillFormat::setTileFlipMode(FillFormat::TileFlipMode tileFlipMode)
{
    if (!d) d = new FillFormatPrivate;
    d->tileFlipMode = tileFlipMode;
}

std::optional<bool> FillFormat::rotateWithShape() const
{
    if (d)  return d->rotWithShape;
    return {};
}

void FillFormat::setRotateWithShape(bool val)
{
    if (!d) d = new FillFormatPrivate;
    d->rotWithShape = val;
}

Color FillFormat::foregroundColor() const
{
    if (d) return d->foregroundColor;
    return Color();
}

void FillFormat::setForegroundColor(const Color &color)
{
    if (!d) d = new FillFormatPrivate;
    d->foregroundColor = color;
}

Color FillFormat::backgroundColor() const
{
    if (d) return d->backgroundColor;
    return Color();
}

void FillFormat::setBackgroundColor(const Color &color)
{
    if (!d) d = new FillFormatPrivate;
    d->backgroundColor = color;
}

std::optional<FillFormat::PatternType> FillFormat::patternType()
{
    if (d) return d->patternType;
    return {};
}

void FillFormat::setPatternType(FillFormat::PatternType patternType)
{
    if (!d) d = new FillFormatPrivate;
    d->patternType = patternType;
}

void FillFormat::setPicture(const QImage &picture)
{
    if (!d) d = new FillFormatPrivate;
    if (picture.isNull()) {
        d->blipImage.reset();
        d->blipID.clear();
    }
    else {
        QByteArray ba;
        QBuffer buffer(&ba);
        buffer.open(QIODevice::WriteOnly);
        picture.save(&buffer, "PNG");

        d->blipImage = QSharedPointer<MediaFile>(new MediaFile(ba, QStringLiteral("png"), QStringLiteral("image/png")));
        //no possibility to register the new MediaFile in workbook, this will be done later.
    }
}

QImage FillFormat::picture() const
{
    if (d && d->blipImage) {
        QImage pic;
        pic.loadFromData(d->blipImage->contents());
        return pic;
    }
    return {};
}

std::optional<FillFormat::PictureFillMode> FillFormat::pictureFillMode() const
{
    if (d) return d->fillMode;
    return std::nullopt;
}
void FillFormat::setPictureFillMode(std::optional<PictureFillMode> mode)
{
    if (!d) d = new FillFormatPrivate;
    d->fillMode = mode;
}

std::optional<RelativeRect> FillFormat::pictureSourceRect() const
{
    if (d) return d->blipSrcRect;
    return {};
}

void FillFormat::setPictureSourceRect(const RelativeRect &rect)
{
    if (!d) d = new FillFormatPrivate;
    d->blipSrcRect = rect;
}

std::optional<int> FillFormat::pictureDpi() const
{
    if (d) return d->dpi;
    return {};
}

void FillFormat::setPictureDpi(int dpi)
{
    if (!d) d = new FillFormatPrivate;
    d->dpi = dpi;
}

std::optional<RelativeRect> FillFormat::pictureStretchRect() const
{
    if (d) return d->blipFillRect;
    return {};
}
void FillFormat::setPictureStretchRect(const RelativeRect &rect)
{
    if (!d) d = new FillFormatPrivate;
    d->blipFillRect = rect;
    d->fillMode = PictureFillMode::Stretch;
}

Coordinate FillFormat::pictureHorizontalOffset() const
{
    if (d) return d->tx;
    return {};
}
void FillFormat::setPictureHorizontalOffset(const Coordinate &offset)
{
    if (!d) d = new FillFormatPrivate;
    d->tx = offset;
    d->fillMode = PictureFillMode::Tile;
}

Coordinate FillFormat::pictureVerticalOffset() const
{
    if (d) return d->ty;
    return {};
}
void FillFormat::setPictureVerticalOffset(const Coordinate &offset)
{
    if (!d) d = new FillFormatPrivate;
    d->ty = offset;
    d->fillMode = PictureFillMode::Tile;
}

std::optional<double> FillFormat::pictureHorizontalScale() const
{
    if (d) return d->sx;
    return {};
}
void FillFormat::setPictureHorizontalScale(double scale)
{
    if (!d) d = new FillFormatPrivate;
    d->sx = scale;
    d->fillMode = PictureFillMode::Tile;
}

std::optional<double> FillFormat::pictureVerticalScale() const
{
    if (d) return d->sy;
    return {};
}
void FillFormat::setPictureVerticalScale(double scale)
{
    if (!d) d = new FillFormatPrivate;
    d->sy = scale;
    d->fillMode = PictureFillMode::Tile;
}

std::optional<FillFormat::Alignment> FillFormat::tileAlignment()
{
    if (d) return d->tileAlignment;
    return {};
}
void FillFormat::setTileAlignment(FillFormat::Alignment alignment)
{
    if (!d) d = new FillFormatPrivate;
    d->tileAlignment = alignment;
    d->fillMode = PictureFillMode::Tile;
}

std::optional<FillFormat::PictureCompression> FillFormat::pictureCompression() const
{
    if (d) return d->blipCompression;
    return {};
}
void FillFormat::setPictureCompression(FillFormat::PictureCompression compression)
{
    if (!d) d = new FillFormatPrivate;
    d->blipCompression = compression;
}

std::optional<double> FillFormat::pictureAlpha() const
{
    if (d) return d->alpha;
    return {};
}
void FillFormat::setPictureAlpha(double alpha)
{
    if (!d) d = new FillFormatPrivate;
    d->alpha = alpha;
}

bool FillFormat::isValid() const
{
    if (d) return true;
    return false;
}

void FillFormat::setPictureID(int id)
{
    if (!d) d = new FillFormatPrivate;
    d->blipID = QString("rId%1").arg(id);
}

int FillFormat::registerBlip(Workbook *workbook)
{
    if (!d || !d->blipImage) return -1;
    if (d->blipImage->contents().isEmpty()) return -1;

    workbook->addMediaFile(d->blipImage);
    return d->blipImage->index();
}

void FillFormat::loadBlip(Workbook *workbook, Relationships *relationships)
{
    //After reading blip fill each blip has blipId that is somewhere registered in relationships.
    //We need to find this relation and find the corresponding MediaFile in
    //worksbooks mediaFiles.

    if (!d || d->blipID.isEmpty()) return;

    if (!d->blipImage) {
        auto r = relationships->getRelationshipById(d->blipID);
        auto mediaFiles = workbook->mediaFiles();

        QString name = r.target;

        const auto parts = splitPath(workbook->chartFiles().first().lock()->filePath());
        QString path = QDir::cleanPath(parts.first() + QLatin1String("/") + name);

        bool exist = false;
        for (const auto &mf : mediaFiles) {
            if (auto media = mf.lock(); media->fileName() == path) {
                //already exist
                exist = true;
                d->blipImage = media;
                break;
            }
        }
        if (!exist) {
            d->blipImage = QSharedPointer<MediaFile>(new MediaFile(path));
            workbook->addMediaFile(d->blipImage, true);
        }
    }
    else qDebug() << "Warning: blip image is already loaded.";
}

void FillFormat::write(QXmlStreamWriter &writer) const
{
    switch (d->type) {
        case FillType::NoFill : writeNoFill(writer); break;
        case FillType::SolidFill : writeSolidFill(writer); break;
        case FillType::GradientFill : writeGradientFill(writer); break;
        case FillType::PictureFill : writePictureFill(writer); break;
        case FillType::PatternFill : writePatternFill(writer); break;
        case FillType::GroupFill : writeGroupFill(writer); break;
    }
}

void FillFormat::read(QXmlStreamReader &reader)
{
    const QString &name = reader.name().toString();
    FillType t;
    fromString(name, t);

    if (!d) {
        d = new FillFormatPrivate;
        d->type = t;
    }

    switch (d->type) {
        case FillType::NoFill: break; //NO-OP
        case FillType::SolidFill : readSolidFill(reader); break;
        case FillType::GradientFill : readGradientFill(reader); break;
        case FillType::PictureFill : readPictureFill(reader); break;
        case FillType::PatternFill : readPatternFill(reader); break;
        case FillType::GroupFill : readGroupFill(reader); break;
        default: break;
    }
}

bool FillFormat::operator==(const FillFormat &other) const
{
    if (d == other.d) return true;
    if (!d || !other.d) return false;
    return *this->d.constData() == *other.d.constData();
}

bool FillFormat::operator!=(const FillFormat &other) const
{
    return !(operator==(other));
}

FillFormat::operator QVariant() const
{
    const auto& cref
#if QT_VERSION >= 0x060000 // Qt 6.0 or over
        = QMetaType::fromType<FillFormat>();
#else
        = qMetaTypeId<FillFormat>() ;
#endif
    return QVariant(cref, this);
}

void FillFormat::readNoFill(QXmlStreamReader &reader)
{
    Q_ASSERT(reader.name() == QLatin1String("noFill"));

    //NO_OP
}

void FillFormat::readSolidFill(QXmlStreamReader &reader)
{
    Q_ASSERT(reader.name() == QLatin1String("solidFill"));

    while (!reader.atEnd()) {
        auto token = reader.readNext();
        if (token == QXmlStreamReader::StartElement)
            d->color.read(reader);
        else if (token == QXmlStreamReader::EndElement && reader.name() == QLatin1String("solidFill"))
            break;
    }
}

void FillFormat::readGradientFill(QXmlStreamReader &reader)
{
    Q_ASSERT(reader.name() == QLatin1String("gradFill"));

    const auto &attr = reader.attributes();
    if (attr.hasAttribute(QLatin1String("flip"))) {
        TileFlipMode t;
        fromString(attr.value(QLatin1String("flip")).toString(), t);
        d->tileFlipMode = t;
    }
    parseAttributeBool(attr, QLatin1String("rotWithShape"), d->rotWithShape);

//    <xsd:complexType name="CT_GradientFillProperties">
//        <xsd:sequence>
//          <xsd:element name="gsLst" type="CT_GradientStopList" minOccurs="0" maxOccurs="1"/>
//          <xsd:group ref="EG_ShadeProperties" minOccurs="0" maxOccurs="1"/>
//          <xsd:element name="tileRect" type="CT_RelativeRect" minOccurs="0" maxOccurs="1"/>
//        </xsd:sequence>
//        <xsd:attribute name="flip" type="ST_TileFlipMode" use="optional" default="none"/>
//        <xsd:attribute name="rotWithShape" type="xsd:boolean" use="optional"/>
//      </xsd:complexType>

    while (!reader.atEnd()) {
        auto token = reader.readNext();
        if (token == QXmlStreamReader::StartElement) {
            if (reader.name() == QLatin1String("gsLst")) {
                readGradientList(reader);
            }
            else if (reader.name() == QLatin1String("lin")) {
                //linear shade properties
                const auto &attr = reader.attributes();
                parseAttribute(attr, QLatin1String("ang"), d->linearShadeAngle);
                parseAttributeBool(attr, QLatin1String("scaled"), d->linearShadeScaled);
            }
            else if (reader.name() == QLatin1String("path")) {
                //path shade properties
                const auto &attr = reader.attributes();
                if (attr.hasAttribute(QLatin1String("path"))) {
                    auto s = attr.value(QLatin1String("path"));
                    if (s == QLatin1String("shape")) d->pathShadeType = PathType::Shape;
                    else if (s == QLatin1String("circle")) d->pathShadeType = PathType::Circle;
                    else if (s == QLatin1String("rect")) d->pathShadeType = PathType::Rectangle;
                }

                reader.readNextStartElement();
                if (reader.name() == QLatin1String("fillToRect")) {
                    RelativeRect r;
                    r.read(reader);
                    d->pathShadeRect = r;
                }
            }
            else if (reader.name() == QLatin1String("tileRect")) {
                RelativeRect r;
                r.read(reader);
                d->tileRect = r;
            }
            else reader.skipCurrentElement();
        }
        else if (token == QXmlStreamReader::EndElement && reader.name() == QLatin1String("gradFill"))
            break;
    }
}

void FillFormat::readBlip(QXmlStreamReader &reader)
{
    const auto &name = reader.name();
    const auto &a = reader.attributes();
    parseAttributeString(a, QLatin1String("r:embed"), d->blipID);
    if (a.hasAttribute(QLatin1String("cstate"))) {
        PictureCompression c; fromString(a.value(QLatin1String("cstate")).toString(), c);
        d->blipCompression = c;
    }

    while (!reader.atEnd()) {
        const auto token = reader.readNext();
        if (token == QXmlStreamReader::StartElement) {
            if (reader.name() == QLatin1String("alphaModFix")) {
                parseAttributePercent(reader.attributes(), QLatin1String("amt"), d->alpha);
            }
            else {
                reader.skipCurrentElement();
                //TODO: all other blip effects (16 more)
            }
        }
        else if (token == QXmlStreamReader::EndElement && reader.name() == name) break;
    }
}

void FillFormat::readPictureFill(QXmlStreamReader &reader)
{
    const auto &name = reader.name();
    parseAttributeInt(reader.attributes(), QLatin1String("dpi"), d->dpi);
    parseAttributeBool(reader.attributes(), QLatin1String("rotWithShape"), d->rotWithShape);

    while (!reader.atEnd()) {
        const auto token = reader.readNext();
        if (token == QXmlStreamReader::StartElement) {
            const auto &a = reader.attributes();
            if (reader.name() == QLatin1String("blip")) {
                readBlip(reader);
            }
            else if (reader.name() == QLatin1String("srcRect")) {
                RelativeRect r; r.read(reader);
                d->blipSrcRect = r;
            }
            else if (reader.name() == QLatin1String("tile")) {
                d->fillMode = PictureFillMode::Tile;
                if (a.hasAttribute(QLatin1String("tx"))) d->tx = Coordinate::create(a.value(QLatin1String("tx")));
                if (a.hasAttribute(QLatin1String("ty"))) d->ty = Coordinate::create(a.value(QLatin1String("ty")));
                parseAttributePercent(a, QLatin1String("sx"), d->sx);
                parseAttributePercent(a, QLatin1String("sy"), d->sy);
                if (a.hasAttribute(QLatin1String("flip"))) {
                    TileFlipMode t; fromString(a.value(QLatin1String("flip")).toString(), t);
                    d->tileFlipMode = t;
                }
                if (a.hasAttribute(QLatin1String("algn"))) {
                    Alignment t; fromString(a.value(QLatin1String("algn")).toString(), t);
                    d->tileAlignment = t;
                }
            }
            else if (reader.name() == QLatin1String("stretch")) {
                d->fillMode = PictureFillMode::Stretch;
                reader.readNextStartElement();
                if (reader.name() == QLatin1String("fillRect")) {
                    RelativeRect r; r.read(reader);
                    d->blipFillRect = r;
                }
            }
        }
        else if (token == QXmlStreamReader::EndElement && reader.name() == name) break;
    }
}

void FillFormat::writeNoFill(QXmlStreamWriter &writer) const
{
    writer.writeEmptyElement("a:noFill");
}

void FillFormat::writeSolidFill(QXmlStreamWriter &writer) const
{
    writer.writeStartElement("a:solidFill");
    d->color.write(writer);
    writer.writeEndElement();
}

void FillFormat::writeGradientFill(QXmlStreamWriter &writer) const
{
    writer.writeStartElement(QLatin1String("a:gradFill"));
    if (d->tileFlipMode.has_value()) {
        QString s;
        toString(d->tileFlipMode.value(), s);
        writer.writeAttribute("flip", s);
    }
    if (d->rotWithShape.has_value()) writer.writeAttribute("rotWithShape", d->rotWithShape.value() ? "true" : "false");
    writeGradientList(writer);

    if (d->linearShadeAngle.isValid() || d->linearShadeScaled.has_value()) {
        writer.writeEmptyElement("a:lin");
        if (d->linearShadeAngle.isValid())
            writer.writeAttribute("ang", d->linearShadeAngle.toString());
        if (d->linearShadeScaled.has_value())
            writer.writeAttribute("scaled", d->linearShadeScaled.value() ? "true" : "false");
    }
    else {
        if (d->pathShadeType.has_value() || d->pathShadeRect.has_value()) {
            writer.writeStartElement("a:path");
            if (d->pathShadeType.has_value())
                switch (d->pathShadeType.value()) {
                    case PathType::Shape: writer.writeAttribute("path", "shape"); break;
                    case PathType::Circle: writer.writeAttribute("path", "circle"); break;
                    case PathType::Rectangle: writer.writeAttribute("path", "rect"); break;
                }
            if (d->pathShadeRect.has_value())
                d->pathShadeRect->write(writer, "a:fillToRect");
            writer.writeEndElement();
        }
    }
    if (d->tileRect.has_value() && d->tileRect->isValid())
        d->tileRect->write(writer, "a:tileRect");

    writer.writeEndElement();
}

void FillFormat::readPatternFill(QXmlStreamReader &reader)
{
    const auto &name = reader.name();
    if (reader.attributes().hasAttribute(QLatin1String("prst"))) {
        PatternType t;
        fromString(reader.attributes().value(QLatin1String("prst")).toString(), t);
        d->patternType = t;
    }
    while (!reader.atEnd()) {
        auto token = reader.readNext();
        if (token == QXmlStreamReader::StartElement) {
            if (reader.name() == QLatin1String("fgClr")) {
                d->foregroundColor.read(reader);
            }
            else if (reader.name() == QLatin1String("bgClr")) {
                d->backgroundColor.read(reader);
            }
            else reader.skipCurrentElement();
        }
        else if (token == QXmlStreamReader::EndElement && reader.name() == name)
            break;
    }
}

void FillFormat::writePatternFill(QXmlStreamWriter &writer) const
{
    if (!d->backgroundColor.isValid() && !d->foregroundColor.isValid() && !d->patternType.has_value()) {
        writer.writeEmptyElement(QLatin1String("a:pattFill"));
        return;
    }

    writer.writeStartElement(QLatin1String("a:pattFill"));
    if (d->patternType.has_value()) {
        QString s;
        toString(d->patternType.value(), s);
        writer.writeAttribute(QLatin1String("prst"), s);
    }
    if (d->backgroundColor.isValid()) {
        writer.writeStartElement(QLatin1String("a:bgClr"));
        d->backgroundColor.write(writer);
        writer.writeEndElement();
    }
    if (d->foregroundColor.isValid()) {
        writer.writeStartElement(QLatin1String("a:fgClr"));
        d->foregroundColor.write(writer);
        writer.writeEndElement();
    }
    writer.writeEndElement();
}

void FillFormat::readGroupFill(QXmlStreamReader &reader)
{
    Q_UNUSED(reader);
    //no-op
}

void FillFormat::writeGroupFill(QXmlStreamWriter &writer) const
{
    writer.writeEmptyElement(QLatin1String("a:grpFill"));
}

void FillFormat::writePictureFill(QXmlStreamWriter &writer) const
{
    if (!d->blipImage || d->blipID.isEmpty()) return;
    writer.writeStartElement(QLatin1String("a:blipFill"));
    writeAttribute(writer, QLatin1String("rotWithShape"), d->rotWithShape);
    writeAttribute(writer, QLatin1String("dpi"), d->dpi);
    writer.writeStartElement(QLatin1String("a:blip"));
    writer.writeAttribute(QLatin1String("xmlns:r"), QLatin1String("http://schemas.openxmlformats.org/officeDocument/2006/relationships"));
    writer.writeAttribute(QLatin1String("r:embed"), d->blipID);
    if (d->blipCompression.has_value()) {
        QString s; toString(d->blipCompression.value(), s);
        writer.writeAttribute(QLatin1String("cstate"), s);
    }
    if (d->alpha.has_value()) {
        writer.writeEmptyElement(QLatin1String("a:alphaModFix"));
        writeAttributePercent(writer, QLatin1String("amt"), d->alpha.value());
    }
    writer.writeEndElement(); //a:blip
    if (d->blipSrcRect.has_value()) d->blipSrcRect.value().write(writer, QLatin1String("a:srcRect"));
    if (d->fillMode.has_value()) {
        switch (d->fillMode.value()) {
            case PictureFillMode::Stretch: {
                writer.writeStartElement(QLatin1String("a:stretch"));
                if (d->blipFillRect.has_value()) {
                    d->blipFillRect.value().write(writer, QLatin1String("a:fillRect"));
                }
                writer.writeEndElement(); //a:stretch
                break;
            }
            case PictureFillMode::Tile: {
                writer.writeStartElement(QLatin1String("a:tile"));
                if (d->tx.isValid()) writeAttribute(writer, QLatin1String("tx"), d->tx.toString());
                if (d->ty.isValid()) writeAttribute(writer, QLatin1String("ty"), d->ty.toString());
                writeAttributePercent(writer, QLatin1String("sx"), d->sx);
                writeAttributePercent(writer, QLatin1String("sy"), d->sy);
                if (d->tileFlipMode.has_value()) {
                    QString s; toString(d->tileFlipMode.value(), s);
                    writeAttribute(writer, QLatin1String("flip"), s);
                }
                if (d->tileAlignment.has_value()) {
                    QString s; toString(d->tileAlignment.value(), s);
                    writeAttribute(writer, QLatin1String("algn"), s);
                }
                writer.writeEndElement(); //a:tile
                break;
            }
        }
    }

    writer.writeEndElement(); //a:blipFill
}

void FillFormat::readGradientList(QXmlStreamReader &reader)
{
    while (!reader.atEnd()) {
        auto token = reader.readNext();
        if (token == QXmlStreamReader::StartElement) {
            if (reader.name() == QLatin1String("gs")) {
                double pos = fromST_Percent(reader.attributes().value("pos"));
                reader.readNextStartElement();
                Color col;
                col.read(reader);
                d->gradientList.insert(pos, col);
            }
            else reader.skipCurrentElement();
        }
        else if (token == QXmlStreamReader::EndElement && reader.name() == QLatin1String("gsLst"))
            break;
    }
}

void FillFormat::writeGradientList(QXmlStreamWriter &writer) const
{
    if (d->gradientList.isEmpty()) return;
    writer.writeStartElement(QLatin1String("a:gsLst"));
    for (auto i = d->gradientList.constBegin(); i!= d->gradientList.constEnd(); ++i) {
        writer.writeStartElement(QLatin1String("a:gs"));
        writeAttributePercent(writer, QLatin1String("pos"), i.key());
        i.value().write(writer);
        writer.writeEndElement();
    }
    writer.writeEndElement();
}

#ifndef QT_NO_DEBUG_STREAM
QDebug operator<<(QDebug dbg, const FillFormat &f)
{
    //TODO: rest of parameters
    QDebugStateSaver saver(dbg);
    dbg.setAutoInsertSpaces(false);
    dbg << "QXlsx::FillFormat(" ;
    dbg << "type: "<< static_cast<int>(f.d->type) << ", ";
    if (f.d->color.isValid()) dbg << "color: " << f.d->color << ", ";
    dbg << "gradientList: "<< f.d->gradientList << ", ";
    if (f.d->linearShadeAngle.isValid()) dbg << "linearShadeAngle: " << f.d->linearShadeAngle.toString() << ", ";
    if (f.d->linearShadeScaled.has_value()) dbg << "linearShadeScaled: " << f.d->linearShadeScaled.value() << ", ";
    if (f.d->pathShadeType.has_value()) dbg << "pathShadeType: " << static_cast<int>(f.d->pathShadeType.value()) << ", ";
    if (f.d->pathShadeRect.has_value()) dbg << "pathShadeRect: " << f.d->pathShadeRect.value() << ", ";
    if (f.d->tileRect.has_value()) dbg << "tileRect: " << f.d->tileRect.value() << ", ";
    if (f.d->tileFlipMode.has_value())
        dbg << "tileFlipMode: " << static_cast<int>(f.d->tileFlipMode.value()) << ", ";
    if (f.d->rotWithShape.has_value()) dbg << "rotWithShape: " << f.d->rotWithShape.value() << ", ";
    if (f.d->foregroundColor.isValid()) dbg << "foregroundColor: " << f.d->foregroundColor << ", ";
    if (f.d->backgroundColor.isValid()) dbg << "backgroundColor: " << f.d->backgroundColor << ", ";
    if (f.d->patternType.has_value())
        dbg << "patternType: " << static_cast<int>(f.d->patternType.value());

    return dbg;
}
#endif

}




