#ifndef XLSXLAYOUT_H
#define XLSXLAYOUT_H

#include <QXmlStreamWriter>
#include <QXmlStreamReader>
#include <QSharedData>
#include <QVector3D>
#include <QRectF>
#include <QDebug>
#include <optional>

#include "xlsxglobal.h"
#include "xlsxutility_p.h"

namespace QXlsx {

class LayoutPrivate;

/**
 * @brief The class is used to set the exact position of a chart element.
 */
class QXLSX_EXPORT Layout
{
public:
    /**
     * @brief The LayoutTarget enum specifies  whether to layout the plot area
     * by its inside (not including axis and axis labels) or outside (including
     * axis and axis labels).
     */
    enum class LayoutTarget
    {
        Inner, /**< plot area is laid out by its inside (not including axis and axis labels). */
        Outer /**< plot area is laid out by its outside (including axis and axis labels). */
    };

    /**
     * @brief The LayoutMode enum specifies how to interpret the height and width
     * values of the layout:
     *
     * - either as the Right or Bottom of the chart element (Edge)
     * - or as the Width or Height of the chart element (Factor)

     * Default is LayoutMode::Factor
     */
    enum class LayoutMode
    {
        Edge, /**< rect().width() or rect().height() shall be interpreted as the Right or Bottom of the chart element */
        Factor /**< rect().width() or rect().height() shall be interpreted as the Width or Height of the chart element. */
    };

    Layout();
    Layout(const Layout &other);
    ~Layout();
    Layout &operator=(const Layout &other);

    std::optional<LayoutTarget> layoutTarget() const;
    void setLayoutTarget(LayoutTarget layoutTarget);

    std::optional<LayoutMode> xlayoutMode() const;
    void setXLayoutMode(LayoutMode layoutMode);

    std::optional<LayoutMode> ylayoutMode() const;
    void setYLayoutMode(LayoutMode layoutMode);

    std::optional<LayoutMode> wlayoutMode() const;
    void setWLayoutMode(LayoutMode layoutMode);

    std::optional<LayoutMode> hlayoutMode() const;
    void setHLayoutMode(LayoutMode layoutMode);

    QRectF rect() const;
    void setRect(QRectF rect);

    QPointF position() const;
    void setPosition(QPointF position);

    QSizeF size() const;
    void setSize(QSizeF size);

    double width() const;
    void setWidth(double width);

    double height() const;
    void setHeight(double height);

    double x() const;
    void setX(double x);

    double y() const;
    void setY(double y);

    bool operator ==(const Layout &other) const;
    bool operator !=(const Layout &other) const;

    operator QVariant() const;

    bool isValid() const;

    void read(QXmlStreamReader &reader);
    void write(QXmlStreamWriter &writer, const QString &name) const;
private:
    SERIALIZE_ENUM(LayoutTarget,
    {
        {LayoutTarget::Inner, "inner"},
        {LayoutTarget::Outer, "outer"},
    });
    SERIALIZE_ENUM(LayoutMode,
    {
        {LayoutMode::Factor, "factor"},
        {LayoutMode::Edge, "edge"},
    });
    friend QDebug operator<<(QDebug, const Layout &f);
    QSharedDataPointer<LayoutPrivate> d;
};

QDebug operator<<(QDebug dbg, const Layout &f);

}

Q_DECLARE_METATYPE(QXlsx::Layout)

#endif // XLSXLAYOUT_H
