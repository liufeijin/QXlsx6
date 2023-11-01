#include "xlsxlayout.h"

namespace QXlsx {

class LayoutPrivate: public QSharedData
{
public:
    LayoutPrivate();
    LayoutPrivate(const LayoutPrivate &other);
    ~LayoutPrivate();
    bool operator ==(const LayoutPrivate &other) const;

    std::optional<Layout::LayoutTarget> layoutTarget;
    std::optional<Layout::LayoutMode> xMode;
    std::optional<Layout::LayoutMode> yMode;
    std::optional<Layout::LayoutMode> wMode;
    std::optional<Layout::LayoutMode> hMode;

    std::optional<double> x;
    std::optional<double> y;
    std::optional<double> w;
    std::optional<double> h;
};

Layout::Layout()
{

}

Layout::Layout(const Layout &other) : d(other.d)
{

}

Layout::~Layout()
{

}

Layout &Layout::operator=(const Layout &other)
{
    if (*this != other)
        d = other.d;
    return *this;
}

std::optional<Layout::LayoutTarget> Layout::layoutTarget() const
{
    if (d) return d->layoutTarget;
    return {};
}

void Layout::setLayoutTarget(LayoutTarget layoutTarget)
{
    if (!d) d = new LayoutPrivate;
    d->layoutTarget = layoutTarget;
}

std::optional<Layout::LayoutMode> Layout::xlayoutMode() const
{
    if (d) return d->xMode;
    return {};
}

void Layout::setXLayoutMode(LayoutMode layoutMode)
{
    if (!d) d = new LayoutPrivate;
    d->xMode = layoutMode;
}

std::optional<Layout::LayoutMode> Layout::ylayoutMode() const
{
    if (d) return d->yMode;
    return {};
}

void Layout::setYLayoutMode(LayoutMode layoutMode)
{
    if (!d) d = new LayoutPrivate;
    d->yMode = layoutMode;
}

std::optional<Layout::LayoutMode> Layout::wlayoutMode() const
{
    if (d) return d->wMode;
    return {};
}

void Layout::setWLayoutMode(LayoutMode layoutMode)
{
    if (!d) d = new LayoutPrivate;
    d->wMode = layoutMode;
}

std::optional<Layout::LayoutMode> Layout::hlayoutMode() const
{
    if (d) return d->hMode;
    return {};
}

void Layout::setHLayoutMode(LayoutMode layoutMode)
{
    if (!d) d = new LayoutPrivate;
    d->hMode = layoutMode;
}

QRectF Layout::rect() const
{
    return QRectF(position(), size());
}

void Layout::setRect(QRectF rect)
{
    setPosition(rect.topLeft());
    setSize(rect.size());
}

QPointF Layout::position() const
{
    return {x(), y()};
}

void Layout::setPosition(QPointF position)
{
    setX(position.x());
    setY(position.y());
}

QSizeF Layout::size() const
{
    return {width(), height()};
}

void Layout::setSize(QSizeF size)
{
    setHeight(size.height());
    setWidth(size.width());
}

std::optional<double> Layout::width() const
{
    if (d) return d->w;
    return {};
}

void Layout::setWidth(double width)
{
    if (!d) d = new LayoutPrivate;
    d->wMode = LayoutMode::Factor;
    d->w = width;
}

std::optional<double> Layout::height() const
{
    if (d) return d->h;
    return {};
}

void Layout::setHeight(double height)
{
    if (!d) d = new LayoutPrivate;
    d->hMode = LayoutMode::Factor;
    d->h = height;
}

std::optional<double> Layout::x() const
{
    if (d) return d->x;
    return {};
}

void Layout::setX(double x)
{
    if (!d) d = new LayoutPrivate;
    d->xMode = LayoutMode::Edge;
    d->x = x;
}

std::optional<double> Layout::y() const
{
    if (d) return d->y;
    return {};
}

void Layout::setY(double y)
{
    if (!d) d = new LayoutPrivate;
    d->yMode = LayoutMode::Edge;
    d->y = y;
}



bool Layout::operator ==(const Layout &other) const
{
    if (d == other.d) return true;
    if (!d || !other.d) return false;
    return *this->d.constData() == *other.d.constData();
}

bool Layout::operator !=(const Layout &other) const
{
    return !(operator==(other));
}

Layout::operator QVariant() const
{
    const auto& cref
#if QT_VERSION >= 0x060000 // Qt 6.0 or over
        = QMetaType::fromType<Layout>();
#else
        = qMetaTypeId<Layout>() ;
#endif
    return QVariant(cref, this);
}

bool Layout::isValid() const
{
    if (d) return true;
    return false;
}

void Layout::read(QXmlStreamReader &reader)
{
    const auto &name = reader.name();
    if (!reader.readNextStartElement()) return;

    if (!d) d = new LayoutPrivate;

    while (!reader.atEnd()) {
        auto token = reader.readNext();
        if (token == QXmlStreamReader::StartElement) {
            const auto &val = reader.attributes().value("val");
            if (reader.name() == QLatin1String("layoutTarget")) {
                if (val == QLatin1String("inner")) d->layoutTarget = LayoutTarget::Inner;
                if (val == QLatin1String("outer")) d->layoutTarget = LayoutTarget::Outer;
            }
            else if (reader.name() == QLatin1String("xMode")) {
                if (val == QLatin1String("edge")) d->xMode = LayoutMode::Edge;
                if (val == QLatin1String("factor")) d->xMode = LayoutMode::Factor;
            }
            else if (reader.name() == QLatin1String("yMode")) {
                if (val == QLatin1String("edge")) d->yMode = LayoutMode::Edge;
                if (val == QLatin1String("factor")) d->yMode = LayoutMode::Factor;
            }
            else if (reader.name() == QLatin1String("wMode")) {
                if (val == QLatin1String("edge")) d->wMode = LayoutMode::Edge;
                if (val == QLatin1String("factor")) d->wMode = LayoutMode::Factor;
            }
            else if (reader.name() == QLatin1String("hMode")) {
                if (val == QLatin1String("edge")) d->hMode = LayoutMode::Edge;
                if (val == QLatin1String("factor")) d->hMode = LayoutMode::Factor;
            }
            else if (reader.name() == QLatin1String("x")) d->x = val.toDouble();
            else if (reader.name() == QLatin1String("y")) d->y = val.toDouble();
            else if (reader.name() == QLatin1String("w")) d->w = val.toDouble();
            else if (reader.name() == QLatin1String("h")) d->h = val.toDouble();
            else reader.skipCurrentElement();
        }
        else if (token == QXmlStreamReader::EndElement && reader.name() == name)
            break;
    }
}

void Layout::write(QXmlStreamWriter &writer, const QString &name) const
{
    if (!d) {
        writer.writeEmptyElement(name);
        return;
    }

    writer.writeStartElement(name);
    writer.writeStartElement("c:manualLayout");

    if (d->layoutTarget.has_value()) {
        writer.writeEmptyElement("c:layoutTarget");
        switch (d->layoutTarget.value()) {
            case LayoutTarget::Inner: writer.writeAttribute("val", "inner"); break;
            case LayoutTarget::Outer: writer.writeAttribute("val", "outer"); break;
        }
    }
    if (d->xMode.has_value()) {
        writer.writeEmptyElement("c:xMode");
        switch (d->xMode.value()) {
            case LayoutMode::Edge: writer.writeAttribute("val", "edge"); break;
            case LayoutMode::Factor: writer.writeAttribute("val", "factor"); break;
        }
    }
    if (d->yMode.has_value()) {
        writer.writeEmptyElement("c:yMode");
        switch (d->yMode.value()) {
            case LayoutMode::Edge: writer.writeAttribute("val", "edge"); break;
            case LayoutMode::Factor: writer.writeAttribute("val", "factor"); break;
        }
    }
    if (d->wMode.has_value()) {
        writer.writeEmptyElement("c:wMode");
        switch (d->wMode.value()) {
            case LayoutMode::Edge: writer.writeAttribute("val", "edge"); break;
            case LayoutMode::Factor: writer.writeAttribute("val", "factor"); break;
        }
    }
    if (d->hMode.has_value()) {
        writer.writeEmptyElement("c:hMode");
        switch (d->hMode.value()) {
            case LayoutMode::Edge: writer.writeAttribute("val", "edge"); break;
            case LayoutMode::Factor: writer.writeAttribute("val", "factor"); break;
        }
    }
    if (d->x.has_value()) {
        writer.writeEmptyElement("c:x");
        writer.writeAttribute("val", QString::number(d->x.value()));
    }
    if (d->y.has_value()) {
        writer.writeEmptyElement("c:y");
        writer.writeAttribute("val", QString::number(d->y.value()));
    }
    if (d->w.has_value()) {
        writer.writeEmptyElement("c:w");
        writer.writeAttribute("val", QString::number(d->w.value()));
    }
    if (d->h.has_value()) {
        writer.writeEmptyElement("c:h");
        writer.writeAttribute("val", QString::number(d->h.value()));
    }

    writer.writeEndElement(); //c:manualLayout
    writer.writeEndElement(); //name
}

LayoutPrivate::LayoutPrivate()
{

}

LayoutPrivate::LayoutPrivate(const LayoutPrivate &other) : QSharedData(other),
    layoutTarget(other.layoutTarget), xMode(other.xMode), yMode(other.yMode),
    wMode(other.wMode), hMode(other.hMode), x(other.x), y(other.y), w(other.w),
    h(other.h)
{

}

LayoutPrivate::~LayoutPrivate()
{

}

bool LayoutPrivate::operator ==(const LayoutPrivate &other) const
{
    if (layoutTarget != other.layoutTarget) return false;
    if (xMode != other.xMode) return false;
    if (yMode != other.yMode)  return false;
    if (wMode != other.wMode) return false;
    if (hMode != other.hMode) return false;
    if (x != other.x) return false;
    if (y != other.y) return false;
    if (w != other.w) return false;
    if (h != other.h) return false;
    return true;
}

QDebug operator<<(QDebug dbg, const Layout &f)
{
    QDebugStateSaver state(dbg);
    dbg.nospace() << "QXlsx::Layout(";
    dbg.nospace() << "layoutTarget=";
    QString s;
    if (f.d->layoutTarget.has_value()) {
        dbg.nospace() << f.toString(f.d->layoutTarget.value()) << ", ";
    }
    else dbg.nospace() << "not def, ";
    if (f.d->xMode.has_value()) {
        dbg.nospace() << f.toString(f.d->xMode.value()) << ", ";
    }
    else dbg.nospace() << "not def, ";
    if (f.d->yMode.has_value()) {
        dbg.nospace() << f.toString(f.d->yMode.value()) << ", ";
    }
    else dbg.nospace() << "not def, ";
    if (f.d->wMode.has_value()) {
        dbg.nospace() << f.toString(f.d->wMode.value()) << ", ";
    }
    else dbg.nospace() << "not def, ";
    if (f.d->hMode.has_value()) {
        dbg.nospace() << f.toString(f.d->hMode.value()) << ", ";
    }
    else dbg.nospace() << "not def, ";
    if (f.d->x.has_value()) {
        dbg.nospace() << f.d->x.value() << ", ";
    }
    else dbg.nospace() << "not def, ";
    if (f.d->y.has_value()) {
        dbg.nospace() << f.d->y.value() << ", ";
    }
    else dbg.nospace() << "not def, ";
    if (f.d->w.has_value()) {
        dbg.nospace() << f.d->w.value() << ", ";
    }
    else dbg.nospace() << "not def, ";
    if (f.d->h.has_value()) {
        dbg.nospace() << f.d->h.value();
    }
    else dbg.nospace() << "not def";
    dbg.nospace() << ")";
    return dbg;
}

}
