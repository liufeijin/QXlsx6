#ifndef XLSXMARKERFORMAT_H
#define XLSXMARKERFORMAT_H

#include <QFont>
#include <QColor>
#include <QByteArray>
#include <QList>
#include <QExplicitlySharedDataPointer>
#include <QVariant>
#include <QXmlStreamReader>
#include <QXmlStreamWriter>

#include "xlsxglobal.h"
#include "xlsxshapeformat.h"

namespace QXlsx {

class Worksheet;
class WorksheetPrivate;
class RichStringPrivate;
class SharedStrings;

class MarkerFormatPrivate;

/**
 * @brief The MarkerFormat class represents the series marker.
 *
 * The class is _implicitly shareable_: the deep copy occurs only in the non-const methods.
 *
 * The marker has type(), size() in pt and shape() properties.
 */
class QXLSX_EXPORT MarkerFormat
{
public:
    enum class MarkerType {
        None,
        Square,
        Diamond,
        Triangle,
        Cross,
        Star,
        Dot,
        Dash,
        Circle,
        Plus,
        Picture,
        X,
        Auto
    };

    MarkerFormat();
    MarkerFormat(MarkerType type);
    MarkerFormat(const MarkerFormat &other);
    MarkerFormat &operator=(const MarkerFormat &rhs);
    ~MarkerFormat();

    std::optional<MarkerType> type() const;
    void setType(MarkerType type);

    /**
     * @brief returns the marker size
     * @return valid optional value if the size was set, `nullopt` otherwise.
     */
    std::optional<int> size() const;
    /**
     * @brief sets the marker size.
     * @param size a number 2..72. If the marker size is not set, the default size is 5.
     */
    void setSize(int size);

    ShapeFormat shape() const;
    ShapeFormat &shape();
    void setShape(ShapeFormat shape);

    void write(QXmlStreamWriter &writer) const;
    void read(QXmlStreamReader &reader);

    bool isValid() const;

    bool operator==(const MarkerFormat &other) const;
    bool operator!=(const MarkerFormat &other) const;
private:
    SERIALIZE_ENUM(MarkerType,
    {
        {MarkerType::X, "x"},
        {MarkerType::None, "none"},
        {MarkerType::Square, "square"},
        {MarkerType::Dot, "dot"},
        {MarkerType::Dash, "dash"},
        {MarkerType::Diamond, "diamond"},
        {MarkerType::Auto, "auto"},
        {MarkerType::Plus, "plus"},
        {MarkerType::Star, "star"},
        {MarkerType::Cross, "cross"},
        {MarkerType::Circle, "circle"},
        {MarkerType::Picture, "picture"},
        {MarkerType::Triangle, "triangle"},
    });

    friend QDebug operator<<(QDebug, const MarkerFormat &f);

    QSharedDataPointer<MarkerFormatPrivate> d;
};

#ifndef QT_NO_DEBUG_STREAM
  QDebug operator<<(QDebug dbg, const MarkerFormat &f);
#endif

}

Q_DECLARE_TYPEINFO(QXlsx::MarkerFormat, Q_MOVABLE_TYPE);

#endif // XLSXMARKERFORMAT_H
