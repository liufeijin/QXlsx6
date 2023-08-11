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

QT_BEGIN_NAMESPACE_XLSX

class Worksheet;
class WorksheetPrivate;
class RichStringPrivate;
class SharedStrings;

class MarkerFormatPrivate;

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

    std::optional<int> size() const;
    void setSize(int size);

    ShapeFormat shape() const;
    ShapeFormat &shape();
    void setShape(ShapeFormat shape);

    void write(QXmlStreamWriter &writer) const;
    void read(QXmlStreamReader &reader);

    bool isValid() const;
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

QT_END_NAMESPACE_XLSX

#endif // XLSXMARKERFORMAT_H
