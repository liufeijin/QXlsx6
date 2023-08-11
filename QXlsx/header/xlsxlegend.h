#ifndef XLSXLEGEND_H
#define XLSXLEGEND_H

#include <QSharedData>
#include <QXmlStreamReader>
#include <QXmlStreamWriter>

#include <optional>

#include "xlsxglobal.h"
#include "xlsxlayout.h"
#include "xlsxshapeformat.h"
#include "xlsxtext.h"
#include "xlsxutility_p.h"

QT_BEGIN_NAMESPACE_XLSX

class LegendPrivate;

class QXLSX_EXPORT LegendEntry
{
public:
    LegendEntry();
    LegendEntry(int index, bool visible);
    LegendEntry(int index, bool visible, const Text &text);

    int idx = -1;
    std::optional<bool> visible;
    Text entry;

private:
    friend class Legend;
    void read(QXmlStreamReader &reader);
    void write(QXmlStreamWriter &writer, const QString &name) const;
};

class QXLSX_EXPORT Legend
{
public:
    enum class Position
    {
        Top,
        Bottom,
        Left,
        TopRight,
        Right
    };

    Legend();
    Legend(Position position);
    Legend(const Legend &other);
    ~Legend();
    Legend &operator=(const Legend &other);

    bool isValid() const;

    std::optional<Position> position() const;
    void setPosition(Position position);

    std::optional<bool> overlay() const;
    void setOverlay(bool overlay);

    Layout layout() const;
    Layout &layout();
    void setLayout(const Layout &layout);

    ShapeFormat shape() const;
    ShapeFormat &shape();
    void setShape(const ShapeFormat &shape);

    Text text() const;
    Text &text();
    void setText(const Text &text);

    void setEntryVisible(int index, bool visible);
    std::optional<bool> entryVisible(int index) const;

    /**
     * @brief addEntry adds a new legend entry with the text and visible for series
     * with index or modifies the existing legend entry.
     *
     * Use this method to manually remove the series entry from the legend, as
     * by default the chart legend does not have any entries (this means all the
     * series are visible in the chart legend, and they have the default text).
     *
     * @param index series index
     * @param text entry text
     * @param visible visibility of the entry.
     */
    void addEntry(int index, const Text &text, bool visible);
    /**
     * @brief entry returns legend entry for the series with index index. Such an entry
     * must be added to the legend manually, as by default the chart legend does not
     * have any entries (this means all the series are visible in the chart legend, and
     * they have the default text).
     * @param index m
     * @return
     */
    LegendEntry &entry(int index);

    void read(QXmlStreamReader &reader);
    void write(QXmlStreamWriter &writer) const;
private:

    SERIALIZE_ENUM(Position,
    {
        {Position::Top, "t"},
        {Position::Bottom, "b"},
        {Position::Left, "l"},
        {Position::TopRight, "tr"},
        {Position::Right, "r"},
    });
    QSharedDataPointer<LegendPrivate> d;
};


QT_END_NAMESPACE_XLSX

#endif // XLSXLEGEND_H
