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

namespace QXlsx {

class LegendPrivate;

/**
 * @brief The LegendEntry class specifies properties for a single legend entry: series,
 * trend lines etc.
 */
class QXLSX_EXPORT LegendEntry
{
public:
    /**
     * @brief creates an invalid LegendEntry.
     */
    LegendEntry();
    /**
     * @brief creates a LegendEntry and sets its _index_.
     * @param index the entry index (series, trend lines etc.).
     */
    LegendEntry(int index);
    /**
     * @brief creates a LegendEntry for a series with _index_, sets the _text_ and makes it _visible_.
     * @param index the entry index (series, trend lines etc.).
     * @param visible visibility of the LegendEntry.
     * @param text text of the LegendEntry.
     */
    LegendEntry(int index, bool visible, const TextFormat &text);

    int idx = -1;
    std::optional<bool> visible;
    TextFormat entry;

    bool isValid() const;
    bool operator==(const LegendEntry &other) const;

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
    /**
     * @brief creates an invalid Legend.
     */
    Legend();
    /**
     * @brief creates Legend positioned at #position.
     * @param position legend position.
     */
    Legend(Position position);
    Legend(const Legend &other);
    ~Legend();
    Legend &operator=(const Legend &other);

    bool isValid() const;
    /**
     * @brief creates Legend with default parameters.
     * The legend created will be valid, but all its parameters will not be set.
     * @return new Legend.
     */
    static Legend defaultLegend();

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

    TextFormat text() const;
    TextFormat &text();
    void setText(const TextFormat &text);

    /**
     * @brief sets visibility to the specific legend entry.
     * @param index index of the entry.
     * @param visible visibility of the legend entry.
     *
     * By default all legend entries are visible. You can use this method to change
     * the visibility of a specific entry.
     */
    void setEntryVisible(int index, bool visible);
    /**
     * @brief returns the visibility of a specific legend entry.
     * @param index the entry index.
     * @return valid std::optional if the visibility of the entry was changed,
     * nullopt if the visibility of the entry is the default.
     */
    std::optional<bool> entryVisible(int index) const;

    /**
     * @brief returns legend entry for the series with index index.
     * @param index the entry index. If no entry with this index is present in legend,
     * creates a new legend entry.
     * @return reference to the legend entry.
     */
    LegendEntry &entry(int index); //TODO: test this method

    void read(QXmlStreamReader &reader);
    void write(QXmlStreamWriter &writer) const;

    bool operator ==(const Legend &other) const;
    bool operator !=(const Legend &other) const;
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


}

#endif // XLSXLEGEND_H
