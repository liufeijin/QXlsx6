#ifndef XLSXLABEL_H
#define XLSXLABEL_H

#include "xlsxglobal.h"

#include <QSharedData>
#include <QXmlStreamReader>
#include <QXmlStreamWriter>
#include <optional>

#include "xlsxutility_p.h"
#include "xlsxmain.h"
#include "xlsxtext.h"
#include "xlsxshapeformat.h"
#include "xlsxlayout.h"

QT_BEGIN_NAMESPACE_XLSX

class LabelPrivate;
class SharedLabelProperties;

class QXLSX_EXPORT Label
{
public:
    enum class Position
    {
        BestFit,
        Left,
        Right,
        Top,
        Bottom,
        Center,
        InBase,
        InEnd,
        OutEnd
    };

    enum ShowParameter
    {
        ShowLegendKey = 1,
        ShowValue = 2,
        ShowCategory = 4,
        ShowSeries = 8,
        ShowPercent = 16,
        ShowBubbleSize = 32
    };
    Q_DECLARE_FLAGS(ShowParameters, ShowParameter)

    Label();
    Label(int index, ShowParameters show, Position position);
    Label(const Label &other);
    Label &operator=(const Label &other);
    ~Label();

    int index() const;
    void setIndex(int index);

    std::optional<bool> visible() const;
    void setVisible(bool visible);

    SharedLabelProperties properties() const;
    void setProperties(SharedLabelProperties properties);

    Layout layout() const;
    void setLayout(const Layout &layout);

    Text text() const;
    void setText(const Text &text);

    // TODO: the rest of
    void setShowCategory(bool show);

    void setPosition(Position pos);
    Position position() const;

    void read(QXmlStreamReader &reader);
    void write(QXmlStreamWriter &writer) const;

    bool isValid() const;

    bool operator ==(const Label &other) const;
    bool operator !=(const Label &other) const;
    operator QVariant() const;

private:
    friend QDebug operator<<(QDebug dbg, const Label &f);
    SERIALIZE_ENUM(Position, {
        {Position::BestFit, "bestFit"},
        {Position::Left, "l"},
        {Position::Right, "r"},
        {Position::Top, "t"},
        {Position::Bottom, "b"},
        {Position::Center, "ctr"},
        {Position::InBase, "inBase"},
        {Position::InEnd, "inEnd"},
        {Position::OutEnd, "outEnd"},
    });

    QSharedDataPointer<LabelPrivate> d;
};

QDebug operator<<(QDebug dbg, const Label &f);

class QXLSX_EXPORT SharedLabelProperties
{
public:
    Label::ShowParameters showFlags;
    std::optional<Label::Position> pos;
    QString numberFormat;
    bool formatSourceLinked = false;
    ShapeFormat shape;
    Text text;
    QString separator;

    void read(QXmlStreamReader &reader);
    void write(QXmlStreamWriter &writer) const;

    bool operator==(const SharedLabelProperties &other) const;
    bool operator!=(const SharedLabelProperties &other) const;
};

QDebug operator<<(QDebug dbg, const SharedLabelProperties &f);

class LabelsPrivate;

class QXLSX_EXPORT Labels
{
public:
    Labels();
    Labels(Label::ShowParameters showFlags, Label::Position pos);
    Labels(const Labels &other);
    ~Labels();
    Labels &operator=(const Labels &other);

    bool operator==(const Labels &other) const;
    bool operator!=(const Labels &other) const;
    operator QVariant() const;

    bool isValid() const;

    std::optional<bool> visible() const;
    void setVisible(bool visible);

    std::optional<bool> showLeaderLines() const;
    void setShowLeaderLines(bool showLeaderLines);

    ShapeFormat leaderLines() const;
    ShapeFormat &leaderLines();
    void setLeaderLines(const ShapeFormat &format);

    SharedLabelProperties defaultProperties() const;
    void setDefaultProperties(SharedLabelProperties defaultProperties);

    /**
     * @brief addLabel adds \label to the list of labels
     * @param label
     */
    void addLabel(const Label &label);
    /**
     * @brief addLabel adds new label to the list of labels
     * @param index data point index
     * @param showFlags parameters of a new label
     * @param position position of a new label
     */
    void addLabel(int index, Label::ShowParameters showFlags, Label::Position position);
    /**
     * @brief label
     * @param index index of a label in the labels list.
     * @return
     */
    Label label(int index) const;
    Label &label(int index);
    Label &labelForPoint(int index);
    Label labelForPoint(int index) const;

    void read(QXmlStreamReader &reader);
    void write(QXmlStreamWriter &writer) const;
private:
    friend QDebug operator<<(QDebug dbg, const Labels &f);
    QSharedDataPointer<LabelsPrivate> d;
};

class LabelsPrivate : public QSharedData
{
public:
    LabelsPrivate();
    LabelsPrivate(const LabelsPrivate &other);
    ~LabelsPrivate();

    bool operator ==(const LabelsPrivate &other) const;

    QList<Label> labels;
    bool visible = false;
    std::optional<bool> showLeaderLines;
    ShapeFormat leaderLines;
    SharedLabelProperties defaultProperties;
};

Q_DECLARE_OPERATORS_FOR_FLAGS(Label::ShowParameters)

QDebug operator<<(QDebug dbg, const Labels &f);

QT_END_NAMESPACE_XLSX

Q_DECLARE_METATYPE(QXlsx::Labels)
Q_DECLARE_METATYPE(QXlsx::Label)


#endif // XLSXLABEL_H
