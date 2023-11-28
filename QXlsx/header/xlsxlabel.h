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

namespace QXlsx {

class LabelPrivate;
//TODO: full documentation of Label

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

    Layout layout() const;
    void setLayout(const Layout &layout);

    Text text() const;
    void setText(const Text &text);


    // Properties common to Label and Labels

    ShowParameters showParameters() const;
    void setShowParameters(ShowParameters showParameters);
    void setShowParameter(ShowParameter parameter, bool value);
    /**
     * @brief test if parameter is set.
     * @param parameter parameter to test.
     * @return `true` if parameter is set, `false` if the parameter is not set or label is not valid.
     */
    bool testShowParameter(ShowParameter parameter) const;

    void setShowCategory(bool show);
    void setShowLegendKey(bool show);
    void setShowValue(bool show);
    void setShowSeries(bool show);
    void setShowPercent(bool show);
    void setShowBubbleSize(bool show);

    bool testShowCategory() const;
    bool testShowLegendKey() const;
    bool testShowValue() const;
    bool testShowSeries() const;
    bool testShowPercent() const;
    bool testShowBubbleSize() const;

    /**
     * @brief sets the label position.
     * @param pos Position enum value.
     *
     * If not set, Position::BestFit is assumed.
     */
    void setPosition(Position pos);
    /**
     * @brief returns the label's position.
     * @return Position enum value if it was set, `nullopt` otherwise. If not set,
     * Position::BestFit is assumed.
     */
    std::optional<Position> position() const;

    QString numberFormat() const;
    void setNumberFormat(const QString &numberFormat, bool formatSourceLinked = false);
    ShapeFormat shape() const;
    ShapeFormat &shape();
    void setShape(const ShapeFormat &shape);
    TextFormat textFormat() const;
    TextFormat &textFormat();
    void setTextFormat(const TextFormat &textFormat);
    QString separator() const;
    void setSeparator(const QString &separator);
    /**
     * @brief clears all properties that were set individually for this
     * particular label (#showParameters(), #position(), #numberFormat(),
     * #shape(), #textFormat(), #separator()).
     */
    void clearProperties();


    void read(QXmlStreamReader &reader);
    void write(QXmlStreamWriter &writer) const;

    bool isValid() const;

    bool operator ==(const Label &other) const;
    bool operator !=(const Label &other) const;
    operator QVariant() const;

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
private:
    friend QDebug operator<<(QDebug dbg, const Label &f);
    friend class Labels;
    QSharedDataPointer<LabelPrivate> d;
};

QDebug operator<<(QDebug dbg, const Label &f);

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

    Label::ShowParameters showParameters() const;
    void setShowParameters(Label::ShowParameters showParameters);
    void setShowParameter(Label::ShowParameter parameter, bool value);
    /**
     * @brief tests whether the show parameter is set.
     * @param parameter parameter to test.
     */
    bool testShowParameter(Label::ShowParameter parameter) const;

    void setShowCategory(bool show);
    void setShowLegendKey(bool show);
    void setShowValue(bool show);
    void setShowSeries(bool show);
    void setShowPercent(bool show);
    void setShowBubbleSize(bool show);

    bool testShowCategory() const;
    bool testShowLegendKey() const;
    bool testShowValue() const;
    bool testShowSeries() const;
    bool testShowPercent() const;
    bool testShowBubbleSize() const;

    /**
     * @brief sets the default label position that will be used for all labels in
     * this labels collection.
     * @param pos Position enum value.
     *
     * If not set, Position::BestFit is assumed.
     */
    void setPosition(Label::Position pos);
    /**
     * @brief returns the default label position that is used for all labels in
     * this labels collection.
     * @return Position enum value if it was set, `nullopt` otherwise. If not set,
     * Position::BestFit is assumed.
     */
    std::optional<Label::Position> position() const;
    /**
     * @brief returns the default number format that is used for all labels in
     * this labels collection.
     */
    QString numberFormat() const;
    /**
     * @brief sets the default number format that will be used for all labels in
     * this labels collection.
     * @param numberFormat A string that contains the format.
     * @param formatSourceLinked An additional informative flag.
     */
    void setNumberFormat(const QString &numberFormat, bool formatSourceLinked = false);
    /**
     * @brief returns the label shape format that is used for all labels in
     * this labels collection.
     */
    ShapeFormat shape() const;
    /**
     * @brief returns the label shape format that is used for all labels in
     * this labels collection.
     */
    ShapeFormat &shape();
    /**
     * @brief sets the label shape format that will be used for all labels in
     * this labels collection.
     * @param shape If valid, labels will be shown with this shape format.
     * If invalid, labels will get the default shape format.
     */
    void setShape(const ShapeFormat &shape);
    /**
     * @brief returns the text format that is used for all labels in
     * this labels collection.
     */
    TextFormat textFormat() const;
    /**
     * @brief returns the text format that is used for all labels in
     * this labels collection.
     */
    TextFormat &textFormat();
    /**
     * @brief sets the text format that will be used for all labels in
     * this labels collection.
     * @param textFormat If valid, labels will be shown with this text format.
     * If invalid, labels will get the default text format.
     */
    void setTextFormat(const TextFormat &textFormat);
    /**
     * @brief returns the separator that is used for all labels in
     * this labels collection.
     */
    QString separator() const;
    /**
     * @brief sets the separator that will be used for all labels in
     * this labels collection.
     * @param separator A string that will be used to divide the label sections.
     */
    void setSeparator(const QString &separator);
    /**
     * @brief clears all properties that were set for this labels collection
     * (#showParameters(), #position(), #numberFormat(), #shape(),
     * #textFormat(), #separator()).
     *
     * If some properties were set for an individual label, this label will be
     * shown according to its properties.
     */
    void clearDefaultProperties();

    /**
     * @brief addLabel adds @a label to the list of labels
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
    std::optional<std::reference_wrapper<Label> > label(int index);
    /**
     * @brief labelForPoint
     * @param index index of a data point
     * @return
     */
    std::optional<std::reference_wrapper<Label>> labelForPoint(int index);
    Label labelForPoint(int index) const;

    void read(QXmlStreamReader &reader);
    void write(QXmlStreamWriter &writer) const;
    QList<std::reference_wrapper<FillFormat> > fills();
private:
    friend QDebug operator<<(QDebug dbg, const Labels &f);
    QSharedDataPointer<LabelsPrivate> d;
};

Q_DECLARE_OPERATORS_FOR_FLAGS(Label::ShowParameters)

QDebug operator<<(QDebug dbg, const Labels &f);

}

Q_DECLARE_METATYPE(QXlsx::Labels)
Q_DECLARE_METATYPE(QXlsx::Label)
Q_DECLARE_TYPEINFO(QXlsx::Label, Q_MOVABLE_TYPE);
Q_DECLARE_TYPEINFO(QXlsx::Labels, Q_MOVABLE_TYPE);


#endif // XLSXLABEL_H
