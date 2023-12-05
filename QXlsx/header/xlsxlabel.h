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
/**
 * @brief The Label class represents an individual series label in the case this
 * label should have properties distinct from other labels.
 *
 * The class has the same methods as the Labels class but these methods set /
 * return the parameter of only one label.
 */
class QXLSX_EXPORT Label
{
public:
    /**
     * @brief The Position enum specifies the label position relative to the
     * data marker.
     *
     * Some positions are not applicable to all series type. Tests are needed to
     * check all combinations.
     *
     * Position | Bar | Line | Scatter
     * ----|----|----|----
     * BestFit | - | - | -
     * Left | - | + | +
     * Right | - | + | +
     * Top | - | + | +
     * Bottom | - | + | +
     * Center | + | + | +
     * InBase | + | - | -
     * InEnd | + | - | -
     * OutEnd | + | - | -
     *
     */
    enum class Position
    {
        BestFit, /**< Best position.  */
        Left, /**< To the left of data marker. */
        Right, /**< To the right of data marker. */
        Top, /**< Above the data marker. Not applicable to bar charts. */
        Bottom, /**< Below the data marker. */
        Center, /**< Centered on the data marker. */
        InBase, /**< Inside the base of the data marker. */
        InEnd, /**< Inside the end of the data marker. */
        OutEnd /**< Outside the end of the data marker. */
    };
    /**
     * @brief The ShowParameter enum specifies the elements that are shown in
     * the label text.
     */
    enum ShowParameter
    {
        ShowLegendKey = 1, /**< Show the legend key. */
        ShowValue = 2, /**< Show the data point value. */
        ShowCategory = 4, /**< Show the data point category. */
        ShowSeries = 8, /**< Show the series name. */
        ShowPercent = 16, /**< Show the percentage. */
        ShowBubbleSize = 32 /**< Show the bubble size. */
    };
    Q_DECLARE_FLAGS(ShowParameters, ShowParameter)

    /**
     * @brief Creats an invalid (default) Label.
     */
    Label();
    /**
     * @brief Creates Label for data point @a index with show parameters and
     * position specified.
     * @param index Zero-based index of a data point.
     * @param show A combination of elements shown in the label text.
     * @param position Position of the label.
     */
    Label(int index, ShowParameters show, Position position);
    /**
     * @brief Creates a shallow copy of @a other.
     * @param other
     */
    Label(const Label &other);
    Label &operator=(const Label &other);
    ~Label();
    /**
     * @brief returns the index of the data point this label is attached to or
     * -1 if the label's not attached.
     */
    int index() const;
    /**
     * @brief sets the index of the data point this label is attached to.
     * @param index Zero-based index of the data point.
     */
    void setIndex(int index);
    /**
     * @brief returns whether the label is visible.
     *
     * The default value is `true`.
     */
    std::optional<bool> visible() const;
    /**
     * @brief sets the visibility of the label.
     * @param visible
     *
     * If not set, `true` is assumed.
     */
    void setVisible(bool visible);
    /**
     * @brief returns the layout parameters of the label.
     */
    Layout layout() const;
    /**
     * @brief returns the layout parameters of the label.
     */
    Layout &layout();
    /**
     * @brief sets the layout parameters of the label.
     * @param layout
     */
    void setLayout(const Layout &layout);
    /**
     * @brief returns the text of the label.
     */
    Text text() const;
    /**
     * @brief sets the text of the label.
     *
     * If the text is set, then #showParameters() is ignored and @a text is
     * shown instead.
     *
     * To remove the text use `setText(Text())`.
     *
     * @param text
     */
    void setText(const Text &text);


    // Properties common to Label and Labels

    /**
     * @brief returns the elements shown in the label text.
     * @return A combination of Label::ShowParameter flags.
     * @sa #testShowParameter(), #testShowCategory(), #testShowLegendKey() etc.
     */
    ShowParameters showParameters() const;
    /**
     * @brief sets the elements shown in the label text.
     * @param showParameters A combination of Label::ShowParameter flags.
     * @sa #setShowParameter(), #setShowCategory(), #setShowLegendKey() etc.
     */
    void setShowParameters(ShowParameters showParameters);
    /**
     * @brief sets the flag of the element shown in the label text.
     * @param parameter A ShowParameter enum value.
     * @param value `true` if show, `false` if do not show.
     * @sa #setShowParameters(), #setShowCategory(), #setShowLegendKey() etc.
     *
     * If not set, the label will get the default flags either specified
     * in the Labels object or by the spreadsheet app.
     */
    void setShowParameter(ShowParameter parameter, bool value);
    /**
     * @brief tests whether the the element is shown in the label text.
     * @param parameter element to test.
     * @sa #showParameters().
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
     * If not set, the label will get the default position either specified
     * in the Labels object or by the spreadsheet app.
     */
    void setPosition(Position pos);
    /**
     * @brief returns the label's position.
     * @return Position enum value if it was set, `nullopt` otherwise. If not set,
     * Position::BestFit is assumed.
     */
    std::optional<Position> position() const;
    /**
     * @brief returns the number format that is used for the label.
     */
    QString numberFormat() const;
    /**
     * @brief sets the number format that will be used for the label.
     * @param numberFormat A string that contains the format. If empty, the
     * label will get the default number format either specified in the Labels
     * object or by the spreadsheet app.
     * @param formatSourceLinked An additional informative flag.
     */
    void setNumberFormat(const QString &numberFormat, bool formatSourceLinked = false);
    /**
     * @brief returns the label shape format.
     */
    ShapeFormat shape() const;
    /**
     * @overload
     * @brief returns the label shape format.
     */
    ShapeFormat &shape();
    /**
     * @brief sets the label shape format.
     * @param shape If valid, the label will be shown with this shape format.
     * If invalid, the label will get the default shape format either specified
     * in the Labels object or by the spreadsheet app.
     */
    void setShape(const ShapeFormat &shape);
    /**
     * @brief returns the label text format.
     */
    TextFormat textFormat() const;
    /**
     * @brief returns the label text format.
     */
    TextFormat &textFormat();
    /**
     * @brief sets the label text format.
     * @param textFormat If valid, the label will be shown with this text
     * format. If invalid, the label will get the default text format either
     * specified in the Labels object or by the spreadsheet app.
     */
    void setTextFormat(const TextFormat &textFormat);
    /**
     * @brief returns the label separator.
     */
    QString separator() const;
    /**
     * @brief sets the label separator.
     * @param separator A string that will be used to divide the label sections.
     * If empty, the label will get the default separator either specified in
     * the Labels object or by the spreadsheet app.
     */
    void setSeparator(const QString &separator);
    /**
     * @brief clears all properties that were set for this particular label
     * (#showParameters(), #position(), #numberFormat(),
     * #shape(), #textFormat(), #separator()).
     */
    void clearProperties();

    bool isValid() const;

    bool operator ==(const Label &other) const;
    bool operator !=(const Label &other) const;
    operator QVariant() const;
private:
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
    void read(QXmlStreamReader &reader);
    void write(QXmlStreamWriter &writer) const;
    friend QDebug operator<<(QDebug dbg, const Label &f);
    friend class Labels;
    friend class SharedLabelProperties;
    QSharedDataPointer<LabelPrivate> d;
};

QDebug operator<<(QDebug dbg, const Label &f);

class LabelsPrivate;

/**
 * @brief The Labels class represents parameters of labels in the chart series.
 *
 * The series labels may have default parameters specified with a number of
 * methods: #leaderLines(), #showParameters(), #position(), #numberFormat(),
 * #shape(), #textFormat(), #separator().
 *
 * Each individual label attached to a specific data point may have its
 * own parameters set via #label() and #labelForPoint(), but such a label needs
 * to be added manually with #addLabel().
 *
 * You don't need to add labels for all data points if these labels are shown
 * with the same parameters.
 *
 * To show / hide all labels use #setVisible(). To reset the default parameters
 * use #clearDefaultProperties(). To remove attached labels use #removeLabel(),
 * #removeLabelForPoint() and #removeLabels().
 */
class QXLSX_EXPORT Labels
{
public:
    /**
     * @brief creates invalid (default) Labels.
     */
    Labels();
    /**
     * @brief creates Labels with shown elements and their position specified.
     * @param showFlags specifies which elements to show in the labels text.
     * @param pos position of labels.
     */
    Labels(Label::ShowParameters showFlags, Label::Position pos);
    /**
     * @brief creates a shallow copy of @a other.
     * @param other
     */
    Labels(const Labels &other);
    ~Labels();
    Labels &operator=(const Labels &other);

    bool operator==(const Labels &other) const;
    bool operator!=(const Labels &other) const;
    operator QVariant() const;
    /**
     * @brief returns whether any of the labels parameters were set.
     */
    bool isValid() const;
    /**
     * @brief returns whether the labels are visible.
     *
     * The default value is `false` (not visible).
     */
    std::optional<bool> visible() const;
    /**
     * @brief sets visibility of labels.
     * @param visible If `true` then labels will be shown.
     *
     * If not set, `false` (not visible) is assumed.
     */
    void setVisible(bool visible);
    /**
     * @brief returns whether to show the leader lines attached to labels.
     *
     * The default value is `false` (do not show).
     */
    std::optional<bool> showLeaderLines() const;
    /**
     * @brief sets visibility of the leader lines attached to labels.
     * @param showLeaderLines If `true` then a line will be drawn between a
     * label and a data point.
     *
     * If not set, `false` (do not show) is assumed.
     */
    void setShowLeaderLines(bool showLeaderLines);
    /**
     * @brief returns the leader lines shape format.
     */
    ShapeFormat leaderLines() const;
    /**
     * @overload
     * @brief returns the leader lines shape format.
     */
    ShapeFormat &leaderLines();
    /**
     * @brief sets the leader lines shape format.
     * @param format
     *
     * This method does not change the visibility of leader lines.
     */
    void setLeaderLines(const ShapeFormat &format);
    /**
     * @brief returns the elements shown in the labels text.
     * @return A combination of Label::ShowParameter flags.
     * @sa #testShowParameter(), #testShowCategory(), #testShowLegendKey() etc.
     */
    Label::ShowParameters showParameters() const;
    /**
     * @brief sets the elements shown in the labels text.
     * @param showParameters A combination of Label::ShowParameter flags.
     * @sa #setShowParameter(), #setShowCategory(), #setShowLegendKey() etc.
     */
    void setShowParameters(Label::ShowParameters showParameters);
    /**
     * @brief sets the flag of the element shown in the labels text.
     * @param parameter A ShowParameter enum value.
     * @param value `true` if show, `false` if do not show.
     * @sa #setShowParameters(), #setShowCategory(), #setShowLegendKey() etc.
     */
    void setShowParameter(Label::ShowParameter parameter, bool value);
    /**
     * @brief tests whether the the element is shown in the labels text.
     * @param parameter element to test.
     * @sa #showParameters().
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
     * @brief sets the default label position that will be used for all labels
     * in this labels collection.
     * @param pos Position enum value.
     *
     * If not set, Position::BestFit is assumed.
     */
    void setPosition(Label::Position pos);
    /**
     * @brief returns the default label position that is used for all labels in
     * this labels collection.
     * @return Position enum value if it was set, `nullopt` otherwise. If not
     * set, Position::BestFit is assumed.
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
     * @overload
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
     * If some properties were set for a specific label, this label will be
     * shown according to its properties.
     */
    void clearDefaultProperties();

    /**
     * @brief adds @a label to the list of labels.
     * @param label
     * @return `true` if @a label was added, `false` if the labels list already
     * has the label with the same data point index.
     */
    bool addLabel(const Label &label);
    /**
     * @brief adds new label to the list of labels.
     * @param index data point index.
     * @param showFlags parameters of a new label.
     * @param position position of a new label.
     * @return `true` if a new label was added, `false` if the labels list
     * already has the label with the same data point @a index.
     */
    bool addLabel(int index, Label::ShowParameters showFlags, Label::Position position);
    /**
     * @brief removes the label with @a index.
     * @param index valid label index (0 to #labelsCount()-1).
     * @return `true` if the label was found and removed, `false` otherwise.
     */
    bool removeLabel(int index);
    /**
     * @brief removes the label that was associated with data point @a index.
     * @param index data point index.
     * @return `true` if the label was found and removed, `false` otherwise.
     * @sa label(), labelForPoint(), addLabel()
     */
    bool removeLabelForPoint(int index);
    /**
     * @brief removes all labels.
     *
     * Removes all labels with specific parameters, leaving default labels
     * parameters untouched.
     */
    void removeLabels();
    /**
     * @brief returns the label with @a index.
     * @param index index of a label in the labels list (0 to labelsCount()-1).
     * @return A copy of the label.
     */
    Label label(int index) const;
    /**
     * @brief returns the label with @a index.
     * @param index valid index of a label in the labels list (0 to
     * #labelsCount()-1).
     * @return a reference to the label.
     *
     * If @a index is invalid, the behaviour is undefined.
     *
     * The labels list contains the label with @a index only if it was manually
     * added with #addLabel() or in the spreadsheet app.
     *
     * To get the label by the data point index use #labelForPoint(). Compare:
     *
     * @code
     * chart->labels(0).addLabel(4, //label will be attached to the 5th bar of a bar series
     * Label::ShowValue | Label::ShowCategory,
     * Label::Position::Top);
     * chart->labels(0).label(0).setNumberFormat("0.000"); // address the label by its index
     * chart->labels(0).labelForPoint(4).setSeparator(" - "); // address the label by data point index
     * @endcode
     */
    Label& label(int index);
    /**
     * @brief returns the label associated with data point @a index.
     * @param index index of a data point.
     *
     * If there's no label associated with data point @a index, creates a
     * default one. You can test @a index with #hasLabelForPoint().
     *
     * @return valid reference if @a index is valid, `nullopt` otherwise.
     *
     * Compare:
     *
     * @code
     * chart->labels(0).addLabel(4, //label will be attached to the 5th bar of a bar series
     * Label::ShowValue | Label::ShowCategory,
     * Label::Position::Top);
     * chart->labels(0).label(0).setNumberFormat("0.000"); // address the label by its index
     * chart->labels(0).labelForPoint(4).setSeparator(" - "); // address the label by data point index
     * @endcode
     */
    std::optional<std::reference_wrapper<Label>> labelForPoint(int index);
    /**
     * @brief returns the label associated with data point @a index.
     * @param index index of a data point.
     * @return Valid Label object if such label exists, invalid one otherwise.
     * @sa label(), addLabel().
     */
    Label labelForPoint(int index) const;
    /**
     * @brief returns the count of labels that have parameters distinct from the
     * default parameters, that is the count of labels manually added to the
     * labels list with #addLabel() or in the spreadsheet app.
     */
    int labelsCount() const;
    /**
     * @brief returns whether the labels list contains the label associated with
     * the data point @a index.
     * @param index valid (non-negative) data point index.
     */
    bool hasLabelForPoint(int index) const;
private:
    friend class SubChart;
    friend class Chart;
    friend class Series;
    void read(QXmlStreamReader &reader);
    void write(QXmlStreamWriter &writer) const;
    QList<std::reference_wrapper<FillFormat> > fills();
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
