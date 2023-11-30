#ifndef XLSXSHEETVIEW_H
#define XLSXSHEETVIEW_H

#include "xlsxglobal.h"

#include <QString>
#include <QXmlStreamReader>
#include <QXmlStreamWriter>

#include <optional>

#include "xlsxcellreference.h"
#include "xlsxcellrange.h"
#include "xlsxmain.h"
#include "xlsxutility_p.h"

namespace QXlsx {

/**
     * @brief The ViewPane struct represents a view pane in the sheet view
     */
struct ViewPane
{
    /**
     * @brief The Pane enum specifies the pane that is active in the sheet view.
     */
    enum class Type
    {
        BottomRight,
        TopRight,
        BottomLeft,
        TopLeft
    };
    /**
     * @brief The State enum specifies whether the pane has horizontal /
     * vertical splits, and whether those splits are frozen.
     */
    enum class State
    {
        Split, /**< Panes are split, but not frozen. In this state, the split
bars are adjustable by the user. */
        Frozen, /**< Panes are frozen, but were not split being frozen. In
this state, when the panes are unfrozen again, a single pane results, with no split.
In this state, the split bars are not adjustable. */
        FrozenSplit /**< Panes are frozen and were split before being frozen. In
this state, when the panes are unfrozen again, the split remains, but is adjustable. */
    };
    /**
     * @brief Specifies the horizontal position of the split, in 1/20th of a
     * point; 0 (zero) if none.  If the pane is frozen, this value indicates the
     * number of columns visible in the top pane.
     *
     * If not set, 0 is assumed.
     */
    std::optional<double> xSplit; //default=0
    /**
     * @brief Specifies the vertical position of the split, in 1/20th of a point;
     * 0 (zero) if none.  If the pane is frozen, this value indicates the number
     * of rows visible in the left pane.
     *
     * If not set, 0 is assumed.
     */
    std::optional<double> ySplit; //default=0
    /**
     * @brief Specifies the location of the top left visible cell in the bottom
     * right pane (when in Left-To-Right mode).
     */
    CellReference topLeftCell;
    /**
     * @brief Specifies the pane that is active.
     *
     * If not set, Type::TopLeft is assumed.
     */
    std::optional<Type> activePane; //default = TopLeft
    /**
     * @brief Indicates whether the pane has horizontal / vertical splits, and
     * whether those splits are frozen.
     *
     * If not set, State::Split is assumed.
     * @note Excel and LibreOffice seem to not understand Frozen and FrozenSplit
     * states.
     */
    std::optional<State> paneState; //default = Split
    /**
     * @brief returns whether the pane is valid, that is whether any of its
     * parameters is set.
     */
    bool isValid() const;

    SERIALIZE_ENUM(Type, {
        {Type::TopLeft, "topLeft"},
        {Type::BottomRight, "bottomRight"},
        {Type::TopRight, "topRight"},
        {Type::BottomLeft, "bottomLeft"},
    });
    SERIALIZE_ENUM(State, {
        {State::Frozen, "frozen"},
        {State::FrozenSplit, "frozenSplit"},
        {State::Split, "split"}
    });
};

/**
 * @brief The Selection struct represents a selected region in the worksheet.
 */
struct QXLSX_EXPORT Selection
{
    /**
     * @brief The pane to which this selection belongs.
     *
     * If not set, the default value is PaneType::TopLeft.
     */
    std::optional<ViewPane::Type> pane;
    /**
     * @brief Location of the active cell.
     *
     * If invalid, the default value is (1,1), that is cell "A1".
     */
    CellReference activeCell;
    /**
     * @brief 0-based index of the range reference (in #selectedRanges) that
     * contains the active cell.
     *
     * Only used when #selectedRanges has more that one element. Therefore, this
     * value needs to be aware of the order in which the range references are
     * added to #selectedRanges.
     *
     * When this value is out of range then #activeCell can be used.
     *
     * If not set, the default value is 0.
     */
    std::optional<int> activeCellId;
    /**
     * @brief Ranges of the selection. Can be non-contiguous set of ranges.
     */
    QList<CellRange> selectedRanges;

    bool isValid() const;
    void read(QXmlStreamReader &reader);
    void write(QXmlStreamWriter &writer, const QLatin1String &name) const;
};


/**
 * @brief The SheetView struct represents one of the 'views' into the sheet.
 *
 * All SheetView members are accessible via AbstractSheet::view() and AbstractSheet::addView() methods.
 *
 * Some parameters are common to both worksheets and chartsheets: #tabSelected,
 * #zoomScale, #workbookViewId. They are accessible via the corresponding AbstractSheet methods.
 *
 * Most of parameters are applicable only to worksheets: #windowProtection, #showFormulas,
 * #showGridLines, #showRowColHeaders, #showZeros, #rightToLeft, #showRuler, #showOutlineSymbols,
 * #showWhiteSpace, #defaultGridColor, #type, #topLeftCell, #colorId, #zoomScaleNormal,
 * #zoomScalePageBreakView, #zoomScalePageLayoutView, #selection, #pane, . These parameters are accessible
 * via the corresponding Worksheet methods.
 *
 * One parameter is valid to only chartsheets: #zoomToFit. The corresponding methods
 * are added to the Chartsheet class.
 */
struct QXLSX_EXPORT SheetView
        //TODO: convert to implicitly shareable class
{
    /**
     * @brief The Type enum defines the kinds of view available to an application
     * when rendering a SpreadsheetML document.
     */
    enum class Type
    {
        Normal, /**< Specifies that the worksheet should be displayed without regard to pagination.*/
        PageBreakPreview, /**<  Specifies that the worksheet should be displayed
showing where pages would break if it were printed. */
        PageLayout /**< Specifies that the worksheet should be displayed
mimicking how it would look if printed. */
    };

    /**
     * @brief Flag indicating whether the panes in the window are locked due to
     * workbook protection.
     *
     * This is an option when the workbook structure is protected.
     *
     * If not set, the default value is `false`.
     *
     * This parameter is not applicable to chartsheets.
     */
    std::optional<bool> windowProtection;
    /**
     * @brief Flag indicating whether this sheet should display formulas.
     *
     * If not set, the default value is `false`.
     *
     * This parameter is not applicable to chartsheets.
     */
    std::optional<bool> showFormulas;
    /**
     * @brief Flag indicating whether this sheet should display gridlines.
     *
     * If not set, the default value is `true`.
     *
     * This parameter is not applicable to chartsheets.
     */
    std::optional<bool> showGridLines;
    /**
     * @brief Flag indicating whether the sheet should display row and column headings.
     *
     * If not set, the default value is `true`.
     *
     * This parameter is not applicable to chartsheets.
     */
    std::optional<bool> showRowColHeaders;
    /**
     * @brief Flag indicating whether the window should show 0 (zero) in cells
     * containing zero value.
     *
     * When `false`, cells with zero value appear blank instead of showing the
     * number zero.
     *
     * If not set, the default value is `true`.
     *
     * This parameter is not applicable to chartsheets.
     */
    std::optional<bool> showZeros;
    /**
     * @brief Flag indicating whether the sheet is in 'right to left' display
     * mode.
     *
     * When in this mode, Column A is on the far right, Column B is one column
     * left of Column A, and so on. Also, information in cells is displayed in
     * the Right to Left format.
     *
     * If not set, the default value is `false`.
     *
     * This parameter is not applicable to chartsheets.
     */
    std::optional<bool> rightToLeft;
    /**
     * @brief Flag indicating whether this sheet is selected.
     *
     * When only one sheet is selected and active, this value should be in synch
     * with the activeTab value.
     *
     * If not set, the default value is `false`.
     */
    std::optional<bool> tabSelected;
    /**
     * @brief Show the ruler in page layout view.
     *
     * If not set, the default value is `false`.
     *
     * This parameter is not applicable to chartsheets.
     */
    std::optional<bool> showRuler;
    /**
     * @brief Flag indicating whether the sheet has outline symbols visible.
     *
     * If not set, the default value is `true`.
     *
     * This parameter is not applicable to chartsheets.
     */
    std::optional<bool> showOutlineSymbols;
    /**
     * @brief Flag indicating whether page layout view shall display margins.
     *
     * `false` means do not display left, right, top (header), and bottom
     * (footer) margins (even when there is data in the header or footer).
     *
     * If not set, the default value is `true`.
     *
     * This parameter is not applicable to chartsheets.
     */
    std::optional<bool> showWhiteSpace;
    /**
     * @brief Flag indicating that the application should use the default grid
     * lines color (system dependent).
     *
     * Overrides any color specified in colorId.
     *
     * If not set, the default value is `true`.
     *
     * This parameter is not applicable to chartsheets.
     */
    std::optional<bool> defaultGridColor;
    /**
     * @brief View type.
     *
     * If not set, the default value is Type::Normal.
     *
     * This parameter is not applicable to chartsheets.
     */
    std::optional<Type> type;
    /**
     * @brief Location of the top left visible cell.
     *
     * This parameter is not applicable to chartsheets.
     */
    CellReference topLeftCell;
    /**
     * @brief  Index to the color value for row/column text headings and
     * gridlines. This is an 'index color value' (ICV) rather than rgb value.
     *
     * If not set, the default value is 64.
     *
     * This parameter is not applicable to chartsheets.
     */
    std::optional<int> colorId;
    /**
     * @brief Window zoom magnification for current view representing percent
     * values.
     *
     * This parameter is restricted to values ranging from 10 to 400.
     *
     * If not set, the default value is 100.
     */
    std::optional<int> zoomScale;
    /**
     * @brief Zoom magnification to use when in normal view, representing percent values.
     *
     * This attribute is restricted to values ranging from 10 to 400. Zero value implies
     * automatic setting. Horizontal & Vertical scale together.
     *
     * If not set, the default value is 0 (auto zoom scale in normal view).
     *
     * This parameter is not applicable to chartsheets.
     */
    std::optional<int> zoomScaleNormal;
    /**
     * @brief Zoom magnification to use when in page break view, representing percent values.
     *
     * This attribute is restricted to values ranging from 10 to 400. Zero value implies
     * automatic setting. Horizontal & Vertical scale together.
     *
     * If not set, the default value is 0 (auto zoom scale in page break view).
     *
     * This parameter is not applicable to chartsheets.
     */
    std::optional<int> zoomScalePageBreakView;
    /**
     * @brief Zoom magnification to use when in page layout view, representing percent values.
     *
     * This attribute is restricted to values ranging from 10 to 400. Zero value implies
     * automatic setting. Horizontal & Vertical scale together.
     *
     * If not set, the default value is 0 (auto zoom scale in page layout view).
     *
     * This parameter is not applicable to chartsheets.
     */
    std::optional<int> zoomScalePageLayoutView;
    /**
     * @brief Zero-based index of this workbook view, pointing to a specific workbookView
     * element in the bookViews collection.
     */
    int workbookViewId = 0; //required
    /**
     * @brief Flag indicating whether chart sheet is zoom to fit window.
     *
     * The default value is `false` according to ECMA-376, but to mimic the Excel behavior if
     * the parameter is not set directly, then on writing the document zoomToFit is written as `true`.
     *
     * This parameter is not applicable to worksheets.
     */
    std::optional<bool> zoomToFit;
    /**
     * @brief Selection parameters of the worksheet.
     *
     * This parameter is not applicable to chartsheets.
     */
    Selection selection;
    /**
     * @brief Pane parameters of the view.
     *
     * This parameter is not applicable to chartsheets.
     */
    ViewPane pane;

//TODO:    std::optional<CT_PivotSelection> pivotSelection;

    ExtensionList extLst;

    bool isValid() const;
    void read(QXmlStreamReader &reader);
    void write(QXmlStreamWriter &writer, const QLatin1String &name, bool worksheet = true) const;
};

}

#endif // XLSXSHEETVIEW_H
