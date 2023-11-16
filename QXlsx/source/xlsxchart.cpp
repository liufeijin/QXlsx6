// xlsxchart.cpp

#include <QtGlobal>
#include <QString>
#include <QIODevice>
#include <QXmlStreamReader>
#include <QXmlStreamWriter>
#include <QDebug>
#include <QObject>
#include <QVector>
#include <QMap>
#include <QList>

#include <memory>

#include "xlsxchart.h"
#include "xlsxcellrange.h"
#include "xlsxutility_p.h"
#include "xlsxabstractooxmlfile_p.h"
#include "xlsxseries.h"
#include "xlsxworkbook.h"
#include "QBuffer"

namespace QXlsx {

class SubChart
{
public:
    SubChart(Chart::Type type);

    Chart::Type type = Chart::Type::None;

    /* Below are SubChart properties */

    //+area, +area3d, +line, +line3d, +bar, +bar3d, +surface, +surface3d,
    //+scatter, +pie, +pie3d, +doughnut, +ofpie, +radar, +stock, +bubble
    QList<Series> seriesList;
    //+area, +area3d, +line, +line3d
    std::optional<Chart::Grouping> grouping;
    //+bar, +bar3d
    std::optional<Chart::BarGrouping> barGrouping;

    //+line, +line3d, +scatter, +radar, +bar, +bar3d, +area, +area3d,
    //+pie, +pie3d, +doughnut, +ofpie, +bubble
    std::optional<bool> varyColors;

    //+line, +line3d, +stock, +scatter, +radar, +bar, +bar3d, +area, +area3d,
    //+pie, +pie3d, +doughnut, +ofpie, +bubble
    Labels labels;

    //+line, +line3d, +stock, +area, +area3d,
    std::optional<ShapeFormat> dropLines; //drop lines can be default

    //+line, +stock,
    std::optional<ShapeFormat> hiLowLines; //can be default

    //+line, +stock
    UpDownBar upDownBars; //optional

    //+line
    std::optional<bool> marker;
    std::optional<bool> smooth;

    //+area, +area3d, +line, +line3d, +bar, +bar3d, +surface, +surface3d,
    //+scatter, +pie, +pie3d, +doughnut, +ofpie, +radar, +stock, +bubble
    QList<int> axesIds;

    //+line3d, +bar3d, +area3d
    std::optional<int> gapDepth;

    //+scatter
    Chart::ScatterStyle scatterStyle = Chart::ScatterStyle::Marker;

    //+radar
    Chart::RadarStyle radarStyle = Chart::RadarStyle::Standard;

    //+bar, +bar3d
    Chart::BarDirection barDirection = Chart::BarDirection::Column;

    //+bar, +bar3d, +ofpie,
    std::optional<int> gapWidth;

    //+bar
    std::optional<int> overlap; //in %

    //+bar, +ofpie
    QList<ShapeFormat> serLines;

    //+bar3D
    std::optional<Series::BarShape> barShape;

    //+pie, +doughnut
    std::optional<int> firstSliceAng; // [0..360] ? why not Angle?

    //+doughnut
    std::optional<int> holeSize; // in %, 1..90

    //+ofpie
    Chart::OfPieType ofPieType = Chart::OfPieType::Pie;
    std::optional<Chart::SplitType> splitType;
    std::optional<double> splitPos;
    QList<int> customSplit;
    std::optional<int> secondPieSize; // in %, 5..200

    //+bubble
    std::optional<bool> bubble3D;
    std::optional<bool> showNegBubbles;
    std::optional<int> bubbleScale; // in % [0..300]
    std::optional<Chart::BubbleSizeRepresents> bubbleSizeRepresents;

    //+surface, +surface3d
    std::optional<bool> wireframe;
    QMap<int, ShapeFormat> bandFormats;

    Series *addSeries(int index);
    bool read(QXmlStreamReader &reader);
    void write(QXmlStreamWriter &writer) const;
    friend class Chart;
private:
    void loadAreaChart(QXmlStreamReader &reader);
    void loadSurfaceChart(QXmlStreamReader &reader);
    void loadBubbleChart(QXmlStreamReader &reader);
    void loadPieChart(QXmlStreamReader &reader);
    void loadLineChart(QXmlStreamReader &reader);
    void loadBarChart(QXmlStreamReader &reader);
    void loadScatterChart(QXmlStreamReader &reader);
    void loadStockChart(QXmlStreamReader &reader);
    void loadRadarChart(QXmlStreamReader &reader);
    void readBandFormats(QXmlStreamReader &reader);
    void readDropLines(QXmlStreamReader &reader, ShapeFormat &shape);

    void saveAreaChart(QXmlStreamWriter &writer) const;
    void saveSurfaceChart(QXmlStreamWriter &writer) const;
    void saveBubbleChart(QXmlStreamWriter &writer) const;
    void savePieChart(QXmlStreamWriter &writer) const;
    void saveLineChart(QXmlStreamWriter &writer) const;
    void saveBarChart(QXmlStreamWriter &writer) const;
    void saveScatterChart(QXmlStreamWriter &writer) const;
    void saveStockChart(QXmlStreamWriter &writer) const;
    void saveRadarChart(QXmlStreamWriter &writer) const;
    void saveBandFormats(QXmlStreamWriter &writer) const;
};

    class ChartPrivate : public AbstractOOXmlFilePrivate
    {
        Q_DECLARE_PUBLIC(Chart)

    public:
        ChartPrivate(Chart *q, Chart::CreateFlag flag);
        ~ChartPrivate();
        ChartPrivate &operator=(const ChartPrivate &other);
    public:
        bool loadXmlChart(QXmlStreamReader &reader);
        bool loadXmlPlotArea(QXmlStreamReader &reader);
        void addAxis(Axis::Type type, Axis::Position pos, QString title);
        void addAxis(const Axis &axis);
    public:
        void saveXmlChart(QXmlStreamWriter &writer) const;
    public:
        ///CT_ChartSpace properties
        std::optional<bool> date1904;
        QString language;
        std::optional<bool> roundedCorners;
        std::optional<int> styleID; //1..48
        //    std::optional<ColorMapping> colorMappingOverwrite;
        //    std::optional<PivotSource> pivotSource;
        //    std::optional<Protection> protection;
        ShapeFormat chartSpaceShape;
        TextFormat textProperties;
        //    std::optional<ExternalData> externalData;
        HeaderFooter headerFooter;
        PageMargins pageMargins;
        PageSetup pageSetup;
        //    std::optional<Reference> userShapes;
        ExtensionList chartSpaceExtList;

        ///CT_Chart properties
        Title title;
        std::optional<bool> autoTitleDeleted;
        //    QList<PivotFormat> pivotFormats;
        //    std::optional<View3D> view3D;
        //    std::optional<Surface> floor;
        //    std::optional<Surface> sideWall;
        //    std::optional<Surface> backWall;
        Legend legend;
        std::optional<bool> plotVisOnly;
        std::optional<Chart::DisplayBlanksAs> dispBlanksAs;
        std::optional<bool> showDLblsOverMax;
        ExtensionList chartExtList;

        ///CT_PlotArea properties
        Layout layout;
        QList<SubChart> subcharts;
        QList<Axis> axisList;
        std::optional<DataTable> dTable;
        ShapeFormat plotAreaShape;
        ExtensionList plotAreaExtList;

        AbstractSheet* sheet;
    private:
        friend class Axis;
    };

ChartPrivate::ChartPrivate(Chart *q, Chart::CreateFlag flag)
    : AbstractOOXmlFilePrivate(q, flag)
{

}

ChartPrivate::~ChartPrivate()
{
}

ChartPrivate &ChartPrivate::operator=(const ChartPrivate &other)
{
    date1904 = other.date1904;
    language = other.language;
    roundedCorners = other.roundedCorners;
    styleID = other.styleID; //1..48
    //colorMappingOverwrite = other.colorMappingOverwrite;
    //pivotSource = other.pivotSource;
    //protection = other.protection;
    chartSpaceShape = other.chartSpaceShape;
    textProperties = other.textProperties;
    //externalData = other.externalData;
    headerFooter = other.headerFooter;
    pageMargins = other.pageMargins;
    pageSetup = other.pageSetup;
    //userShapes = other.userShapes;
    chartSpaceExtList = other.chartSpaceExtList;

    title = other.title;
    autoTitleDeleted = other.autoTitleDeleted;
    //pivotFormats = other.pivotFormats;
    //view3D = other.view3D;
    //floor = other.floor;
    //sideWall = other.sideWall;
    //backWall = other.backWall;
    legend = other.legend;
    plotVisOnly = other.plotVisOnly;
    dispBlanksAs = other.dispBlanksAs;
    showDLblsOverMax = other.showDLblsOverMax;
    chartExtList = other.chartExtList;

    layout = other.layout;
    subcharts = other.subcharts;
    axisList = other.axisList;
    dTable = other.dTable;
    plotAreaShape = other.plotAreaShape;
    plotAreaExtList = other.plotAreaExtList;
    return *this;
}

Chart::Chart(AbstractSheet *parent, CreateFlag flag)
    : AbstractOOXmlFile(new ChartPrivate(this, flag))
{
    Q_D(Chart);

    d->sheet = parent;
}

Chart &Chart::operator=(const Chart &other)
{
    Q_D(Chart);
    *d = *other.d_func();
    return *this;
}

Chart::~Chart()
{
}

QList<int> Chart::addSeries(const CellRange &range, AbstractSheet *sheet,
                      bool firstRowContainsHeaders,
                      bool firstColumnContainsCategoryData,
                      bool columnBased,
                      int subchart)
{
    Q_D(Chart);

    if (!range.isValid())
        return {};
    if (sheet && sheet->type() != AbstractSheet::Type::Worksheet)
        return {};
    if (!sheet && d->sheet->type() != AbstractSheet::Type::Worksheet)
        return {};

    QString sheetName = sheet ? sheet->name() : d->sheet->name();
    //In case sheetName contains space or '
    sheetName = escapeSheetName(sheetName);

    QList<int> result;

    if (range.columnCount() == 1 || range.rowCount() == 1) {
        auto seriesIdx = addSeries(subchart);
        auto series = this->series(seriesIdx);
        if (series) {
            series->setValueData(sheetName + QLatin1String("!") + range.toString(true, true));
            result << seriesIdx;
        }
    }
    else if (columnBased) {
        //Column based series, first column is category data, rest are value data
        int firstDataRow = range.firstRow();
        int firstDataColumn = range.firstColumn();

        if (firstRowContainsHeaders)
            firstDataRow += 1;


        CellRange categorySubRange(firstDataRow,
                           range.firstColumn(),
                           range.lastRow(),
                           range.firstColumn());
        QString categoryReference = sheetName + QLatin1String("!") + categorySubRange.toString(true, true);
        if (firstColumnContainsCategoryData)
            firstDataColumn += 1;
        for (int col = firstDataColumn; col <= range.lastColumn(); ++col) {
            CellRange subRange(firstDataRow, col, range.lastRow(), col);
            auto seriesIdx = addSeries(subchart);
            auto series = this->series(seriesIdx);
            if (series) {
                if (firstColumnContainsCategoryData)
                    series->setCategoryData(categoryReference);
                series->setValueData(sheetName + QLatin1String("!") + subRange.toString(true, true));

                if (firstRowContainsHeaders) {
                    CellRange subRange(range.firstRow(), col, range.firstRow(), col);
                    series->setNameReference(sheetName + QLatin1String("!") + subRange.toString(true, true));
                }
                result << seriesIdx;
            }
        }
    }
    else {
        //Row based series
        int firstDataRow = range.firstRow();
        int firstDataColumn = range.firstColumn();

        if (firstRowContainsHeaders)
            firstDataColumn += 1;

        CellRange categorySubRange(range.firstRow(), firstDataColumn, range.firstRow(), range.lastColumn());
        QString categoryReference = sheetName + QLatin1String("!") + categorySubRange.toString(true, true);
        if (firstColumnContainsCategoryData)
            firstDataRow += 1;
        for (int row = firstDataRow; row <= range.lastRow(); ++row) {
            CellRange subRange(row, firstDataColumn, row, range.lastColumn());
            auto seriesIdx = addSeries(subchart);
            auto series = this->series(seriesIdx);
            if (series) {
                if (firstColumnContainsCategoryData)
                    series->setCategoryData(categoryReference);
                series->setValueData(sheetName + QLatin1String("!") + subRange.toString(true, true));

                if (firstRowContainsHeaders) {
                    CellRange subRange(row, range.firstColumn(), row, range.firstColumn());
                    series->setNameReference(sheetName + QLatin1String("!") + subRange.toString(true, true));
                }
                result << seriesIdx;
            }
        }
    }
    return result;
}

int Chart::addSeries(const CellRange &keyRange, const CellRange &valRange,
                                AbstractSheet *sheet, bool keyRangeIncludesHeader, int subchart)
{
    Q_D(Chart);

    if (!keyRange.isValid() || !valRange.isValid())
        return {};
    if (sheet && sheet->type() != AbstractSheet::Type::Worksheet)
        return {};
    if (!sheet && d->sheet->type() != AbstractSheet::Type::Worksheet)
        return {};

    QString sheetName = sheet ? sheet->name() : d->sheet->name();
    //In case sheetName contains space or '
    sheetName = escapeSheetName(sheetName);

    if (!(keyRange.columnCount() == 1 || keyRange.rowCount() == 1))
        return {};
    if (!(valRange.columnCount() == 1 || valRange.rowCount() == 1))
        return {};

    auto seriesIdx = addSeries(subchart);
    auto series = this->series(seriesIdx);
    if (!series) return -1;

    CellRange subRange = keyRange;
    if (keyRange.columnCount() == 1) {
        //Column based series
        if (keyRangeIncludesHeader) subRange.setFirstRow(subRange.firstRow()+1);
    }
    if (keyRange.rowCount() == 1) {
        //Row based series
        if (keyRangeIncludesHeader) subRange.setFirstColumn(subRange.firstColumn()+1);
    }
    series->setCategoryData(sheetName + QLatin1String("!") + subRange.toString(true, true));
    subRange = valRange;
    if (valRange.columnCount() == 1) {
        //Column based series
        if (keyRangeIncludesHeader) subRange.setFirstRow(subRange.firstRow()+1);
    }
    if (valRange.rowCount() == 1) {
        //Row based series
        if (keyRangeIncludesHeader) subRange.setFirstColumn(subRange.firstColumn()+1);
    }
    series->setValueData(sheetName + QLatin1String("!") + subRange.toString(true, true));

    if(keyRangeIncludesHeader) {
        CellRange subRange(valRange.firstRow(), valRange.firstColumn(), valRange.firstRow(), valRange.firstColumn());
        series->setNameReference(sheetName + QLatin1String("!") + subRange.toString(true, true));
    }
    return seriesCount()-1;
}

int Chart::addSeries(int subchart)
{
    Q_D(Chart);

    if (subchart < 0 || subchart >= d->subcharts.size()) return -1;

    d->subcharts[subchart].addSeries(seriesCount());
    return seriesCount()-1;
}

Series* Chart::series(int index)
{
    Q_D(Chart);

    for (auto &sub: d->subcharts) {
        for (auto &ser: sub.seriesList)
            if (ser.index() == index) return &ser;
    }

    return nullptr;
}

Series* Chart::seriesByOrder(int order)
{
    Q_D(Chart);
    for (auto &sub: d->subcharts) {
        for (auto &ser: sub.seriesList)
            if (ser.order() == order) return &ser;
    }

    return nullptr;
}

bool Chart::hasAxis(int id) const
{
    Q_D(const Chart);

    for (const auto &ax: qAsConst(d->axisList)) {
        if (ax.id() == id) return true;
    }
    return false;
}

int Chart::seriesCount() const
{
    Q_D(const Chart);

    int count = 0;

    for (const auto &sub: qAsConst(d->subcharts))
        count += sub.seriesList.size();

    return count;
}

bool Chart::removeSeries(int index)
{
    Q_D(Chart);
    bool wasRemoved = false;

    if (index >= 0 && index < seriesCount()) {
        for (auto &sub: d->subcharts) {
            for (int i=sub.seriesList.size()-1; i>=0; --i) {
                if (sub.seriesList[i].index() == index) {
                    sub.seriesList.removeAt(i);
                    wasRemoved = true;
                }
            }
        }
        for (int i=d->subcharts.size()-1; i>=0; --i) {
            if (d->subcharts[i].seriesList.isEmpty())
                d->subcharts.removeAt(i);
        }
    }
    return wasRemoved;
}

bool Chart::removeSeries(const Series &series)
{
    Q_D(Chart);
    bool wasRemoved = false;

    for (auto &sub: d->subcharts) {
        for (int i=sub.seriesList.size()-1; i>=0; --i) {
            if (sub.seriesList[i] == series) {
                sub.seriesList.removeAt(i);
                wasRemoved = true;
            }
        }
    }
    for (int i=d->subcharts.size()-1; i>=0; --i) {
        if (d->subcharts[i].seriesList.isEmpty())
            d->subcharts.removeAt(i);
    }
    return wasRemoved;
}

void Chart::removeAllSeries()
{
    Q_D(Chart);
    for (auto &sub: d->subcharts) sub.seriesList.clear();
}

void Chart::setChartLineFormat(const LineFormat &format)
{
    Q_D(Chart);

    d->chartSpaceShape.setLine(format);
}

void Chart::setPlotAreaLineFormat(const LineFormat &format)
{
    Q_D(Chart);

    d->plotAreaShape.setLine(format);
}

void Chart::setChartShape(const ShapeFormat &shape)
{
    Q_D(Chart);
    d->chartSpaceShape = shape;
}

ShapeFormat &Chart::chartShape()
{
    Q_D(Chart);
    return d->chartSpaceShape;
}

HeaderFooter &Chart::headerFooter()
{
    Q_D(Chart);
    return d->headerFooter;
}

PageMargins &Chart::pageMargins()
{
    Q_D(Chart);
    return d->pageMargins;
}

PageSetup &Chart::pageSetup()
{
    Q_D(Chart);
    return d->pageSetup;
}

void Chart::setPlotAreaShape(const ShapeFormat &shape)
{
    Q_D(Chart);
    d->plotAreaShape = shape;
}

ShapeFormat &Chart::plotAreaShape()
{
    Q_D(Chart);
    return d->plotAreaShape;
}

TextFormat Chart::textProperties() const
{
    Q_D(const Chart);
    return d->textProperties;
}

TextFormat &Chart::textProperties()
{
    Q_D(Chart);
    return d->textProperties;
}

void Chart::setTextProperties(const TextFormat &textProperties)
{
    Q_D(Chart);
    d->textProperties = textProperties;
}

int Chart::addAxis(Axis::Type type, Axis::Position pos, QString title)
{
    Q_D(Chart);

    d->addAxis(type, pos, title);
    return d->axisList.size()-1;
}

int Chart::addAxis(const Axis &axis)
{
    Q_D(Chart);

    d->addAxis(axis);
    return d->axisList.size()-1;
}


Axis *Chart::axis(int idx)
{
    Q_D(Chart);

    if (idx < 0 || idx >= d->axisList.size()) return nullptr;

    return &d->axisList[idx];
}

Axis *Chart::axis(Axis::Position pos)
{
    Q_D(Chart);

    for (auto &ax: d->axisList) {
        if (ax.position() == pos) return &ax;
    }

    return nullptr;
}

QList<Axis> Chart::axes() const
{
    Q_D(const Chart);
    return d->axisList;
}

QList<int> Chart::axesIDs() const
{
    Q_D(const Chart);
    QList<int> result;

    for (const auto &axis: qAsConst(d->axisList)) result << axis.id();
    return result;
}

Axis *Chart::axis(Axis::Type type)
{
    Q_D(Chart);

    for (auto &ax: d->axisList) {
        if (ax.type() == type) return &ax;
    }

    return nullptr;
}

Axis *Chart::axis(Axis::Type type, Axis::Position pos)
{
    Q_D(Chart);

    for (auto &ax: d->axisList) {
        if (ax.type() == type && ax.position() == pos) return &ax;
    }

    return nullptr;
}

bool Chart::removeAxis(int axisID)
{
    Q_D(Chart);

    //search for the axis
    auto ax = std::find_if(d->axisList.begin(), d->axisList.end(), [axisID](const Axis &axis){
        return axis.id() == axisID;
    });
    if (ax == d->axisList.end()) return false;

    //search for the axis in series

    auto ser = std::find_if(d->subcharts.begin(), d->subcharts.end(), [axisID](const SubChart &c){
        return c.axesIds.contains(axisID);
    });
    if (ser != d->subcharts.end()) return false;
    d->axisList.removeOne(*ax);
    return true;
}

bool Chart::removeAxis(Axis *axis)
{
    return removeAxis(axis->id());
}

void Chart::removeAxes()
{
    Q_D(Chart);
    d->axisList.clear();
    for (auto &sub: d->subcharts) {
        sub.axesIds.clear();
    }
}

QList<int> Chart::seriesThatUseAxis(int axisID) const
{
    Q_D(const Chart);
    QList<int> result;

    for (const auto &sub: qAsConst(d->subcharts)) {
        if (sub.axesIds.contains(axisID)) {
            for (const auto &s: sub.seriesList)
                result.append(s.index());
        }
    }

    return result;
}

QList<int> Chart::seriesThatUseAxis(Axis *axis) const
{
    return seriesThatUseAxis(axis->id());
}

int Chart::axesCount() const
{
    Q_D(const Chart);
    return d->axisList.size();
}

void Chart::setTitle(const QString &title)
{
    Q_D(Chart);

    d->title.setPlainString(title);
}

void Chart::setTitle(const Title &title)
{
    Q_D(Chart);
    d->title = title;
}

Title &Chart::title()
{
    Q_D(Chart);
    return d->title;
}

Title Chart::title() const
{
    Q_D(const Chart);
    return d->title;
}

std::optional<bool> Chart::autoTitleDeleted() const
{
    Q_D(const Chart);
    return d->autoTitleDeleted;
}

void Chart::setAutoTitleDeleted(bool value)
{
    Q_D(Chart);
    d->autoTitleDeleted = value;
}

std::optional<bool> Chart::plotOnlyVisibleCells() const
{
    Q_D(const Chart);
    return d->plotVisOnly;
}

void Chart::setPlotOnlyVisibleCells(bool value)
{
    Q_D(Chart);
    d->plotVisOnly = value;
}

std::optional<Chart::DisplayBlanksAs> Chart::displayBlanksAs() const
{
    Q_D(const Chart);
    return d->dispBlanksAs;
}

void Chart::setDisplayBlanksAs(DisplayBlanksAs value)
{
    Q_D(Chart);
    d->dispBlanksAs = value;
}

std::optional<bool> Chart::showDataLabelsOverMaximum() const
{
    Q_D(const Chart);
    return d->showDLblsOverMax;
}

void Chart::setShowDataLabelsOverMaximum(bool value)
{
    Q_D(Chart);
    d->showDLblsOverMax = value;
}

void Chart::setSeriesAxesIDs(Series *series, const QList<int> &axesIds)
{
    Q_D(Chart);

    if (d->subcharts.isEmpty()) return;

    //first we need to find a subchart with the exact axes
    SubChart *subchart = nullptr;
    for (auto &sub: d->subcharts) {
        if (sub.axesIds == axesIds) {
            subchart = &sub;
            break;
        }
    }

    //second we need to find the series in subcharts
    SubChart *oldSubchart = nullptr;
    for (auto &sub: d->subcharts) {
        if (sub.seriesList.contains(*series)) {
            oldSubchart = &sub;
            break;
        }
    }

    //series not found
    if (!oldSubchart) return;

    //if oldsubchart has the same axes do nothing
    if (oldSubchart->axesIds == axesIds) return;

    //create new subchart if necessary
    if (!subchart) {
        SubChart sub(oldSubchart->type);
        d->subcharts << sub;
        sub.axesIds = axesIds;
        subchart = &sub;
    }

    //move the series from the old subchart to the new one
    subchart->seriesList << *series;
    oldSubchart->seriesList.removeOne(*series);
}

void Chart::moveSeries(int oldOrder, int newOrder)
{
    auto s1 = seriesByOrder(oldOrder);
    auto s2 = seriesByOrder(newOrder);
    if (!s1 || !s2) return;
    s1->setOrder(newOrder);
    s2->setOrder(oldOrder);
}

QList<int> Chart::addDefaultAxes(Type type)
{
    QList<int> result;
    switch (type) {
    case Type::Pie:
    case Type::Pie3D:
    case Type::Doughnut:
    case Type::OfPie:
    case Type::None: break;
    case Type::Line:
    case Type::Area:
    case Type::Bubble:
    case Type::Bar:
    case Type::Radar:
    case Type::Stock: {
        auto ax1 = addAxis(Axis::Type::Category, Axis::Position::Bottom);
        auto ax2 = addAxis(Axis::Type::Value, Axis::Position::Left);
        axis(ax1)->setCrossAxis(axis(ax2));
        axis(ax2)->setCrossAxis(axis(ax1));
        result << axis(ax1)->id();
        result << axis(ax2)->id();
        break;
    }
    case Type::Scatter: {
        auto ax1 = addAxis(Axis::Type::Value, Axis::Position::Bottom);
        auto ax2 = addAxis(Axis::Type::Value, Axis::Position::Left);
        axis(ax1)->setCrossAxis(axis(ax2));
        axis(ax2)->setCrossAxis(axis(ax1));
        result << axis(ax1)->id();
        result << axis(ax2)->id();
        break;
    }
    case Type::Area3D:
    case Type::Line3D:
    case Type::Bar3D:
    case Type::Surface:
    case Type::Surface3D: {
        auto ax1 = addAxis(Axis::Type::Category, Axis::Position::Bottom);
        auto ax2 = addAxis(Axis::Type::Value, Axis::Position::Left);
        auto ax3 = addAxis(Axis::Type::Series, Axis::Position::Bottom);
        axis(ax1)->setCrossAxis(axis(ax2));
        axis(ax2)->setCrossAxis(axis(ax1));
        axis(ax3)->setCrossAxis(axis(ax2));
        result << axis(ax1)->id();
        result << axis(ax2)->id();
        result << axis(ax3)->id();
        break;
    }
    }
    return result;
}

void Chart::addSubchart(Type type, const QList<int> &axesIDs)
{
    Q_D(Chart);
    SubChart sub(type);
    if (axesIDs.isEmpty()) {
        auto axes = addDefaultAxes(type);
        sub.axesIds = axes;
    }
    else
        sub.axesIds = axesIDs;

    d->subcharts << sub;
}

void Chart::setSubchartAxes(int subchartIndex, const QList<int> &axesIDs)
{
    Q_D(Chart);
    if (subchartIndex < 0 || subchartIndex >= d->subcharts.size()) return;
    d->subcharts[subchartIndex].axesIds = axesIDs;
}

QList<int> Chart::subchartAxes(int subchartIndex) const
{
    Q_D(const Chart);
    if (subchartIndex < 0 || subchartIndex >= d->subcharts.size()) return {};
    return d->subcharts[subchartIndex].axesIds;
}

bool Chart::subchartHasSeries(int subchartIndex) const
{
    Q_D(const Chart);
    if (subchartIndex < 0 || subchartIndex >= d->subcharts.size()) return false;
    return !d->subcharts[subchartIndex].seriesList.isEmpty();
}

bool Chart::removeSubchart(int subchartIndex)
{
    Q_D(Chart);

    if (subchartIndex < 0 || subchartIndex >= d->subcharts.size()) return false;
    d->subcharts.removeAt(subchartIndex);
    return true;
}

int Chart::subchartsCount() const
{
    Q_D(const Chart);
    return d->subcharts.size();
}

Chart::Type Chart::subchartType(int subchartIndex) const
{
    Q_D(const Chart);
    if (subchartIndex < 0 || subchartIndex >= d->subcharts.size()) return Chart::Type::None;
    return d->subcharts[subchartIndex].type;
}

void Chart::setSubchartType(int subchartIndex, Type type)
{
    Q_D(Chart);
    if (subchartIndex < 0 || subchartIndex >= d->subcharts.size()) return;
    d->subcharts[subchartIndex].type = type;
}

void Chart::setDefaultLegend()
{
    Q_D(Chart);
    d->legend = Legend::defaultLegend();
}

void Chart::setLegend(Legend::Position position, bool overlay)
{
    Q_D(Chart);

    d->legend.setPosition(position);
    d->legend.setOverlay(overlay);
}

Legend &Chart::legend()
{
    Q_D(Chart);
    return d->legend;
}

void Chart::removeLegend()
{
    Q_D(Chart);
    d->legend = {};
}

Layout Chart::layout() const
{
    Q_D(const Chart);
    return d->layout;
}

Layout &Chart::layout()
{
    Q_D(Chart);
    return d->layout;
}

void Chart::setLayout(const Layout &layout)
{
    Q_D(Chart);
    d->layout = layout;
}

void Chart::saveToXmlFile(QIODevice *device) const
{
    Q_D(const Chart);

    QXmlStreamWriter writer(device);

    writer.writeStartDocument(QStringLiteral("1.0"), true);
    writer.writeStartElement(QStringLiteral("c:chartSpace"));
    writer.writeAttribute(QStringLiteral("xmlns:c"),
                          QStringLiteral("http://schemas.openxmlformats.org/drawingml/2006/chart"));
    writer.writeAttribute(QStringLiteral("xmlns:a"),
                          QStringLiteral("http://schemas.openxmlformats.org/drawingml/2006/main"));
    writer.writeAttribute(QStringLiteral("xmlns:r"),
                          QStringLiteral("http://schemas.openxmlformats.org/officeDocument/2006/relationships"));

    // 1. date1904
    writeEmptyElement(writer, QLatin1String("c:date1904"), d->date1904);
    // 2. lang
    writeEmptyElement(writer, QLatin1String("c:lang"), d->language);
    // 3. roundedCorners
    writeEmptyElement(writer, QLatin1String("c:roundedCorners"), d->roundedCorners);
    // 4. style
    writeEmptyElement(writer, QLatin1String("c:style"), d->styleID);
    // 5. clrMapOvr
    // 6. pivotSource
    // 7. protection
    // 8. chart
    d->saveXmlChart(writer);
    // 9. spPr
    if (d->chartSpaceShape.isValid()) d->chartSpaceShape.write(writer, "c:spPr");
    // 10. txPr
    if (d->textProperties.isValid()) d->textProperties.write(writer, QLatin1String("c:txPr"));
    // 11. externalData
    // 12. printSettings
    if (d->headerFooter.isValid() || d->pageMargins.isValid() || d->pageSetup.isValid()) {
        writer.writeStartElement(QLatin1String("c:printSettings"));
        d->headerFooter.write(writer, QLatin1String("c:headerFooter"));
        d->pageMargins.write(writer);
        d->pageSetup.writeChartsheet(writer);
    }
    // 13. userShapes
    // 14. extLst
    if (d->chartSpaceExtList.isValid()) d->chartSpaceExtList.write(writer, QLatin1String("c:extLst"));
    writer.writeEndElement();// c:chartSpace
    writer.writeEndDocument();
}

bool Chart::loadFromXmlFile(QIODevice *device)
{
    Q_D(Chart);

    QXmlStreamReader reader(device);
    while (!reader.atEnd()) {
        auto token = reader.readNext();
        if (token == QXmlStreamReader::StartElement) {
            const auto &a = reader.attributes();
            if (reader.name() == QLatin1String("chartSpace")) {
                //dive into
            }
            else if (reader.name() == QLatin1String("chart")) {
                if (!d->loadXmlChart(reader))
                    return false;
            }
            else if (reader.name() == QLatin1String("date1904")) {
                parseAttributeBool(a, QLatin1String("val"), d->date1904);
            }
            else if (reader.name() == QLatin1String("lang")) {
                d->language = a.value(QLatin1String("val")).toString();
            }
            else if (reader.name() == QLatin1String("roundedCorners")) {
                parseAttributeBool(a, QLatin1String("val"), d->roundedCorners);
            }
            else if (reader.name() == QLatin1String("style")) {
                parseAttributeInt(a, QLatin1String("val"), d->styleID);
            }
            else if (reader.name() == QLatin1String("clrMapOvr")) {
                //TODO
                reader.skipCurrentElement();
            }
            else if (reader.name() == QLatin1String("pivotSource")) {
                //TODO
                reader.skipCurrentElement();
            }
            else if (reader.name() == QLatin1String("protection")) {
                //TODO
                reader.skipCurrentElement();
            }
            else if (reader.name() == QLatin1String("spPr"))
                d->chartSpaceShape.read(reader);
            else if (reader.name() == QLatin1String("txPr"))
                d->textProperties.read(reader);
            else if (reader.name() == QLatin1String("externalData")) {
                //TODO
                reader.skipCurrentElement();
            }
            else if (reader.name() == QLatin1String("printSettings")) {//dive into
            }
            else if (reader.name() == QLatin1String("headerFooter"))
                d->headerFooter.read(reader);
            else if (reader.name() == QLatin1String("pageMargins"))
                d->pageMargins.read(reader);
            else if (reader.name() == QLatin1String("pageSetup"))
                d->pageSetup.read(reader);
            else if (reader.name() == QLatin1String("userShapes")) {
                //TODO
                reader.skipCurrentElement();
            }
            else if (reader.name() == QLatin1String("extLst"))
                d->chartSpaceExtList.read(reader);
            else
                reader.skipCurrentElement();
        }
    }

    //relate mediafiles


    return true;
}


bool ChartPrivate::loadXmlChart(QXmlStreamReader &reader)
{
    const auto &name = reader.name();
    while (!reader.atEnd()) {
        auto token = reader.readNext();
        if (token == QXmlStreamReader::StartElement) {
            const auto &a = reader.attributes();

            if (reader.name() == QLatin1String("plotArea")) {
                if (!loadXmlPlotArea(reader))
                    return false;
            }
            else if (reader.name() == QLatin1String("title"))
                title.read(reader);
            else if (reader.name() == QLatin1String("autoTitleDeleted"))
                parseAttributeBool(a, QLatin1String("val"), autoTitleDeleted);
            else if (reader.name() == QLatin1String("pivotFmts")) {
                //TODO
                reader.skipCurrentElement();
            }
            else if (reader.name() == QLatin1String("view3D")) {
                //TODO
                reader.skipCurrentElement();
            }
            else if (reader.name() == QLatin1String("floor")) {
                //TODO
                reader.skipCurrentElement();
            }
            else if (reader.name() == QLatin1String("sideWall")) {
                //TODO
                reader.skipCurrentElement();
            }
            else if (reader.name() == QLatin1String("backWall")) {
                //TODO
                reader.skipCurrentElement();
            }
            else if (reader.name() == QLatin1String("legend")) {
                legend.read(reader);
            }
            else if (reader.name() == QLatin1String("plotVisOnly"))
                parseAttributeBool(a, QLatin1String("val"), plotVisOnly);
            else if (reader.name() == QLatin1String("dispBlanksAs")) {
                 Chart::DisplayBlanksAs d;
                 Chart::fromString(a.value(QLatin1String("val")).toString(), d);
                 dispBlanksAs = d;
            }
            else if (reader.name() == QLatin1String("showDLblsOverMax"))
                parseAttributeBool(a, QLatin1String("val"), showDLblsOverMax);
            else if (reader.name() == QLatin1String("extLst"))
                chartExtList.read(reader);
            else reader.skipCurrentElement();
        }
        else if (token == QXmlStreamReader::EndElement && reader.name() == name)
            break;
    }
    return true;
}

bool ChartPrivate::loadXmlPlotArea(QXmlStreamReader &reader)
{
    const auto &name = reader.name();
    while (!reader.atEnd()) {
        auto token = reader.readNext();
        if (token == QXmlStreamReader::StartElement) {
            if (reader.name() == QLatin1String("layout")) {
                layout.read(reader);
            }
            else if (reader.name().endsWith(QLatin1String("Chart"))) {
                SubChart c(Chart::Type::None);
                if (c.read(reader))
                    subcharts << c;
                else {
                    qDebug() << "[debug] failed to load subchart";
                    return false;
                }
            }
            else if (reader.name().endsWith(QLatin1String("Ax"))) {
                Axis ax; ax.read(reader);
                axisList << ax;
            }
            else if (reader.name() == QLatin1String("dTable")) {
                DataTable t; t.read(reader);
                dTable = t;
            }
            else if (reader.name() == QLatin1String("spPr")) {
                plotAreaShape.read(reader);
            }
            else if (reader.name() == QLatin1String("extLst")) {
                plotAreaExtList.read(reader);
            }
            else reader.skipCurrentElement();
        }
        else if (token == QXmlStreamReader::EndElement && reader.name() == name)
            break;
    }

    return true;
}

void ChartPrivate::addAxis(Axis::Type type, Axis::Position pos, QString title)
{
    int axisId = 1;

    auto hasAxis = [this](int id) -> bool {
        for (const auto &ax: qAsConst(axisList)) if (ax.id() == id) return true;
        return false;
    };

    while (hasAxis(axisId)) axisId++;

    auto axis = Axis(type, pos);
    axis.setId(axisId);
    if (!title.isEmpty()) axis.setTitle(title);

    axisList.append(axis);
}

void ChartPrivate::addAxis(const Axis &axis)
{
    int axisId = 1;

    auto hasAxis = [this](int id) -> bool {
        for (const auto &ax: qAsConst(axisList)) if (ax.id() == id) return true;
        return false;
    };

    while (hasAxis(axisId)) axisId++;
    axisList.append(axis);
    axisList.last().setId(axisId);
}

void SubChart::loadAreaChart(QXmlStreamReader &reader)
{
    const auto &a = reader.attributes();
    if (reader.name() == QLatin1String("grouping")) {
        Chart::Grouping g; Chart::fromString(a.value(QLatin1String("val")).toString(), g);
        grouping = g;
    }
    else if (reader.name() == QLatin1String("varyColors"))
        parseAttributeBool(a, QLatin1String("val"), varyColors);
    else if (reader.name() == QLatin1String("ser")) {
        Series ser(Series::Type::Area); ser.read(reader);
        seriesList << ser;
    }
    else if (reader.name() == QLatin1String("dLbls"))
        labels.read(reader);
    else if (reader.name() == QLatin1String("dropLines")) {
        ShapeFormat f;
        readDropLines(reader, f);
        dropLines = f;
    }
    else if (reader.name() == QLatin1String("axId"))
        axesIds << a.value(QLatin1String("val")).toInt();
    else if (reader.name() == QLatin1String("gapDepth"))
        parseAttributePercent(a, QLatin1String("val"), gapDepth);
    else reader.skipCurrentElement();
}

void SubChart::loadSurfaceChart(QXmlStreamReader &reader)
{
    const auto &a = reader.attributes();
    if (reader.name() == QLatin1String("wireframe"))
        parseAttributeBool(a, QLatin1String("val"), wireframe);
    else if (reader.name() == QLatin1String("ser")) {
        Series ser(Series::Type::Surface); ser.read(reader);
        seriesList << ser;
    }
    else if (reader.name() == QLatin1String("bandFmts"))
        readBandFormats(reader);
    else if (reader.name() == QLatin1String("axId"))
        axesIds << a.value(QLatin1String("val")).toInt();
    else  reader.skipCurrentElement();
}

void SubChart::loadBubbleChart(QXmlStreamReader &reader)
{
    const auto &a = reader.attributes();
    if (reader.name() == QLatin1String("varyColors"))
        parseAttributeBool(a, QLatin1String("val"), varyColors);
    else if (reader.name() == QLatin1String("ser")) {
        Series ser(Series::Type::Bubble); ser.read(reader);
        seriesList << ser;
    }
    else if (reader.name() == QLatin1String("dLbls"))
        labels.read(reader);
    else if (reader.name() == QLatin1String("axId"))
        axesIds << a.value(QLatin1String("val")).toInt();
    else if (reader.name() == QLatin1String("bubble3D"))
        parseAttributeBool(a, QLatin1String("val"), bubble3D);
    else if (reader.name() == QLatin1String("bubbleScale"))
        parseAttributePercent(a, QLatin1String("val"), bubbleScale);
    else if (reader.name() == QLatin1String("showNegBubbles"))
        parseAttributeBool(a, QLatin1String("val"), showNegBubbles);
    else if (reader.name() == QLatin1String("sizeRepresents")) {
        auto val = a.value(QLatin1String("val"));
        if (val == QLatin1String("area")) bubbleSizeRepresents = Chart::BubbleSizeRepresents::Area;
        if (val == QLatin1String("w")) bubbleSizeRepresents = Chart::BubbleSizeRepresents::Width;
    }
    else  reader.skipCurrentElement();
}

void SubChart::loadPieChart(QXmlStreamReader &reader)
{
    const auto &a = reader.attributes();
    if (reader.name() == QLatin1String("varyColors"))
        parseAttributeBool(a, QLatin1String("val"), varyColors);
    else if (reader.name() == QLatin1String("ser")) {
        Series ser(Series::Type::Pie); ser.read(reader);
        seriesList << ser;
    }
    else if (reader.name() == QLatin1String("dLbls"))
        labels.read(reader);
    else if (reader.name() == QLatin1String("axId"))
        axesIds << a.value(QLatin1String("val")).toInt();
    else if (reader.name() == QLatin1String("firstSliceAng"))
        parseAttributeInt(a, QLatin1String("val"), firstSliceAng);
    else if (reader.name() == QLatin1String("gapWidth"))
        parseAttributePercent(a, QLatin1String("val"), gapWidth);
    else if (reader.name() == QLatin1String("splitType")) {
        Chart::SplitType t; Chart::fromString(a.value(QLatin1String("val")).toString(), t);
        splitType = t;
    }
    else if (reader.name() == QLatin1String("splitPos"))
        parseAttributeDouble(a, QLatin1String("val"), splitPos);
    else if (reader.name() == QLatin1String("secondPiePt")) {
        int i;
        parseAttributeInt(a, QLatin1String("val"), i);
        customSplit << i;
    }
    else if (reader.name() == QLatin1String("secondPieSize"))
        parseAttributePercent(a, QLatin1String("val"), secondPieSize);
    else if (reader.name() == QLatin1String("serLines")) {
        reader.readNextStartElement();
        if (reader.name() == QLatin1String("spPr")) {
            ShapeFormat sh;
            sh.read(reader);
            serLines << sh;
        }
    }
    else if (reader.name() == QLatin1String("holeSize"))
        parseAttributePercent(a, QLatin1String("val"), holeSize);
    else if (reader.name() == QLatin1String("ofPieType")) {
        auto s = a.value(QLatin1String("val"));
        if (s == QLatin1String("pie")) ofPieType = Chart::OfPieType::Pie;
        if (s == QLatin1String("bar")) ofPieType = Chart::OfPieType::Bar;
    }
    else reader.skipCurrentElement();
}

void SubChart::loadLineChart(QXmlStreamReader &reader)
{
    const auto &a = reader.attributes();
    if (reader.name() == QLatin1String("grouping")) {
        Chart::Grouping g; Chart::fromString(a.value(QLatin1String("val")).toString(), g);
        grouping = g;
    }
    else if (reader.name() == QLatin1String("varyColors"))
        parseAttributeBool(a, QLatin1String("val"), varyColors);
    else if (reader.name() == QLatin1String("ser")) {
        Series ser(Series::Type::Line); ser.read(reader);
        seriesList << ser;
    }
    else if (reader.name() == QLatin1String("dLbls"))
        labels.read(reader);
    else if (reader.name() == QLatin1String("dropLines")) {
        ShapeFormat f;
        readDropLines(reader, f);
        dropLines = f;
    }
    else if (reader.name() == QLatin1String("hiLowLines")) {
        ShapeFormat f;
        readDropLines(reader, f);
        hiLowLines = f;
    }
    else if (reader.name() == QLatin1String("upDownBars"))
        upDownBars.read(reader);
    else if (reader.name() == QLatin1String("marker"))
        parseAttributeBool(a, QLatin1String("val"), marker);
    else if (reader.name() == QLatin1String("smooth"))
        parseAttributeBool(a, QLatin1String("val"), smooth);
    else if (reader.name() == QLatin1String("axId"))
        axesIds << a.value(QLatin1String("val")).toInt();
    else if (reader.name() == QLatin1String("gapDepth"))
        parseAttributePercent(a, QLatin1String("val"), gapDepth);
    else reader.skipCurrentElement();
}

void SubChart::loadBarChart(QXmlStreamReader &reader)
{
    const auto &a = reader.attributes();
    if (reader.name() == QLatin1String("grouping")) {
        Chart::BarGrouping g; Chart::fromString(a.value(QLatin1String("val")).toString(), g);
        barGrouping = g;
    }
    else if (reader.name() == QLatin1String("varyColors"))
        parseAttributeBool(a, QLatin1String("val"), varyColors);
    else if (reader.name() == QLatin1String("ser")) {
        Series ser(Series::Type::Bar); ser.read(reader);
        seriesList << ser;
    }
    else if (reader.name() == QLatin1String("dLbls"))
        labels.read(reader);
    else if (reader.name() == QLatin1String("barDir")) {
        auto s = a.value(QLatin1String("val"));
        if (s == QLatin1String("bar")) barDirection = Chart::BarDirection::Bar;
        if (s == QLatin1String("col")) barDirection = Chart::BarDirection::Column;
    }
    else if (reader.name() == QLatin1String("overlap"))
        parseAttributePercent(a, QLatin1String("val"), overlap);
    else if (reader.name() == QLatin1String("shape")) {
        Series::BarShape b; Series::fromString(a.value(QLatin1String("val")).toString(), b);
        barShape = b;
    }
    else if (reader.name() == QLatin1String("axId"))
        axesIds << a.value(QLatin1String("val")).toInt();
    else if (reader.name() == QLatin1String("gapDepth"))
        parseAttributePercent(a, QLatin1String("val"), gapDepth);
    else if (reader.name() == QLatin1String("gapWidth"))
        parseAttributePercent(a, QLatin1String("val"), gapWidth);
    else if (reader.name() == QLatin1String("serLines")) {
        reader.readNextStartElement();
        if (reader.name() == QLatin1String("spPr")) {
            ShapeFormat sh;
            sh.read(reader);
            serLines << sh;
        }
    }
    else reader.skipCurrentElement();
}

void SubChart::loadScatterChart(QXmlStreamReader &reader)
{
    const auto &a = reader.attributes();
    if (reader.name() == QLatin1String("scatterStyle")) {
        Chart::ScatterStyle g; Chart::fromString(a.value(QLatin1String("val")).toString(), g);
        scatterStyle = g;
    }
    else if (reader.name() == QLatin1String("varyColors"))
        parseAttributeBool(a, QLatin1String("val"), varyColors);
    else if (reader.name() == QLatin1String("ser")) {
        Series ser(Series::Type::Scatter); ser.read(reader);
        seriesList << ser;
    }
    else if (reader.name() == QLatin1String("dLbls"))
        labels.read(reader);
    else if (reader.name() == QLatin1String("axId"))
        axesIds << a.value(QLatin1String("val")).toInt();
    else reader.skipCurrentElement();
}

void SubChart::loadStockChart(QXmlStreamReader &reader)
{
    const auto &a = reader.attributes();
    if (reader.name() == QLatin1String("ser")) {
        Series ser(Series::Type::Line); ser.read(reader);
        seriesList << ser;
    }
    else if (reader.name() == QLatin1String("dropLines")) {
        ShapeFormat f;
        readDropLines(reader, f);
        dropLines = f;
    }
    else if (reader.name() == QLatin1String("hiLowLines")) {
        ShapeFormat f;
        readDropLines(reader, f);
        hiLowLines = f;
    }
    else if (reader.name() == QLatin1String("dLbls"))
        labels.read(reader);
    else if (reader.name() == QLatin1String("upDownBars"))
        upDownBars.read(reader);
    else if (reader.name() == QLatin1String("axId"))
        axesIds << a.value(QLatin1String("val")).toInt();
    else reader.skipCurrentElement();
}

void SubChart::loadRadarChart(QXmlStreamReader &reader)
{
    const auto &a = reader.attributes();
    if (reader.name() == QLatin1String("radarStyle")) {
        Chart::RadarStyle g; Chart::fromString(a.value(QLatin1String("val")).toString(), g);
        radarStyle = g;
    }
    else if (reader.name() == QLatin1String("varyColors"))
        parseAttributeBool(a, QLatin1String("val"), varyColors);
    else if (reader.name() == QLatin1String("ser")) {
        Series ser(Series::Type::Radar); ser.read(reader);
        seriesList << ser;
    }
    else if (reader.name() == QLatin1String("dLbls"))
        labels.read(reader);
    else if (reader.name() == QLatin1String("axId"))
        axesIds << a.value(QLatin1String("val")).toInt();
    else if (reader.name() == QLatin1String("extLst")) {
        reader.skipCurrentElement();
    }
}

void ChartPrivate::saveXmlChart(QXmlStreamWriter &writer) const
{
    writer.writeStartElement(QStringLiteral("c:chart"));
    // 1. title
    if (title.isValid()) title.write(writer, QLatin1String("c:title"));
    // 2. autoTitleDeleted
    writeEmptyElement(writer, QLatin1String("c:autoTitleDeleted"), autoTitleDeleted);
    // 3. pivotFmts
    // 4. view3D
    // 5. floor
    // 6. sideWall
    // 7. backWall
    // 8. plotArea
    writer.writeStartElement(QStringLiteral("c:plotArea"));

    layout.write(writer, QLatin1String("c:layout"));

    for (const auto &sub: qAsConst(subcharts)) {
        sub.write(writer);
    }
    for (auto &axis: axisList) {
        axis.write(writer);
    }

    if (dTable.has_value()) dTable->write(writer, QLatin1String("c:dTable"));
    if (plotAreaExtList.isValid()) plotAreaExtList.write(writer, QLatin1String("c:extLst"));

    if (plotAreaShape.isValid()) plotAreaShape.write(writer, "c:spPr");
    writer.writeEndElement(); // c:plotArea

    // 9. legend
    if (legend.isValid()) legend.write(writer);
    // 10. plotVisOnly
    writeEmptyElement(writer, QLatin1String("c:plotVisOnly"), plotVisOnly);
    // 11. dispBlanksAs
    if (dispBlanksAs.has_value())
        writeEmptyElement(writer, QLatin1String("c:dispBlanksAs"), Chart::toString(dispBlanksAs.value()));
    // 12. showDLblsOverMax
    writeEmptyElement(writer, QLatin1String("c:showDLblsOverMax"), showDLblsOverMax);
    // 13. extLst
    chartExtList.write(writer, QLatin1String("c:extLst"));
    writer.writeEndElement(); // c:chart
}

void SubChart::readBandFormats(QXmlStreamReader &reader)
{
    reader.readNextStartElement();

    auto readFmt = [](QXmlStreamReader &reader) -> QPair<int, ShapeFormat> {
        int idx = -1;
        ShapeFormat sh;
        const auto& name = reader.name();
        while (!reader.atEnd()) {
            auto token = reader.readNext();
            if (token == QXmlStreamReader::StartElement) {
                if (reader.name() == QLatin1String("idx"))
                    parseAttributeInt(reader.attributes(), QLatin1String("val"), idx);
                else if (reader.name() == QLatin1String("spPr"))
                    sh.read(reader);
                else if (reader.name() == QLatin1String("extLst")) {
                    reader.skipCurrentElement();
                }
            }
            else if (token == QXmlStreamReader::EndElement && reader.name() == name)
                break;
        }
        return {idx, sh};
    };

    const auto& name = reader.name();
    while (!reader.atEnd()) {
        auto token = reader.readNext();
        if (token == QXmlStreamReader::StartElement) {
             if (reader.name() == QLatin1String("bandFmt")) {
                 auto p = readFmt(reader);
                 bandFormats.insert(p.first, p.second);
             }
        }
        else if (token == QXmlStreamReader::EndElement && reader.name() == name)
            break;
    }
}

void SubChart::readDropLines(QXmlStreamReader &reader, ShapeFormat &shape)
{
    const auto &name = reader.name();
    while (!reader.atEnd()) {
        auto token = reader.readNext();
        if (token == QXmlStreamReader::StartElement) {
            if (reader.name() == QLatin1String("spPr"))
                shape.read(reader);
            else reader.skipCurrentElement();
        }
        else if (token == QXmlStreamReader::EndElement && reader.name() == name)
            break;
    }
}

void SubChart::saveAreaChart(QXmlStreamWriter &writer) const
{
    writer.writeStartElement(type == Chart::Type::Area3D ? QLatin1String("c:area3DChart") : QLatin1String("c:areaChart"));
    if (grouping.has_value())
        writeEmptyElement(writer, QLatin1String("c:grouping"), Chart::toString(grouping.value()));
    writeEmptyElement(writer, QLatin1String("c:varyColors"), varyColors);
    for (const auto &ser: qAsConst(seriesList)) ser.write(writer);
    if (labels.isValid()) labels.write(writer);
    if (dropLines.has_value()) {
        if (dropLines->isValid()) {
            writer.writeStartElement(QLatin1String("c:dropLines"));
            dropLines->write(writer, QLatin1String("c:spPr"));
            writer.writeEndElement();
        }
        else writer.writeEmptyElement(QLatin1String("c:dropLines"));
    }
    if (axesIds.isEmpty()) {
        qDebug()<<"[Error] Axes IDs are empty!";
        return;
    }
    for (const auto id: qAsConst(axesIds)) writeEmptyElement(writer, QLatin1String("c:axId"), id);
    if (gapDepth.has_value() && type == Chart::Type::Area3D)
        writeEmptyElement(writer, QLatin1String("c:gapDepth"), toST_PercentInt(gapDepth.value()));
    writer.writeEndElement();
}

void SubChart::saveSurfaceChart(QXmlStreamWriter &writer) const
{
    writer.writeStartElement(type == Chart::Type::Surface3D ? QLatin1String("c:surface3DChart") : QLatin1String("c:surfaceChart"));
    writeEmptyElement(writer, QLatin1String("c:wireframe"), wireframe);
    for (const auto &ser: qAsConst(seriesList)) ser.write(writer);
    saveBandFormats(writer);
    if (axesIds.isEmpty()) {
        qDebug()<<"[Error] Axes IDs are empty!";
        return;
    }
    for (const auto id: qAsConst(axesIds)) writeEmptyElement(writer, QLatin1String("c:axId"), id);
    writer.writeEndElement();
}

void SubChart::saveBubbleChart(QXmlStreamWriter &writer) const
{
    writer.writeStartElement(QLatin1String("c:bubbleChart"));
    writeEmptyElement(writer, QLatin1String("c:varyColors"), varyColors);
    for (const auto &ser: qAsConst(seriesList)) ser.write(writer);
    if (labels.isValid()) labels.write(writer);
    writeEmptyElement(writer, QLatin1String("c:bubble3D"), bubble3D);
    if (bubbleScale.has_value())
        writeEmptyElement(writer, QLatin1String("c:bubbleScale"), toST_PercentInt(bubbleScale.value()));
    writeEmptyElement(writer, QLatin1String("c:showNegBubbles"), showNegBubbles);
    if (bubbleSizeRepresents.has_value()) {
        if (bubbleSizeRepresents.value() == Chart::BubbleSizeRepresents::Area)
            writeEmptyElement(writer, QLatin1String("c:sizeRepresents"), QLatin1String("area"));
        if (bubbleSizeRepresents.value() == Chart::BubbleSizeRepresents::Width)
            writeEmptyElement(writer, QLatin1String("c:sizeRepresents"), QLatin1String("w"));
    }
    if (axesIds.isEmpty()) {
        qDebug()<<"[Error] Axes IDs are empty!";
        return;
    }
    for (const auto id: qAsConst(axesIds)) writeEmptyElement(writer, QLatin1String("c:axId"), id);

    writer.writeEndElement();
}

void SubChart::savePieChart(QXmlStreamWriter &writer) const
{
    switch (type) {
        case Chart::Type::Pie: writer.writeStartElement(QLatin1String("c:pieChart")); break;
        case Chart::Type::Pie3D: writer.writeStartElement(QLatin1String("c:pie3DChart")); break;
        case Chart::Type::OfPie: writer.writeStartElement(QLatin1String("c:ofPieChart")); break;
        case Chart::Type::Doughnut: writer.writeStartElement(QLatin1String("c:doughnutChart")); break;
        default: return;
    }

    if (type == Chart::Type::OfPie) {
        if (ofPieType == Chart::OfPieType::Bar)
            writeEmptyElement(writer, QLatin1String("c:ofPieType"), QLatin1String("bar"));
        if (ofPieType == Chart::OfPieType::Pie)
            writeEmptyElement(writer, QLatin1String("c:ofPieType"), QLatin1String("pie"));
    }
    writeEmptyElement(writer, QLatin1String("c:varyColors"), varyColors);
    for (const auto &ser: qAsConst(seriesList)) ser.write(writer);
    if (labels.isValid()) labels.write(writer);
    if (type == Chart::Type::Pie || type == Chart::Type::Doughnut) {
        if (firstSliceAng.has_value())
            writeEmptyElement(writer, QLatin1String("c:firstSliceAng"), firstSliceAng);
    }
    if (type == Chart::Type::Doughnut && holeSize.has_value())
        writeEmptyElement(writer, QLatin1String("c:holeSize"), toST_PercentInt(holeSize.value()));
    if (type == Chart::Type::OfPie) {
        if (gapWidth.has_value())
            writeEmptyElement(writer, QLatin1String("c:gapWidth"), toST_PercentInt(gapWidth.value()));
        if (splitType.has_value())
            writeEmptyElement(writer, QLatin1String("c:splitType"), Chart::toString(splitType.value()));
        if (splitPos.has_value())
            writeEmptyElement(writer, QLatin1String("c:splitPos"), QString::number(splitPos.value()));
        if (!customSplit.isEmpty()) {
            writer.writeStartElement(QLatin1String("c:custSplit"));
            for (int i: qAsConst(customSplit)) writeEmptyElement(writer, QLatin1String("c:secondPiePt"), i);
            writer.writeEndElement();
        }
        if (secondPieSize.has_value())
            writeEmptyElement(writer, QLatin1String("c:secondPieSize"), toST_PercentInt(secondPieSize.value()));
        if (!serLines.isEmpty()) {
            for (const auto &ser: qAsConst(serLines)) {
                writer.writeStartElement(QLatin1String("c:serLines"));
                ser.write(writer, QLatin1String("c:spPr"));
                writer.writeEndElement();
            }
        }
    }

    writer.writeEndElement();
}

void SubChart::saveLineChart(QXmlStreamWriter &writer) const
{
    writer.writeStartElement(type == Chart::Type::Line3D ? QLatin1String("c:line3DChart") : QLatin1String("c:lineChart"));
    if (grouping.has_value())
        writeEmptyElement(writer, QLatin1String("c:grouping"), Chart::toString(grouping.value()));
    writeEmptyElement(writer, QLatin1String("c:varyColors"), varyColors);
    for (const auto &ser: qAsConst(seriesList)) ser.write(writer);
    if (labels.isValid()) labels.write(writer);
    if (dropLines.has_value()) {
        if (dropLines->isValid()) {
            writer.writeStartElement(QLatin1String("c:dropLines"));
            dropLines->write(writer, QLatin1String("c:spPr"));
            writer.writeEndElement();
        }
        else writer.writeEmptyElement(QLatin1String("c:dropLines"));
    }
    if (type == Chart::Type::Line) {
        if (hiLowLines.has_value()) {
            if (hiLowLines->isValid()) {
                writer.writeStartElement(QLatin1String("c:hiLowLines"));
                hiLowLines->write(writer, QLatin1String("c:spPr"));
                writer.writeEndElement();
            }
            else writer.writeEmptyElement(QLatin1String("c:hiLowLines"));
        }
        if (upDownBars.isValid()) upDownBars.write(writer, QLatin1String("c:upDownBars"));
        writeEmptyElement(writer, QLatin1String("c:marker"), marker);
        writeEmptyElement(writer, QLatin1String("c:smooth"), smooth);
    }
    if (gapDepth.has_value() && type == Chart::Type::Line3D)
        writeEmptyElement(writer, QLatin1String("c:gapDepth"), toST_PercentInt(gapDepth.value()));
    if (axesIds.isEmpty()) {
        qDebug()<<"[Error] Axes IDs are empty!";
        return;
    }
    for (const auto id: qAsConst(axesIds)) writeEmptyElement(writer, QLatin1String("c:axId"), id);

    writer.writeEndElement();
}

void SubChart::saveBarChart(QXmlStreamWriter &writer) const
{
    switch (type) {
        case Chart::Type::Bar: writer.writeStartElement(QLatin1String("c:barChart")); break;
        case Chart::Type::Bar3D: writer.writeStartElement(QLatin1String("c:bar3DChart")); break;
        default: return;
    }

    if (barDirection == Chart::BarDirection::Bar)
        writeEmptyElement(writer, QLatin1String("c:barDir"), QLatin1String("bar"));
    if (barDirection == Chart::BarDirection::Column)
        writeEmptyElement(writer, QLatin1String("c:barDir"), QLatin1String("col"));
    if (barGrouping.has_value())
        writeEmptyElement(writer, QLatin1String("c:grouping"), Chart::toString(barGrouping.value()));
    writeEmptyElement(writer, QLatin1String("c:varyColors"), varyColors);
    for (const auto &ser: qAsConst(seriesList)) ser.write(writer);
    if (labels.isValid()) labels.write(writer);

    if (gapWidth.has_value())
        writeEmptyElement(writer, QLatin1String("c:gapWidth"), toST_PercentInt(gapWidth.value()));
    if (gapDepth.has_value() && type == Chart::Type::Bar3D)
        writeEmptyElement(writer, QLatin1String("c:gapDepth"), toST_PercentInt(gapDepth.value()));

    if (type == Chart::Type::Bar) {
        if (overlap.has_value())
            writeEmptyElement(writer, QLatin1String("c:overlap"), toST_PercentInt(overlap.value()));
        if (!serLines.isEmpty()) {
            for (const auto &ser: qAsConst(serLines)) {
                writer.writeStartElement(QLatin1String("c:serLines"));
                ser.write(writer, QLatin1String("c:spPr"));
                writer.writeEndElement();
            }
        }
    }
    if (type == Chart::Type::Bar3D && barShape.has_value())
        writeEmptyElement(writer, QLatin1String("c:shape"), Series::toString(barShape.value()));
    if (axesIds.isEmpty()) {
        qDebug()<<"[Error] Axes IDs are empty!";
        return;
    }
    for (const auto id: qAsConst(axesIds)) writeEmptyElement(writer, QLatin1String("c:axId"), id);

    writer.writeEndElement();
}

void SubChart::saveScatterChart(QXmlStreamWriter &writer) const
{
    writer.writeStartElement(QLatin1String("c:scatterChart"));

    writeEmptyElement(writer, QLatin1String("c:scatterStyle"), Chart::toString(scatterStyle));
    writeEmptyElement(writer, QLatin1String("c:varyColors"), varyColors);
    for (const auto &ser: qAsConst(seriesList)) ser.write(writer);
    if (labels.isValid()) labels.write(writer);
    if (axesIds.isEmpty()) {
        qDebug()<<"[Error] Axes IDs are empty!";
        return;
    }
    for (const auto id: qAsConst(axesIds)) writeEmptyElement(writer, QLatin1String("c:axId"), id);

    writer.writeEndElement();
}

void SubChart::saveStockChart(QXmlStreamWriter &writer) const
{
    writer.writeStartElement(QLatin1String("c:stockChart"));

    for (const auto &ser: qAsConst(seriesList)) ser.write(writer);
    if (labels.isValid()) labels.write(writer);
    if (dropLines.has_value()) {
        if (dropLines->isValid()) {
            writer.writeStartElement(QLatin1String("c:dropLines"));
            dropLines->write(writer, QLatin1String("c:spPr"));
            writer.writeEndElement();
        }
        else writer.writeEmptyElement(QLatin1String("c:dropLines"));
    }
    if (hiLowLines.has_value()) {
        if (hiLowLines->isValid()) {
            writer.writeStartElement(QLatin1String("c:hiLowLines"));
            hiLowLines->write(writer, QLatin1String("c:spPr"));
            writer.writeEndElement();
        }
        else writer.writeEmptyElement(QLatin1String("c:hiLowLines"));
    }
    if (upDownBars.isValid()) upDownBars.write(writer, QLatin1String("c:upDownBars"));
    if (axesIds.isEmpty()) {
        qDebug()<<"[Error] Axes IDs are empty!";
        return;
    }
    for (const auto id: qAsConst(axesIds)) writeEmptyElement(writer, QLatin1String("c:axId"), id);

    writer.writeEndElement();
}

void SubChart::saveRadarChart(QXmlStreamWriter &writer) const
{
    writer.writeStartElement(QLatin1String("c:radarChart"));

    writeEmptyElement(writer, QLatin1String("c:radarStyle"), Chart::toString(radarStyle));
    writeEmptyElement(writer, QLatin1String("c:varyColors"), varyColors);
    for (const auto &ser: qAsConst(seriesList)) ser.write(writer);
    if (labels.isValid()) labels.write(writer);
    if (axesIds.isEmpty()) {
        qDebug()<<"[Error] Axes IDs are empty!";
        return;
    }
    for (const auto id: qAsConst(axesIds)) writeEmptyElement(writer, QLatin1String("c:axId"), id);

    writer.writeEndElement();
}

void SubChart::saveBandFormats(QXmlStreamWriter &writer) const
{
    if (bandFormats.isEmpty()) return;
    writer.writeStartElement(QLatin1String("c:bandFmts"));
    for (int key: bandFormats.keys()) {
        writer.writeStartElement(QLatin1String("c:bandFmt"));
        writeEmptyElement(writer, QLatin1String("c:idx"), key);
        if (const auto &val = bandFormats.value(key); val.isValid())
            val.write(writer, QLatin1String("c:spPr"));
        writer.writeEndElement();
    }
    writer.writeEndElement();
}

void UpDownBar::read(QXmlStreamReader &reader)
{
    const auto& name = reader.name();
    while (!reader.atEnd()) {
        auto token = reader.readNext();
        if (token == QXmlStreamReader::StartElement) {
            if (reader.name() == QLatin1String("gapWidth"))
                parseAttributePercent(reader.attributes(), QLatin1String("val"), gapWidth);
            else if (reader.name() == QLatin1String("upBars")) {
                reader.readNextStartElement();
                upBar.read(reader);
            }
            else if (reader.name() == QLatin1String("downBars")) {
                reader.readNextStartElement();
                downBar.read(reader);
            }
            else if (reader.name() == QLatin1String("extLst")) {
                reader.skipCurrentElement();
            }
        }
        else if (token == QXmlStreamReader::EndElement && reader.name() == name)
            break;
    }
}

void UpDownBar::write(QXmlStreamWriter &writer, const QString &name) const
{
    if (!isValid()) return;
    writer.writeStartElement(name);
    if (gapWidth.has_value()) {
        writer.writeEmptyElement(QLatin1String("c:gapWidth"));
        writer.writeAttribute(QLatin1String("val"), toST_PercentInt(gapWidth.value()));
    }
    if (upBar.isValid()) {
        writer.writeStartElement(QLatin1String("c:upBars"));
        upBar.write(writer, QLatin1String("c:spPr"));
        writer.writeEndElement();
    }
    if (downBar.isValid()) {
        writer.writeStartElement(QLatin1String("c:downBars"));
        downBar.write(writer, QLatin1String("c:spPr"));
        writer.writeEndElement();
    }
    writer.writeEndElement();
}

bool UpDownBar::isValid() const
{
    return gapWidth.has_value() || upBar.isValid() || downBar.isValid();
}

SubChart::SubChart(Chart::Type type) : type{type}
{
    if (type == Chart::Type::Pie || type == Chart::Type::Pie3D ||
        type == Chart::Type::Doughnut || type == Chart::Type::OfPie ||
        type == Chart::Type::Bubble)
    varyColors = true;
}

Series *SubChart::addSeries(int index)
{
    auto series = Series(index, index);
    switch (type) {
        case Chart::Type::None:
            break;
        case Chart::Type::Area:
        case Chart::Type::Area3D: series.setType(Series::Type::Area);
            break;
        case Chart::Type::Stock:
        case Chart::Type::Line:
        case Chart::Type::Line3D: series.setType(Series::Type::Line);
            break;
        case Chart::Type::Radar: series.setType(Series::Type::Radar);
            break;
        case Chart::Type::Scatter: series.setType(Series::Type::Scatter);
            break;
        case Chart::Type::Pie:
        case Chart::Type::Pie3D:
        case Chart::Type::OfPie: series.setType(Series::Type::Pie);
            break;
        case Chart::Type::Doughnut: series.setType(Series::Type::Pie);
            break;
        case Chart::Type::Bar:
        case Chart::Type::Bar3D: series.setType(Series::Type::Bar);
            break;
        case Chart::Type::Surface:
        case Chart::Type::Surface3D: series.setType(Series::Type::Surface);
            break;
        case Chart::Type::Bubble: series.setType(Series::Type::Bubble);
            break;
    }

    //set axes to the new series
    //series.setAxes(d->subcharts.last().axesIds);

    //append series to the last subchart
    seriesList << series;
    return &seriesList.last();
}

bool SubChart::read(QXmlStreamReader &reader)
{
    const auto& name = reader.name();
    Chart::fromString(name.toString(), type);

    if (type == Chart::Type::None) {
        qDebug() << "[undefined chart type] " << name;
        return false;
    }

    while (!reader.atEnd()) {
        auto token = reader.readNext();
        if (token == QXmlStreamReader::StartElement) {
            switch (type) {
                case Chart::Type::Area:
                case Chart::Type::Area3D:
                    loadAreaChart(reader); break;
                case Chart::Type::Surface:
                case Chart::Type::Surface3D:
                    loadSurfaceChart(reader); break;
                case Chart::Type::Pie:
                case Chart::Type::Pie3D:
                case Chart::Type::Doughnut:
                case Chart::Type::OfPie:
                    loadPieChart(reader); break;
                case Chart::Type::Line:
                case Chart::Type::Line3D:
                    loadLineChart(reader); break;
                case Chart::Type::Bubble:
                    loadBubbleChart(reader); break;
                case Chart::Type::Bar:
                case Chart::Type::Bar3D:
                    loadBarChart(reader); break;
                case Chart::Type::Scatter:
                    loadScatterChart(reader); break;
                case Chart::Type::Stock:
                    loadStockChart(reader); break;
                case Chart::Type::Radar:
                    loadRadarChart(reader);
                default: break;
            }
        }
        else if (token == QXmlStreamReader::EndElement && reader.name() == name)
            break;
    }
    return true;
}

void SubChart::write(QXmlStreamWriter &writer) const
{
    if (type == Chart::Type::None) {
        qDebug() << "[undefined chart type] ";
        return;
    }

    switch (type) {
        case Chart::Type::Area:
        case Chart::Type::Area3D:
            saveAreaChart(writer); break;
        case Chart::Type::Surface:
        case Chart::Type::Surface3D:
            saveSurfaceChart(writer); break;
        case Chart::Type::Pie:
        case Chart::Type::Pie3D:
        case Chart::Type::Doughnut:
        case Chart::Type::OfPie:
            savePieChart(writer); break;
        case Chart::Type::Line:
        case Chart::Type::Line3D:
            saveLineChart(writer); break;
        case Chart::Type::Bubble:
            saveBubbleChart(writer); break;
        case Chart::Type::Bar:
        case Chart::Type::Bar3D:
            saveBarChart(writer); break;
        case Chart::Type::Scatter:
            saveScatterChart(writer); break;
        case Chart::Type::Stock:
            saveStockChart(writer); break;
        case Chart::Type::Radar:
            saveRadarChart(writer);
        default: break;
    }
}

std::optional<Chart::Grouping> Chart::grouping(int subchartIndex) const
{
    Q_D(const Chart);

    if (subchartIndex < 0 || subchartIndex >= d->subcharts.size()) return {};
    return d->subcharts[subchartIndex].grouping;
}

void Chart::setGrouping(int subchartIndex, Chart::Grouping grouping)
{
    Q_D(Chart);

    if (subchartIndex < 0 || subchartIndex >= d->subcharts.size()) return;
    d->subcharts[subchartIndex].grouping = grouping;
}

std::optional<Chart::BarGrouping> Chart::barGrouping(int subchartIndex) const
{
    Q_D(const Chart);

    if (subchartIndex < 0 || subchartIndex >= d->subcharts.size()) return {};
    return d->subcharts[subchartIndex].barGrouping;
}

void Chart::setBarGrouping(int subchartIndex, Chart::BarGrouping grouping)
{
    Q_D(Chart);

    if (subchartIndex < 0 || subchartIndex >= d->subcharts.size()) return;

    d->subcharts[subchartIndex].barGrouping = grouping;
}

std::optional<bool> Chart::varyColors(int subchartIndex) const
{
    Q_D(const Chart);

    if (subchartIndex < 0 || subchartIndex >= d->subcharts.size()) return {};
    return d->subcharts[subchartIndex].varyColors;
}

void Chart::setVaryColors(int subchartIndex, bool varyColors)
{
    Q_D(Chart);

    if (subchartIndex < 0 || subchartIndex >= d->subcharts.size()) return;
    d->subcharts[subchartIndex].varyColors = varyColors;
}

Labels Chart::labels(int subchartIndex) const
{
    Q_D(const Chart);

    if (subchartIndex < 0 || subchartIndex >= d->subcharts.size()) return {};
    return d->subcharts[subchartIndex].labels;
}

Labels &Chart::labels(int subchartIndex)
{
    Q_D(Chart);

    return d->subcharts[subchartIndex].labels;
}

void Chart::setLabels(int subchartIndex, const Labels &labels)
{
    Q_D(Chart);

    if (subchartIndex < 0 || subchartIndex >= d->subcharts.size()) return;
    d->subcharts[subchartIndex].labels = labels;
}

ShapeFormat Chart::dropLines(int subchartIndex) const
{
    Q_D(const Chart);

    if (subchartIndex < 0 || subchartIndex >= d->subcharts.size()) return {};
    return d->subcharts[subchartIndex].dropLines.value_or(ShapeFormat());
}

ShapeFormat &Chart::dropLines(int subchartIndex)
{
    Q_D(Chart);

    if (!d->subcharts[subchartIndex].dropLines.has_value())
        d->subcharts[subchartIndex].dropLines = ShapeFormat();
    return d->subcharts[subchartIndex].dropLines.value();
}

void Chart::setDropLines(int subchartIndex, const ShapeFormat &dropLines)
{
    Q_D(Chart);

    if (subchartIndex < 0 || subchartIndex >= d->subcharts.size()) return;
    d->subcharts[subchartIndex].dropLines = dropLines;
}

void  Chart::removeDropLines(int subchartIndex)
{
    Q_D(Chart);

    if (subchartIndex < 0 || subchartIndex >= d->subcharts.size()) return;
    d->subcharts[subchartIndex].dropLines.reset();
}

ShapeFormat Chart::hiLowLines(int subchartIndex) const
{
    Q_D(const Chart);

    if (subchartIndex < 0 || subchartIndex >= d->subcharts.size()) return {};
    return d->subcharts[subchartIndex].hiLowLines.value_or(ShapeFormat());
}

ShapeFormat &Chart::hiLowLines(int subchartIndex)
{
    Q_D(Chart);

    if (!d->subcharts[subchartIndex].hiLowLines.has_value())
        d->subcharts[subchartIndex].hiLowLines = ShapeFormat();
    return d->subcharts[subchartIndex].hiLowLines.value();
}

void Chart::setHiLowLines(int subchartIndex, const ShapeFormat &hiLowLines)
{
    Q_D(Chart);

    if (subchartIndex < 0 || subchartIndex >= d->subcharts.size()) return;
    d->subcharts[subchartIndex].hiLowLines = hiLowLines;
}

void Chart::removeHiLowLines(int subchartIndex)
{
    Q_D(Chart);
    if (subchartIndex < 0 || subchartIndex >= d->subcharts.size()) return;
    d->subcharts[subchartIndex].hiLowLines.reset();
}

UpDownBar Chart::upDownBars(int subchartIndex) const
{
    Q_D(const Chart);

    if (subchartIndex < 0 || subchartIndex >= d->subcharts.size()) return {};
    return d->subcharts[subchartIndex].upDownBars;
}

UpDownBar &Chart::upDownBars(int subchartIndex)
{
    Q_D(Chart);

    return d->subcharts[subchartIndex].upDownBars;
}

void Chart::setUpDownBars(int subchartIndex, const UpDownBar &upDownBars)
{
    Q_D(Chart);

    if (subchartIndex < 0 || subchartIndex >= d->subcharts.size()) return;
    d->subcharts[subchartIndex].upDownBars = upDownBars;
}

std::optional<bool> Chart::markerShown(int subchartIndex) const
{
    Q_D(const Chart);
    if (subchartIndex < 0 || subchartIndex >= d->subcharts.size()) return {};
    return d->subcharts[subchartIndex].marker;
}

void Chart::setMarkerShown(int subchartIndex, bool markerShown)
{
    Q_D(Chart);

    if (subchartIndex < 0 || subchartIndex >= d->subcharts.size()) return;
    d->subcharts[subchartIndex].marker = markerShown;
}

std::optional<int> Chart::gapDepth(int subchartIndex) const
{
    Q_D(const Chart);
    if (subchartIndex < 0 || subchartIndex >= d->subcharts.size()) return {};
    return d->subcharts[subchartIndex].gapDepth;
}

void Chart::setGapDepth(int subchartIndex, int gapDepth)
{
    Q_D(Chart);

    if (subchartIndex < 0 || subchartIndex >= d->subcharts.size()) return;
    d->subcharts[subchartIndex].gapDepth = gapDepth;
}

Chart::ScatterStyle Chart::scatterStyle(int subchartIndex) const
{
    Q_D(const Chart);
    if (subchartIndex < 0 || subchartIndex >= d->subcharts.size()) return {};
    return d->subcharts[subchartIndex].scatterStyle;
}

void Chart::setScatterStyle(int subchartIndex, Chart::ScatterStyle scatterStyle)
{
    Q_D(Chart);

    if (subchartIndex < 0 || subchartIndex >= d->subcharts.size()) return;
    d->subcharts[subchartIndex].scatterStyle = scatterStyle;
}

Chart::RadarStyle Chart::radarStyle(int subchartIndex) const
{
    Q_D(const Chart);
    if (subchartIndex < 0 || subchartIndex >= d->subcharts.size()) return {};
    return d->subcharts[subchartIndex].radarStyle;
}

void Chart::setRadarStyle(int subchartIndex, Chart::RadarStyle radarStyle)
{
    Q_D(Chart);

    if (subchartIndex < 0 || subchartIndex >= d->subcharts.size()) return;
    d->subcharts[subchartIndex].radarStyle = radarStyle;
}

Chart::BarDirection Chart::barDirection(int subchartIndex) const
{
    Q_D(const Chart);
    if (subchartIndex < 0 || subchartIndex >= d->subcharts.size()) return {};
    return d->subcharts[subchartIndex].barDirection;
}

void Chart::setBarDirection(int subchartIndex, Chart::BarDirection barDirection)
{
    Q_D(Chart);

    if (subchartIndex < 0 || subchartIndex >= d->subcharts.size()) return;
    d->subcharts[subchartIndex].barDirection = barDirection;
}

std::optional<int> Chart::gapWidth(int subchartIndex) const
{
    Q_D(const Chart);
    if (subchartIndex < 0 || subchartIndex >= d->subcharts.size()) return {};
    return d->subcharts[subchartIndex].gapWidth;
}

void Chart::setGapWidth(int subchartIndex, int gapWidth)
{
    Q_D(Chart);

    if (subchartIndex < 0 || subchartIndex >= d->subcharts.size()) return;
    d->subcharts[subchartIndex].gapWidth = gapWidth;
}

std::optional<int> Chart::overlap(int subchartIndex) const
{
    Q_D(const Chart);
    if (subchartIndex < 0 || subchartIndex >= d->subcharts.size()) return {};
    return d->subcharts[subchartIndex].overlap;
}

void Chart::setOverlap(int subchartIndex, int overlap)
{
    Q_D(Chart);

    if (subchartIndex < 0 || subchartIndex >= d->subcharts.size()) return;
    d->subcharts[subchartIndex].overlap = overlap;
}

QList<ShapeFormat> Chart::seriesLines(int subchartIndex) const
{
    Q_D(const Chart);
    if (subchartIndex < 0 || subchartIndex >= d->subcharts.size()) return {};
    return d->subcharts[subchartIndex].serLines;
}

QList<ShapeFormat> &Chart::seriesLines(int subchartIndex)
{
    Q_D(Chart);

    return d->subcharts[subchartIndex].serLines;
}

void Chart::setSeriesLines(int subchartIndex, const QList<ShapeFormat> &seriesLines)
{
    Q_D(Chart);

    if (subchartIndex < 0 || subchartIndex >= d->subcharts.size()) return;
    d->subcharts[subchartIndex].serLines = seriesLines;
}

std::optional<Series::BarShape> Chart::barShape(int subchartIndex) const
{
    Q_D(const Chart);
    if (subchartIndex < 0 || subchartIndex >= d->subcharts.size()) return {};
    return d->subcharts[subchartIndex].barShape;
}

void Chart::setBarShape(int subchartIndex, Series::BarShape barShape)
{
    Q_D(Chart);

    if (subchartIndex < 0 || subchartIndex >= d->subcharts.size()) return;
    d->subcharts[subchartIndex].barShape = barShape;
}

std::optional<int> Chart::firstSliceAngle(int subchartIndex) const
{
    Q_D(const Chart);
    if (subchartIndex < 0 || subchartIndex >= d->subcharts.size()) return {};
    return d->subcharts[subchartIndex].firstSliceAng;
}

void Chart::setFirstSliceAngle(int subchartIndex, int angle)
{
    Q_D(Chart);

    if (subchartIndex < 0 || subchartIndex >= d->subcharts.size()) return;
    d->subcharts[subchartIndex].firstSliceAng = angle;
}

std::optional<int> Chart::holeSize(int subchartIndex) const
{
    Q_D(const Chart);
    if (subchartIndex < 0 || subchartIndex >= d->subcharts.size()) return {};
    return d->subcharts[subchartIndex].holeSize;
}

void Chart::setHoleSise(int subchartIndex, int holeSize)
{
    Q_D(Chart);

    if (subchartIndex < 0 || subchartIndex >= d->subcharts.size()) return;
    d->subcharts[subchartIndex].holeSize = holeSize;
}

Chart::OfPieType Chart::ofPieType(int subchartIndex) const
{
    Q_D(const Chart);
    if (subchartIndex < 0 || subchartIndex >= d->subcharts.size()) return {};
    return d->subcharts[subchartIndex].ofPieType;
}

void Chart::setOfPieType(int subchartIndex, OfPieType type)
{
    Q_D(Chart);

    if (subchartIndex < 0 || subchartIndex >= d->subcharts.size()) return;
    d->subcharts[subchartIndex].ofPieType = type;
}

std::optional<Chart::SplitType> Chart::splitType(int subchartIndex) const
{
    Q_D(const Chart);
    if (subchartIndex < 0 || subchartIndex >= d->subcharts.size()) return {};
    return d->subcharts[subchartIndex].splitType;
}

void Chart::setSplitType(int subchartIndex, SplitType splitType)
{
    Q_D(Chart);

    if (subchartIndex < 0 || subchartIndex >= d->subcharts.size()) return;
    d->subcharts[subchartIndex].splitType = splitType;
}

std::optional<double> Chart::splitPos(int subchartIndex) const
{
    Q_D(const Chart);
    if (subchartIndex < 0 || subchartIndex >= d->subcharts.size()) return {};
    return d->subcharts[subchartIndex].splitPos;
}

void Chart::setSplitPos(int subchartIndex, double value)
{
    Q_D(Chart);

    if (subchartIndex < 0 || subchartIndex >= d->subcharts.size()) return;
    d->subcharts[subchartIndex].splitPos = value;
}

QList<int> Chart::customSplit(int subchartIndex) const
{
    Q_D(const Chart);
    if (subchartIndex < 0 || subchartIndex >= d->subcharts.size()) return {};
    return d->subcharts[subchartIndex].customSplit;
}

void Chart::setCustomSplit(int subchartIndex, const QList<int> &indexes)
{
    Q_D(Chart);

    if (subchartIndex < 0 || subchartIndex >= d->subcharts.size()) return;
    d->subcharts[subchartIndex].customSplit = indexes;
}

std::optional<int> Chart::secondPieSize(int subchartIndex) const
{
    Q_D(const Chart);
    if (subchartIndex < 0 || subchartIndex >= d->subcharts.size()) return {};
    return d->subcharts[subchartIndex].secondPieSize;
}

void Chart::setSecondPieSize(int subchartIndex, int size)
{
    Q_D(Chart);

    if (subchartIndex < 0 || subchartIndex >= d->subcharts.size()) return;
    d->subcharts[subchartIndex].secondPieSize = size;
}

std::optional<bool> Chart::bubble3D(int subchartIndex) const
{
    Q_D(const Chart);
    if (subchartIndex < 0 || subchartIndex >= d->subcharts.size()) return {};
    return d->subcharts[subchartIndex].bubble3D;
}

void Chart::setBubble3D(int subchartIndex, bool bubble3D)
{
    Q_D(Chart);

    if (subchartIndex < 0 || subchartIndex >= d->subcharts.size()) return;
    d->subcharts[subchartIndex].bubble3D = bubble3D;
}

std::optional<bool> Chart::showNegativeBubbles(int subchartIndex) const
{
    Q_D(const Chart);
    if (subchartIndex < 0 || subchartIndex >= d->subcharts.size()) return {};
    return d->subcharts[subchartIndex].showNegBubbles;
}

void Chart::setShowNegativeBubbles(int subchartIndex, bool show)
{
    Q_D(Chart);

    if (subchartIndex < 0 || subchartIndex >= d->subcharts.size()) return;
    d->subcharts[subchartIndex].showNegBubbles = show;
}

std::optional<int> Chart::bubbleScale(int subchartIndex) const
{
    Q_D(const Chart);
    if (subchartIndex < 0 || subchartIndex >= d->subcharts.size()) return {};
    return d->subcharts[subchartIndex].bubbleScale;
}

void Chart::setBubbleScale(int subchartIndex, int scale)
{
    Q_D(Chart);

    if (subchartIndex < 0 || subchartIndex >= d->subcharts.size()) return;
    d->subcharts[subchartIndex].bubbleScale = scale;
}

std::optional<Chart::BubbleSizeRepresents> Chart::bubbleSizeRepresents(int subchartIndex) const
{
    Q_D(const Chart);
    if (subchartIndex < 0 || subchartIndex >= d->subcharts.size()) return {};
    return d->subcharts[subchartIndex].bubbleSizeRepresents;
}

void Chart::setBubbleSizeRepresents(int subchartIndex, BubbleSizeRepresents value)
{
    Q_D(Chart);

    if (subchartIndex < 0 || subchartIndex >= d->subcharts.size()) return;
    d->subcharts[subchartIndex].bubbleSizeRepresents = value;
}

std::optional<bool> Chart::wireframe(int subchartIndex) const
{
    Q_D(const Chart);
    if (subchartIndex < 0 || subchartIndex >= d->subcharts.size()) return {};
    return d->subcharts[subchartIndex].wireframe;
}

void Chart::setWireframe(int subchartIndex, bool wireframe)
{
    Q_D(Chart);

    if (subchartIndex < 0 || subchartIndex >= d->subcharts.size()) return;
    d->subcharts[subchartIndex].wireframe = wireframe;
}

QMap<int, ShapeFormat> Chart::bandFormats(int subchartIndex) const
{
    Q_D(const Chart);
    if (subchartIndex < 0 || subchartIndex >= d->subcharts.size()) return {};
    return d->subcharts[subchartIndex].bandFormats;
}

void Chart::setBandFormats(int subchartIndex, QMap<int, ShapeFormat> bandFormats)
{
    Q_D(Chart);

    if (subchartIndex < 0 || subchartIndex >= d->subcharts.size()) return;
    d->subcharts[subchartIndex].bandFormats = bandFormats;
}




DataTable &Chart::dataTable()
{
    Q_D(Chart);
    if (!d->dTable.has_value()) d->dTable = DataTable();
    return d->dTable.value();
}

void Chart::setDataTableVisible(bool visible)
{
    Q_D(Chart);
    if (visible && !d->dTable.has_value()) d->dTable = DataTable();
    if (!visible) d->dTable.reset();
}

DataTable Chart::dataTable() const
{
    Q_D(const Chart);
    return d->dTable.value_or(DataTable());
}

void Chart::setDataTable(const DataTable &dataTable)
{
    Q_D(Chart);
    d->dTable = dataTable;
}

std::optional<int> Chart::styleID() const
{
    Q_D(const Chart);
    return d->styleID;
}

void Chart::setStyleID(int id)
{
    Q_D(Chart);
    d->styleID = id;
}

std::optional<bool> Chart::date1904() const
{
    Q_D(const Chart);
    return d->date1904;
}

void Chart::setDate1904(bool isDate1904)
{
    Q_D(Chart);
    d->date1904 = isDate1904;
}

QString Chart::lastUsedLanguage() const
{
    Q_D(const Chart);
    return d->language;
}

void Chart::setLastUsedLanguage(QString lang)
{
    Q_D(Chart);
    d->language = lang;
}

std::optional<bool> Chart::roundedCorners() const
{
    Q_D(const Chart);
    return d->roundedCorners;
}

void Chart::setRoundedCorners(bool rounded)
{
    Q_D(Chart);
    d->roundedCorners = rounded;
}

void DataTable::read(QXmlStreamReader &reader)
{
    const auto &name = reader.name();
    while (!reader.atEnd()) {
        auto token = reader.readNext();
        if (token == QXmlStreamReader::StartElement) {
            const auto &a = reader.attributes();
            if (reader.name() == QLatin1String("showHorzBorder"))
                parseAttributeBool(a, QLatin1String("val"), showHorizontalBorder);
            else if (reader.name() == QLatin1String("showVertBorder"))
                parseAttributeBool(a, QLatin1String("val"), showVerticalBorder);
            else if (reader.name() == QLatin1String("showOutline"))
                parseAttributeBool(a, QLatin1String("val"), showOutline);
            else if (reader.name() == QLatin1String("showKeys"))
                parseAttributeBool(a, QLatin1String("val"), showKeys);
            else if (reader.name() == QLatin1String("spPr"))
                shape.read(reader);
            else if (reader.name() == QLatin1String("txPr"))
                textProperties.read(reader);
            else if (reader.name() == QLatin1String("extLst"))
                extension.read(reader);
        }
        else if (token == QXmlStreamReader::EndElement && reader.name() == name)
            break;
    }
}

void DataTable::write(QXmlStreamWriter &writer, const QString &name) const
{
    writer.writeStartElement(name);
    writeEmptyElement(writer, QLatin1String("c:showHorzBorder"), showHorizontalBorder);
    writeEmptyElement(writer, QLatin1String("c:showVertBorder"), showVerticalBorder);
    writeEmptyElement(writer, QLatin1String("c:showOutline"), showOutline);
    writeEmptyElement(writer, QLatin1String("c:showKeys"), showKeys);
    if (shape.isValid()) shape.write(writer, QLatin1String("c:spPr"));
    if (textProperties.isValid()) textProperties.write(writer, QLatin1String("c:txPr"));
    if (extension.isValid()) extension.write(writer, QLatin1String("c:extLst"));
    writer.writeEndElement();
}

bool DataTable::isValid() const
{
    if (showHorizontalBorder.has_value()) return true;
    if (showVerticalBorder.has_value()) return true;
    if (showOutline.has_value()) return true;
    if (showKeys.has_value()) return true;
    if (shape.isValid()) return true;
    if (textProperties.isValid()) return true;
    if (extension.isValid()) return true;
    return false;
}

void Chart::loadMediaFiles(Workbook *workbook)
{
    Q_D(Chart);

    if (d->relationships->isEmpty()) return;

    QList<std::reference_wrapper<FillFormat>> fills;

    if (d->chartSpaceShape.isValid())
        fills << d->chartSpaceShape.fills();
    if (d->plotAreaShape.isValid())
        fills << d->plotAreaShape.fills();
    for (auto &sub: d->subcharts) {
        if (sub.dropLines.has_value() && sub.dropLines->isValid()) fills << sub.dropLines->fills();
        if (sub.hiLowLines.has_value() && sub.hiLowLines->isValid()) fills << sub.hiLowLines->fills();
        if (sub.upDownBars.isValid()) {
            if (sub.upDownBars.upBar.isValid()) fills << sub.upDownBars.upBar.fills();
            if (sub.upDownBars.downBar.isValid()) fills << sub.upDownBars.downBar.fills();
        }
        for (auto &sl: sub.serLines)
            if (sl.isValid()) fills << sl.fills();
        for (auto &s: sub.bandFormats.values()) {
            if (s.isValid()) fills << s.fills();
        }
        if (sub.labels.isValid()) fills << sub.labels.fills();
    }
    if (d->legend.isValid() && d->legend.shape().isValid())
        fills << d->legend.shape().fills();
    for (auto &ax: d->axisList) {
        if (!ax.isValid()) continue;
        if (ax.majorGridLines().isValid()) fills << ax.majorGridLines().fills();
        if (ax.minorGridLines().isValid()) fills << ax.minorGridLines().fills();
        if (ax.shape().isValid()) fills << ax.shape().fills();
    }

    for (auto &fill: fills) {
        if (fill.get().type() == FillFormat::FillType::PictureFill) {
            fill.get().loadBlip(workbook, d->relationships);
        }
    }
}

void Chart::saveMediaFiles(Workbook *workbook)
{
    Q_D(Chart);

    //In all the FillFormat objects we need to check whether they are blip fill
    //and register the fill picture in the document mediaFiles
    d->relationships->clear();

    QList<std::reference_wrapper<FillFormat>> fills;

    if (d->chartSpaceShape.isValid())
        fills << d->chartSpaceShape.fills();
    if (d->plotAreaShape.isValid())
        fills << d->plotAreaShape.fills();
    for (auto &sub: d->subcharts) {
        if (sub.dropLines.has_value() && sub.dropLines->isValid()) fills << sub.dropLines->fills();
        if (sub.hiLowLines.has_value() && sub.hiLowLines->isValid()) fills << sub.hiLowLines->fills();
        if (sub.upDownBars.isValid()) {
            if (sub.upDownBars.upBar.isValid()) fills << sub.upDownBars.upBar.fills();
            if (sub.upDownBars.downBar.isValid()) fills << sub.upDownBars.downBar.fills();
        }
        for (auto &sl: sub.serLines)
            if (sl.isValid()) fills << sl.fills();
        for (auto &s: sub.bandFormats.values()) {
            if (s.isValid()) fills << s.fills();
        }
        if (sub.labels.isValid()) fills << sub.labels.fills();
    }
    if (d->legend.isValid() && d->legend.shape().isValid())
        fills << d->legend.shape().fills();
    for (auto &ax: d->axisList) {
        if (!ax.isValid()) continue;
        if (ax.majorGridLines().isValid()) fills << ax.majorGridLines().fills();
        if (ax.minorGridLines().isValid()) fills << ax.minorGridLines().fills();
        if (ax.shape().isValid()) fills << ax.shape().fills();
    }

    for (auto &fill: fills) {
        if (fill.get().type() == FillFormat::FillType::PictureFill) {
            int id = fill.get().registerBlip(workbook); //after that blip has unique id
            if (id != -1) {
                d->relationships->addDocumentRelationship(QStringLiteral("/image"),
                                                          QStringLiteral("../media/image%1.%2")
                                                          .arg(id+1)
                                                          .arg("png")); //TODO: check
                fill.get().setPictureID(d->relationships->count());
            }
        }
    }
}

}




