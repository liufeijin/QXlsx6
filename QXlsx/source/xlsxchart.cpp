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
#include "xlsxworksheet.h"
#include "xlsxcellrange.h"
#include "xlsxutility_p.h"
#include "xlsxabstractooxmlfile_p.h"
#include "xlsxseries.h"
#include "xlsxworkbook.h"
#include "xlsxmediafile_p.h"
#include "QBuffer"

namespace QXlsx {

    class CT_XXXChart
    {
    public:
        CT_XXXChart(Chart::Type type);

        Chart::Type type = Chart::Type::None;

        /* Below are CT_XXXChart properties */

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
        ShapeFormat dropLines;

        //+line, +stock,
        ShapeFormat hiLowLines;

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
        Chart::Type chartType;

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
        //    std::optional<PrintSettings> printSettings;
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
        QList<CT_XXXChart> subcharts;
        QList<Axis> axisList;
        std::optional<DataTable> dTable;
        ShapeFormat plotAreaShape;
        ExtensionList plotAreaExtList;

        AbstractSheet* sheet;
    private:
        friend class Axis;
    };

ChartPrivate::ChartPrivate(Chart *q, Chart::CreateFlag flag)
    : AbstractOOXmlFilePrivate(q, flag), chartType(Chart::Type::None)
{

}

ChartPrivate::~ChartPrivate()
{
}

ChartPrivate &ChartPrivate::operator=(const ChartPrivate &other)
{
    chartType = other.chartType;
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
    //printSettings = other.printSettings;
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

/*!
 * \internal
 */
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

Chart::Type Chart::type() const
{
    Q_D(const Chart);

    return d->chartType;
}

void Chart::addSeries(const CellRange &range, AbstractSheet *sheet,
                      bool firstRowContainsHeaders,
                      bool firstColumnContainsCategoryData,
                      bool columnBased)
{
    Q_D(Chart);

    if (!range.isValid())
        return;
    if (sheet && sheet->type() != AbstractSheet::Type::Worksheet)
        return;
    if (!sheet && d->sheet->type() != AbstractSheet::Type::Worksheet)
        return;

    QString sheetName = sheet ? sheet->name() : d->sheet->name();
    //In case sheetName contains space or '
    sheetName = escapeSheetName(sheetName);

    if (range.columnCount() == 1 || range.rowCount() == 1) {
        auto series = addSeries();
        series->setValueData(sheetName + QLatin1String("!") + range.toString(true, true));
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
            auto series = addSeries();
            if (firstColumnContainsCategoryData)
                series->setCategoryData(categoryReference);
            series->setValueData(sheetName + QLatin1String("!") + subRange.toString(true, true));

            if (firstRowContainsHeaders) {
                CellRange subRange(range.firstRow(), col, range.firstRow(), col);
                series->setNameReference(sheetName + QLatin1String("!") + subRange.toString(true, true));
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
            auto series = addSeries();
            if (firstColumnContainsCategoryData)
                series->setCategoryData(categoryReference);
            series->setValueData(sheetName + QLatin1String("!") + subRange.toString(true, true));

            if (firstRowContainsHeaders) {
                CellRange subRange(row, range.firstColumn(), row, range.firstColumn());
                series->setNameReference(sheetName + QLatin1String("!") + subRange.toString(true, true));
            }
        }
    }
}

QXlsx::Series *Chart::addSeries(const CellRange &keyRange, const CellRange &valRange,
                                AbstractSheet *sheet, bool keyRangeIncludesHeader)
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

    auto series = addSeries();

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
    return series;
}

Series *Chart::addSeries()
{
    Q_D(Chart);

    //create default axes
    addDefaultAxes();

    return d->subcharts.last().addSeries(seriesCount());
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

    d->subcharts.clear();

    addDefaultAxes();
}

void Chart::setType(Type type)
{
    Q_D(Chart);

    d->chartType = type;
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

Axis& Chart::addAxis(Axis::Type type, Axis::Position pos, QString title)
{
    Q_D(Chart);

    d->addAxis(type, pos, title);
    return d->axisList.last();
}

void Chart::addAxis(const Axis &axis)
{
    Q_D(Chart);

    d->addAxis(axis);
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

    auto ser = std::find_if(d->subcharts.begin(), d->subcharts.end(), [axisID](const CT_XXXChart &c){
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

QList<Series> Chart::seriesThatUseAxis(int axisID) const
{
    Q_D(const Chart);
    QList<Series> result;

    for (const auto &sub: qAsConst(d->subcharts)) {
        if (sub.axesIds.contains(axisID)) {
            result.append(sub.seriesList);
        }
    }

    return result;
}

QList<Series> Chart::seriesThatUseAxis(Axis *axis) const
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

void Chart::setSeriesAxesIDs(const QList<int> &axesIds)
{
    Q_D(Chart);

    if (d->subcharts.isEmpty()) return;

    if (d->subcharts.size() > 1)
        qDebug() << "[warning] The chart has more than one subchart. The axesIDs "
                    "will be set to the last subchart.";
    d->subcharts.last().axesIds = axesIds;
}

void Chart::setSeriesAxesIDs(Series *series, const QList<int> &axesIds)
{
    Q_D(Chart);

    if (d->subcharts.isEmpty()) return;

    //first we need to find a subchart with the exact axes
    CT_XXXChart *subchart = nullptr;
    for (auto &sub: d->subcharts) {
        if (sub.axesIds == axesIds) {
            subchart = &sub;
            break;
        }
    }

    //second we need to find the series in subcharts
    CT_XXXChart *oldSubchart = nullptr;
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
        CT_XXXChart sub(d->chartType);
        d->subcharts << sub;
        sub.axesIds = axesIds;
        subchart = &sub;
    }

    //move the series from the old subchart to the new one
    subchart->seriesList << *series;
    oldSubchart->seriesList.removeOne(*series);
}

void Chart::setSeriesDefaultAxes()
{
    Q_D(Chart);

    QList<int> axesIds;
    for (const auto &ax: qAsConst(d->axisList)) axesIds << ax.id();
    setSeriesAxesIDs(axesIds);
}

void Chart::moveSeries(int oldOrder, int newOrder)
{
    auto s1 = seriesByOrder(oldOrder);
    auto s2 = seriesByOrder(newOrder);
    if (!s1 || !s2) return;
    s1->setOrder(newOrder);
    s2->setOrder(oldOrder);
}

QList<int> Chart::addDefaultAxes()
{
    Q_D(Chart);

    QList<int> ids;

    switch (d->chartType) {
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
            auto ax1 = axis(Axis::Type::Category, Axis::Position::Bottom);
            if (!ax1) ax1 = &addAxis(Axis::Type::Category, Axis::Position::Bottom);
            ids << ax1->id();
            auto ax2 = axis(Axis::Type::Value, Axis::Position::Left);
            if (!ax2) ax2 = &addAxis(Axis::Type::Value, Axis::Position::Left);
            ids << ax2->id();
            ax1->setCrossAxis(ax2);
            ax2->setCrossAxis(ax1);
            break;
        }
        case Type::Scatter: {
            auto ax1 = axis(Axis::Type::Value, Axis::Position::Bottom);
            if (!ax1) ax1 = &addAxis(Axis::Type::Value, Axis::Position::Bottom);
            ids << ax1->id();
            auto ax2 = axis(Axis::Type::Value, Axis::Position::Left);
            if (!ax2) ax2 = &addAxis(Axis::Type::Value, Axis::Position::Left);
            ids << ax2->id();
            ax1->setCrossAxis(ax2);
            ax2->setCrossAxis(ax1);
            break;
        }
        case Type::Area3D:
        case Type::Line3D:
        case Type::Bar3D:
        case Type::Surface:
        case Type::Surface3D: {
            auto ax1 = axis(Axis::Type::Category, Axis::Position::Bottom);
            if (!ax1) ax1 = &addAxis(Axis::Type::Category, Axis::Position::Bottom);
            ids << ax1->id();
            auto ax2 = axis(Axis::Type::Value, Axis::Position::Left);
            if (!ax2) ax2 = &addAxis(Axis::Type::Value, Axis::Position::Left);
            ids << ax2->id();
            auto ax3 = axis(Axis::Type::Series, Axis::Position::Bottom);
            if (!ax3) ax3 = &addAxis(Axis::Type::Series, Axis::Position::Bottom);
            ids << ax3->id();
            ax1->setCrossAxis(ax2);
            ax2->setCrossAxis(ax1);
            ax3->setCrossAxis(ax2);
            break;
        }
    }

    if (d->subcharts.isEmpty() || d->subcharts.last().type != d->chartType) {
        CT_XXXChart sub(d->chartType);
        d->subcharts << sub;
    }

    if (d->subcharts.last().axesIds.size() < ids.size())
        d->subcharts.last().axesIds = ids;

    return ids;
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

/*!
 * \internal
 */
void Chart::saveToXmlFile(QIODevice *device) const
{
    Q_D(const Chart);

    QXmlStreamWriter writer(device);

    writer.writeStartDocument(QStringLiteral("1.0"), true);

    // L.4.13.2.2 Chart
    //
    //  chartSpace is the root node, which contains an element defining the chart,
    // and an element defining the print settings for the chart.

    writer.writeStartElement(QStringLiteral("c:chartSpace"));

    writer.writeAttribute(QStringLiteral("xmlns:c"),
                          QStringLiteral("http://schemas.openxmlformats.org/drawingml/2006/chart"));
    writer.writeAttribute(QStringLiteral("xmlns:a"),
                          QStringLiteral("http://schemas.openxmlformats.org/drawingml/2006/main"));
    writer.writeAttribute(QStringLiteral("xmlns:r"),
                          QStringLiteral("http://schemas.openxmlformats.org/officeDocument/2006/relationships"));

    /*
    * chart is the root element for the chart. If the chart is a 3D chart,
    * then a view3D element is contained, which specifies the 3D view.
    * It then has a plot area, which defines a layout and contains an element
    * that corresponds to, and defines, the type of chart.
    */
    writeEmptyElement(writer, QLatin1String("c:date1904"), d->date1904);
    writeEmptyElement(writer, QLatin1String("c:lang"), d->language);
    writeEmptyElement(writer, QLatin1String("c:roundedCorners"), d->roundedCorners);
    writeEmptyElement(writer, QLatin1String("c:style"), d->styleID);

    d->saveXmlChart(writer);
    if (d->chartSpaceShape.isValid()) d->chartSpaceShape.write(writer, "c:spPr");
    if (d->textProperties.isValid()) d->textProperties.write(writer, QLatin1String("c:txPr"));
    if (d->chartSpaceExtList.isValid()) d->chartSpaceExtList.write(writer, QLatin1String("c:extLst"));

    writer.writeEndElement();// c:chartSpace
    writer.writeEndDocument();
}

/*!
 * \internal
 */
bool Chart::loadFromXmlFile(QIODevice *device)
{
    Q_D(Chart);

    QXmlStreamReader reader(device);
    while (!reader.atEnd()) {
        auto token = reader.readNext();
        if (token == QXmlStreamReader::StartElement) {
            const auto &a = reader.attributes();
            if (reader.name() == QLatin1String("chart")) {
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
            }
            else if (reader.name() == QLatin1String("pivotSource")) {
                //TODO
            }
            else if (reader.name() == QLatin1String("protection")) {
                //TODO
            }
            else if (reader.name() == QLatin1String("spPr"))
                d->chartSpaceShape.read(reader);
            else if (reader.name() == QLatin1String("txPr"))
                d->textProperties.read(reader);
            else if (reader.name() == QLatin1String("externalData")) {
                //TODO
            }
            else if (reader.name() == QLatin1String("printSettings")) {
                //TODO
            }
            else if (reader.name() == QLatin1String("userShapes")) {
                //TODO
            }
            else if (reader.name() == QLatin1String("extLst"))
                d->chartSpaceExtList.read(reader);
        }
    }

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
            }
            else if (reader.name() == QLatin1String("view3D")) {
                //TODO
            }
            else if (reader.name() == QLatin1String("floor")) {
                //TODO
            }
            else if (reader.name() == QLatin1String("sideWall")) {
                //TODO
            }
            else if (reader.name() == QLatin1String("backWall")) {
                //TODO
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
                CT_XXXChart c(chartType);
                if (c.read(reader))
                    subcharts << c;
                else {
                    qDebug() << "[debug] failed to load chart";
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

void CT_XXXChart::loadAreaChart(QXmlStreamReader &reader)
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
        reader.readNextStartElement();
        dropLines.read(reader);
    }
    else if (reader.name() == QLatin1String("axId"))
        axesIds << a.value(QLatin1String("val")).toInt();
    else if (reader.name() == QLatin1String("gapDepth"))
        parseAttributePercent(a, QLatin1String("val"), gapDepth);
    else reader.skipCurrentElement();
}

void CT_XXXChart::loadSurfaceChart(QXmlStreamReader &reader)
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

void CT_XXXChart::loadBubbleChart(QXmlStreamReader &reader)
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

void CT_XXXChart::loadPieChart(QXmlStreamReader &reader)
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

void CT_XXXChart::loadLineChart(QXmlStreamReader &reader)
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
        reader.readNextStartElement();
        dropLines.read(reader);
    }
    else if (reader.name() == QLatin1String("hiLowLines")) {
        reader.readNextStartElement();
        hiLowLines.read(reader);
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

void CT_XXXChart::loadBarChart(QXmlStreamReader &reader)
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

void CT_XXXChart::loadScatterChart(QXmlStreamReader &reader)
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

void CT_XXXChart::loadStockChart(QXmlStreamReader &reader)
{
    const auto &a = reader.attributes();
    if (reader.name() == QLatin1String("ser")) {
        Series ser(Series::Type::Line); ser.read(reader);
        seriesList << ser;
    }
    else if (reader.name() == QLatin1String("dropLines")) {
        reader.readNextStartElement();
        dropLines.read(reader);
    }
    else if (reader.name() == QLatin1String("hiLowLines"))
        hiLowLines.read(reader);
    else if (reader.name() == QLatin1String("dLbls"))
        labels.read(reader);
    else if (reader.name() == QLatin1String("upDownBars"))
        upDownBars.read(reader);
    else if (reader.name() == QLatin1String("axId"))
        axesIds << a.value(QLatin1String("val")).toInt();
    else reader.skipCurrentElement();
}

void CT_XXXChart::loadRadarChart(QXmlStreamReader &reader)
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
    if (title.isValid()) title.write(writer, QLatin1String("c:title"));
    writeEmptyElement(writer, QLatin1String("c:autoTitleDeleted"), autoTitleDeleted);

    writer.writeStartElement(QStringLiteral("c:plotArea"));

    layout.write(writer, QLatin1String("c:layout"));

    for (const auto &sub: qAsConst(subcharts)) {
        sub.write(writer);
    }

    for (const auto &axis: qAsConst(axisList)) {
        axis.write(writer);
    }

    if (dTable.has_value()) dTable->write(writer, QLatin1String("c:dTable"));
    if (plotAreaExtList.isValid()) plotAreaExtList.write(writer, QLatin1String("c:extLst"));

    if (plotAreaShape.isValid()) plotAreaShape.write(writer, "c:spPr");
    writer.writeEndElement(); // c:plotArea

    // c:legend
    if (legend.isValid()) legend.write(writer);

    writeEmptyElement(writer, QLatin1String("c:plotVisOnly"), plotVisOnly);
    if (dispBlanksAs.has_value()) {
        QString s; Chart::toString(dispBlanksAs.value(), s);
        writeEmptyElement(writer, QLatin1String("c:dispBlanksAs"), s);
    }
    writeEmptyElement(writer, QLatin1String("c:showDLblsOverMax"), showDLblsOverMax);

    writer.writeEndElement(); // c:chart
}

void CT_XXXChart::readBandFormats(QXmlStreamReader &reader)
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

void CT_XXXChart::saveAreaChart(QXmlStreamWriter &writer) const
{
    writer.writeStartElement(type == Chart::Type::Area3D ? QLatin1String("c:area3DChart") : QLatin1String("c:areaChart"));
    if (grouping.has_value()) {
        QString s; Chart::toString(grouping.value(), s);
        writeEmptyElement(writer, QLatin1String("c:grouping"), s);
    }
    writeEmptyElement(writer, QLatin1String("c:varyColors"), varyColors);
    for (const auto &ser: qAsConst(seriesList)) ser.write(writer);
    if (labels.isValid()) labels.write(writer);
    if (dropLines.isValid()) {
        writer.writeStartElement(QLatin1String("c:dropLines"));
        dropLines.write(writer, QLatin1String("c:spPr"));
        writer.writeEndElement();
    }
    for (const auto id: qAsConst(axesIds)) writeEmptyElement(writer, QLatin1String("c:axId"), id);
    if (gapDepth.has_value() && type == Chart::Type::Area3D)
        writeEmptyElement(writer, QLatin1String("c:gapDepth"), toST_PercentInt(gapDepth.value()));
    writer.writeEndElement();
}

void CT_XXXChart::saveSurfaceChart(QXmlStreamWriter &writer) const
{
    writer.writeStartElement(type == Chart::Type::Surface3D ? QLatin1String("c:surface3DChart") : QLatin1String("c:surfaceChart"));
    writeEmptyElement(writer, QLatin1String("c:wireframe"), wireframe);
    for (const auto &ser: qAsConst(seriesList)) ser.write(writer);
    saveBandFormats(writer);
    for (const auto id: qAsConst(axesIds)) writeEmptyElement(writer, QLatin1String("c:axId"), id);
    writer.writeEndElement();
}

void CT_XXXChart::saveBubbleChart(QXmlStreamWriter &writer) const
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
    for (const auto id: qAsConst(axesIds)) writeEmptyElement(writer, QLatin1String("c:axId"), id);

    writer.writeEndElement();
}

void CT_XXXChart::savePieChart(QXmlStreamWriter &writer) const
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
        if (splitType.has_value()) {
            QString s; Chart::toString(splitType.value(), s);
            writeEmptyElement(writer, QLatin1String("c:splitType"), s);
        }
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

void CT_XXXChart::saveLineChart(QXmlStreamWriter &writer) const
{
    writer.writeStartElement(type == Chart::Type::Line3D ? QLatin1String("c:line3DChart") : QLatin1String("c:lineChart"));
    if (grouping.has_value()) {
        QString s; Chart::toString(grouping.value(), s);
        writeEmptyElement(writer, QLatin1String("c:grouping"), s);
    }
    writeEmptyElement(writer, QLatin1String("c:varyColors"), varyColors);
    for (const auto &ser: qAsConst(seriesList)) ser.write(writer);
    if (labels.isValid()) labels.write(writer);
    if (dropLines.isValid()) {
        writer.writeStartElement(QLatin1String("c:dropLines"));
        dropLines.write(writer, QLatin1String("c:spPr"));
        writer.writeEndElement();
    }
    if (type == Chart::Type::Line) {
        if (hiLowLines.isValid()) {
            writer.writeStartElement(QLatin1String("c:hiLowLines"));
            hiLowLines.write(writer, QLatin1String("c:spPr"));
            writer.writeEndElement();
        }
        if (upDownBars.isValid()) upDownBars.write(writer, QLatin1String("c:upDownBars"));
        writeEmptyElement(writer, QLatin1String("c:marker"), marker);
        writeEmptyElement(writer, QLatin1String("c:smooth"), smooth);
    }
    if (gapDepth.has_value() && type == Chart::Type::Line3D)
        writeEmptyElement(writer, QLatin1String("c:gapDepth"), toST_PercentInt(gapDepth.value()));
    for (const auto id: qAsConst(axesIds)) writeEmptyElement(writer, QLatin1String("c:axId"), id);

    writer.writeEndElement();
}

void CT_XXXChart::saveBarChart(QXmlStreamWriter &writer) const
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
    if (barGrouping.has_value()) {
        QString s; Chart::toString(barGrouping.value(), s);
        writeEmptyElement(writer, QLatin1String("c:grouping"), s);
    }
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
    if (type == Chart::Type::Bar3D && barShape.has_value()) {
        QString s; Series::toString(barShape.value(), s);
        writeEmptyElement(writer, QLatin1String("c:shape"), s);
    }
    for (const auto id: qAsConst(axesIds)) writeEmptyElement(writer, QLatin1String("c:axId"), id);

    writer.writeEndElement();
}

void CT_XXXChart::saveScatterChart(QXmlStreamWriter &writer) const
{
    writer.writeStartElement(QLatin1String("c:scatterChart"));

    QString s; Chart::toString(scatterStyle, s);
    writeEmptyElement(writer, QLatin1String("c:scatterStyle"), s);
    writeEmptyElement(writer, QLatin1String("c:varyColors"), varyColors);
    for (const auto &ser: qAsConst(seriesList)) ser.write(writer);
    if (labels.isValid()) labels.write(writer);
    for (const auto id: qAsConst(axesIds)) writeEmptyElement(writer, QLatin1String("c:axId"), id);

    writer.writeEndElement();
}

void CT_XXXChart::saveStockChart(QXmlStreamWriter &writer) const
{
    writer.writeStartElement(QLatin1String("c:stockChart"));

    for (const auto &ser: qAsConst(seriesList)) ser.write(writer);
    if (labels.isValid()) labels.write(writer);
    if (dropLines.isValid()) {
        writer.writeStartElement(QLatin1String("c:dropLines"));
        dropLines.write(writer, QLatin1String("c:spPr"));
        writer.writeEndElement();
    }
    if (hiLowLines.isValid()) {
        writer.writeStartElement(QLatin1String("c:hiLowLines"));
        hiLowLines.write(writer, QLatin1String("c:spPr"));
        writer.writeEndElement();
    }
    if (upDownBars.isValid()) upDownBars.write(writer, QLatin1String("c:upDownBars"));
    for (const auto id: qAsConst(axesIds)) writeEmptyElement(writer, QLatin1String("c:axId"), id);

    writer.writeEndElement();
}

void CT_XXXChart::saveRadarChart(QXmlStreamWriter &writer) const
{
    writer.writeStartElement(QLatin1String("c:radarChart"));

    QString s; Chart::toString(radarStyle, s);
    writeEmptyElement(writer, QLatin1String("c:radarStyle"), s);
    writeEmptyElement(writer, QLatin1String("c:varyColors"), varyColors);
    for (const auto &ser: qAsConst(seriesList)) ser.write(writer);
    if (labels.isValid()) labels.write(writer);
    for (const auto id: qAsConst(axesIds)) writeEmptyElement(writer, QLatin1String("c:axId"), id);

    writer.writeEndElement();
}

void CT_XXXChart::saveBandFormats(QXmlStreamWriter &writer) const
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

CT_XXXChart::CT_XXXChart(Chart::Type type) : type{type}
{
    if (type == Chart::Type::Pie || type == Chart::Type::Pie3D ||
        type == Chart::Type::Doughnut || type == Chart::Type::OfPie ||
        type == Chart::Type::Bubble)
    varyColors = true;
}

Series *CT_XXXChart::addSeries(int index)
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

bool CT_XXXChart::read(QXmlStreamReader &reader)
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

void CT_XXXChart::write(QXmlStreamWriter &writer) const
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

std::optional<Chart::Grouping> Chart::grouping() const
{
    Q_D(const Chart);

    if (d->subcharts.isEmpty()) return {};
    return d->subcharts.last().grouping;
}

void Chart::setGrouping(Chart::Grouping grouping)
{
    Q_D(Chart);

    if (d->subcharts.isEmpty()) {
        d->subcharts << CT_XXXChart(d->chartType);
    }
    d->subcharts.last().grouping = grouping;
}

std::optional<Chart::BarGrouping> Chart::barGrouping() const
{
    Q_D(const Chart);

    if (d->subcharts.isEmpty()) return {};
    return d->subcharts.last().barGrouping;
}

void Chart::setBarGrouping(Chart::BarGrouping grouping)
{
    Q_D(Chart);

    if (d->subcharts.isEmpty()) {
        d->subcharts << CT_XXXChart(d->chartType);
    }
    d->subcharts.last().barGrouping = grouping;
}

std::optional<bool> Chart::varyColors() const
{
    Q_D(const Chart);

    if (d->subcharts.isEmpty()) return {};
    return d->subcharts.last().varyColors;
}

void Chart::setVaryColors(bool varyColors)
{
    Q_D(Chart);

    if (d->subcharts.isEmpty()) d->subcharts << CT_XXXChart(d->chartType);
    d->subcharts.last().varyColors = varyColors;
}

Labels Chart::labels() const
{
    Q_D(const Chart);

    if (d->subcharts.isEmpty()) return {};
    return d->subcharts.last().labels;
}

Labels &Chart::labels()
{
    Q_D(Chart);

    if (d->subcharts.isEmpty()) d->subcharts << CT_XXXChart(d->chartType);
    return d->subcharts.last().labels;
}

void Chart::setLabels(const Labels &labels)
{
    Q_D(Chart);

    if (d->subcharts.isEmpty()) d->subcharts << CT_XXXChart(d->chartType);
    d->subcharts.last().labels = labels;
}

ShapeFormat Chart::dropLines() const
{
    Q_D(const Chart);

    if (d->subcharts.isEmpty()) return {};
    return d->subcharts.last().dropLines;
}

ShapeFormat &Chart::dropLines()
{
    Q_D(Chart);

    if (d->subcharts.isEmpty()) d->subcharts << CT_XXXChart(d->chartType);
    return d->subcharts.last().dropLines;
}

void Chart::setDropLines(const ShapeFormat &dropLines)
{
    Q_D(Chart);

    if (d->subcharts.isEmpty()) d->subcharts << CT_XXXChart(d->chartType);
    d->subcharts.last().dropLines = dropLines;
}

ShapeFormat Chart::hiLowLines() const
{
    Q_D(const Chart);

    if (d->subcharts.isEmpty()) return {};
    return d->subcharts.last().hiLowLines;
}

ShapeFormat &Chart::hiLowLines()
{
    Q_D(Chart);

    if (d->subcharts.isEmpty()) d->subcharts << CT_XXXChart(d->chartType);
    return d->subcharts.last().hiLowLines;
}

void Chart::setHiLowLines(const ShapeFormat &hiLowLines)
{
    Q_D(Chart);

    if (d->subcharts.isEmpty()) d->subcharts << CT_XXXChart(d->chartType);
    d->subcharts.last().hiLowLines = hiLowLines;
}

UpDownBar Chart::upDownBars() const
{
    Q_D(const Chart);

    if (d->subcharts.isEmpty()) return {};
    return d->subcharts.last().upDownBars;
}

UpDownBar &Chart::upDownBars()
{
    Q_D(Chart);

    if (d->subcharts.isEmpty()) d->subcharts << CT_XXXChart(d->chartType);
    return d->subcharts.last().upDownBars;
}

void Chart::setUpDownBars(const UpDownBar &upDownBars)
{
    Q_D(Chart);

    if (d->subcharts.isEmpty()) d->subcharts << CT_XXXChart(d->chartType);
    d->subcharts.last().upDownBars = upDownBars;
}

std::optional<bool> Chart::markerShown() const
{
    Q_D(const Chart);
    if (d->subcharts.isEmpty()) return {};
    return d->subcharts.last().marker;
}

void Chart::setMarkerShown(bool markerShown)
{
    Q_D(Chart);

    if (d->subcharts.isEmpty()) d->subcharts << CT_XXXChart(d->chartType);
    d->subcharts.last().marker = markerShown;
}

std::optional<int> Chart::gapDepth() const
{
    Q_D(const Chart);
    if (d->subcharts.isEmpty()) return {};
    return d->subcharts.last().gapDepth;
}

void Chart::setGapDepth(int gapDepth)
{
    Q_D(Chart);

    if (d->subcharts.isEmpty()) d->subcharts << CT_XXXChart(d->chartType);
    d->subcharts.last().gapDepth = gapDepth;
}

Chart::ScatterStyle Chart::scatterStyle() const
{
    Q_D(const Chart);
    if (d->subcharts.isEmpty()) return {};
    return d->subcharts.last().scatterStyle;
}

void Chart::setScatterStyle(Chart::ScatterStyle scatterStyle)
{
    Q_D(Chart);

    if (d->subcharts.isEmpty()) d->subcharts << CT_XXXChart(d->chartType);
    d->subcharts.last().scatterStyle = scatterStyle;
}

Chart::RadarStyle Chart::radarStyle() const
{
    Q_D(const Chart);
    if (d->subcharts.isEmpty()) return {};
    return d->subcharts.last().radarStyle;
}

void Chart::setRadarStyle(Chart::RadarStyle radarStyle)
{
    Q_D(Chart);

    if (d->subcharts.isEmpty()) d->subcharts << CT_XXXChart(d->chartType);
    d->subcharts.last().radarStyle = radarStyle;
}

Chart::BarDirection Chart::barDirection() const
{
    Q_D(const Chart);
    if (d->subcharts.isEmpty()) return {};
    return d->subcharts.last().barDirection;
}

void Chart::setBarDirection(Chart::BarDirection barDirection)
{
    Q_D(Chart);

    if (d->subcharts.isEmpty()) d->subcharts << CT_XXXChart(d->chartType);
    d->subcharts.last().barDirection = barDirection;
}

std::optional<int> Chart::gapWidth() const
{
    Q_D(const Chart);
    if (d->subcharts.isEmpty()) return {};
    return d->subcharts.last().gapWidth;
}

void Chart::setGapWidth(int gapWidth)
{
    Q_D(Chart);

    if (d->subcharts.isEmpty()) d->subcharts << CT_XXXChart(d->chartType);
    d->subcharts.last().gapWidth = gapWidth;
}

std::optional<int> Chart::overlap() const
{
    Q_D(const Chart);
    if (d->subcharts.isEmpty()) return {};
    return d->subcharts.last().overlap;
}

void Chart::setOverlap(int overlap)
{
    Q_D(Chart);

    if (d->subcharts.isEmpty()) d->subcharts << CT_XXXChart(d->chartType);
    d->subcharts.last().overlap = overlap;
}

QList<ShapeFormat> Chart::seriesLines() const
{
    Q_D(const Chart);
    if (d->subcharts.isEmpty()) return {};
    return d->subcharts.last().serLines;
}

QList<ShapeFormat> &Chart::seriesLines()
{
    Q_D(Chart);

    if (d->subcharts.isEmpty()) d->subcharts << CT_XXXChart(d->chartType);
    return d->subcharts.last().serLines;
}

void Chart::setSeriesLines(const QList<ShapeFormat> &seriesLines)
{
    Q_D(Chart);

    if (d->subcharts.isEmpty()) d->subcharts << CT_XXXChart(d->chartType);
    d->subcharts.last().serLines = seriesLines;
}

std::optional<Series::BarShape> Chart::barShape() const
{
    Q_D(const Chart);
    if (d->subcharts.isEmpty()) return {};
    return d->subcharts.last().barShape;
}

void Chart::setBarShape(Series::BarShape barShape)
{
    Q_D(Chart);

    if (d->subcharts.isEmpty()) d->subcharts << CT_XXXChart(d->chartType);
    d->subcharts.last().barShape = barShape;
}

std::optional<int> Chart::firstSliceAngle() const
{
    Q_D(const Chart);
    if (d->subcharts.isEmpty()) return {};
    return d->subcharts.last().firstSliceAng;
}

void Chart::setFirstSliceAngle(int angle)
{
    Q_D(Chart);

    if (d->subcharts.isEmpty()) d->subcharts << CT_XXXChart(d->chartType);
    d->subcharts.last().firstSliceAng = angle;
}

std::optional<int> Chart::holeSize() const
{
    Q_D(const Chart);
    if (d->subcharts.isEmpty()) return {};
    return d->subcharts.last().holeSize;
}

void Chart::setHoleSise(int holeSize)
{
    Q_D(Chart);

    if (d->subcharts.isEmpty()) d->subcharts << CT_XXXChart(d->chartType);
    d->subcharts.last().holeSize = holeSize;
}

Chart::OfPieType Chart::ofPieType() const
{
    Q_D(const Chart);
    if (d->subcharts.isEmpty()) return {};
    return d->subcharts.last().ofPieType;
}

void Chart::setOfPieType(OfPieType type)
{
    Q_D(Chart);

    if (d->subcharts.isEmpty()) d->subcharts << CT_XXXChart(d->chartType);
    d->subcharts.last().ofPieType = type;
}

std::optional<Chart::SplitType> Chart::splitType() const
{
    Q_D(const Chart);
    if (d->subcharts.isEmpty()) return {};
    return d->subcharts.last().splitType;
}

void Chart::setSplitType(SplitType splitType)
{
    Q_D(Chart);

    if (d->subcharts.isEmpty()) d->subcharts << CT_XXXChart(d->chartType);
    d->subcharts.last().splitType = splitType;
}

std::optional<double> Chart::splitPos() const
{
    Q_D(const Chart);
    if (d->subcharts.isEmpty()) return {};
    return d->subcharts.last().splitPos;
}

void Chart::setSplitPos(double value)
{
    Q_D(Chart);

    if (d->subcharts.isEmpty()) d->subcharts << CT_XXXChart(d->chartType);
    d->subcharts.last().splitPos = value;
}

QList<int> Chart::customSplit() const
{
    Q_D(const Chart);
    if (d->subcharts.isEmpty()) return {};
    return d->subcharts.last().customSplit;
}

void Chart::setCustomSplit(const QList<int> &indexes)
{
    Q_D(Chart);

    if (d->subcharts.isEmpty()) d->subcharts << CT_XXXChart(d->chartType);
    d->subcharts.last().customSplit = indexes;
}

std::optional<int> Chart::secondPieSize() const
{
    Q_D(const Chart);
    if (d->subcharts.isEmpty()) return {};
    return d->subcharts.last().secondPieSize;
}

void Chart::setSecondPieSize(int size)
{
    Q_D(Chart);

    if (d->subcharts.isEmpty()) d->subcharts << CT_XXXChart(d->chartType);
    d->subcharts.last().secondPieSize = size;
}

std::optional<bool> Chart::bubble3D() const
{
    Q_D(const Chart);
    if (d->subcharts.isEmpty()) return {};
    return d->subcharts.last().bubble3D;
}

void Chart::setBubble3D(bool bubble3D)
{
    Q_D(Chart);

    if (d->subcharts.isEmpty()) d->subcharts << CT_XXXChart(d->chartType);
    d->subcharts.last().bubble3D = bubble3D;
}

std::optional<bool> Chart::showNegativeBubbles() const
{
    Q_D(const Chart);
    if (d->subcharts.isEmpty()) return {};
    return d->subcharts.last().showNegBubbles;
}

void Chart::setShowNegativeBubbles(bool show)
{
    Q_D(Chart);

    if (d->subcharts.isEmpty()) d->subcharts << CT_XXXChart(d->chartType);
    d->subcharts.last().showNegBubbles = show;
}

std::optional<int> Chart::bubbleScale() const
{
    Q_D(const Chart);
    if (d->subcharts.isEmpty()) return {};
    return d->subcharts.last().bubbleScale;
}

void Chart::setBubbleScale(int scale)
{
    Q_D(Chart);

    if (d->subcharts.isEmpty()) d->subcharts << CT_XXXChart(d->chartType);
    d->subcharts.last().bubbleScale = scale;
}

std::optional<Chart::BubbleSizeRepresents> Chart::bubbleSizeRepresents() const
{
    Q_D(const Chart);
    if (d->subcharts.isEmpty()) return {};
    return d->subcharts.last().bubbleSizeRepresents;
}

void Chart::setBubbleSizeRepresents(BubbleSizeRepresents value)
{
    Q_D(Chart);

    if (d->subcharts.isEmpty()) d->subcharts << CT_XXXChart(d->chartType);
    d->subcharts.last().bubbleSizeRepresents = value;
}

std::optional<bool> Chart::wireframe() const
{
    Q_D(const Chart);
    if (d->subcharts.isEmpty()) return {};
    return d->subcharts.last().wireframe;
}

void Chart::setWireframe(bool wireframe)
{
    Q_D(Chart);

    if (d->subcharts.isEmpty()) d->subcharts << CT_XXXChart(d->chartType);
    d->subcharts.last().wireframe = wireframe;
}

QMap<int, ShapeFormat> Chart::bandFormats() const
{
    Q_D(const Chart);
    if (d->subcharts.isEmpty()) return {};
    return d->subcharts.last().bandFormats;
}

void Chart::setBandFormats(QMap<int, ShapeFormat> bandFormats)
{
    Q_D(Chart);

    if (d->subcharts.isEmpty()) d->subcharts << CT_XXXChart(d->chartType);
    d->subcharts.last().bandFormats = bandFormats;
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

void Chart::registerBlips(Workbook *workbook)
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

    for (auto &fill: fills) {
        if (fill.get().type() == FillFormat::FillType::BlipFill) {
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

    //writer.writeEmptyElement(QStringLiteral("picture"));
    //writer.writeAttribute(QStringLiteral("r:id"), QStringLiteral("rId%1").arg(d->relationships->count()));
}

int Chart::registerBlip(Workbook *workbook, const QImage &image)
{
    Q_D(Chart);

    QByteArray ba;
    QBuffer buffer(&ba);
    buffer.open(QIODevice::WriteOnly);
    image.save(&buffer, "PNG");

    auto pictureFile = QSharedPointer<MediaFile>(new MediaFile(ba, QStringLiteral("png"), QStringLiteral("image/png")));
    workbook->addMediaFile(pictureFile);


    return d->relationships->count();
}

}



