// xlsxchart_p.h

#ifndef QXLSX_CHART_P_H
#define QXLSX_CHART_P_H

#include <QObject>
#include <QString>
#include <QVector>
#include <QMap>
#include <QList>
#include <QXmlStreamReader>
#include <QXmlStreamWriter>

#include <memory>

#include "xlsxabstractooxmlfile_p.h"
#include "xlsxchart.h"
#include "xlsxseries.h"

QT_BEGIN_NAMESPACE_XLSX

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
    Text textProperties;
//    std::optional<ExternalData> externalData;
//    std::optional<PrintSettings> printSettings;
//    std::optional<Reference> userShapes;
    ExtensionList chartSpaceExtList;

    ///CT_Chart properties
    Title title;
    std::optional<bool> autoTitleDeleted; //deleteAutoGeneratedTitle
//    QList<PivotFormat> pivotFormats;
//    std::optional<View3D> view3D;
//    std::optional<Surface> floor;
//    std::optional<Surface> sideWall;
//    std::optional<Surface> backWall;
    Legend legend;
    std::optional<bool> plotVisOnly; //plotOnlyVisibleCells
    std::optional<Chart::DisplayBlanksAs> dispBlanksAs;
    std::optional<bool> showDLblsOverMax; //showDataLabelsOverMaximum
    ExtensionList chartExtList;

    ///CT_PlotArea properties
    Layout layout;
    QList<CT_XXXChart> subcharts;
    QList<Axis> axisList;
    std::optional<DataTable> dTable;
    ShapeFormat plotAreaShape;
    ExtensionList plotAreaExtList;

    AbstractSheet* sheet;
};

QT_END_NAMESPACE_XLSX

#endif // QXLSX_CHART_P_H
