// xlsxchartsheet.cpp

#include <QtGlobal>
#include <QXmlStreamReader>
#include <QXmlStreamWriter>
#include <QDir>

#include "xlsxchartsheet.h"
#include "xlsxchartsheet_p.h"
#include "xlsxworkbook.h"
#include "xlsxutility_p.h"
#include "xlsxdrawing_p.h"
#include "xlsxdrawinganchor_p.h"
#include "xlsxchart.h"

namespace QXlsx {

ChartsheetPrivate::ChartsheetPrivate(Chartsheet *p, Chartsheet::CreateFlag flag)
    : AbstractSheetPrivate(p, flag), chart(0)
{

}

ChartsheetPrivate::~ChartsheetPrivate()
{
}

/*!
 * \internal
 */
Chartsheet::Chartsheet(const QString &name, int id, Workbook *workbook, CreateFlag flag)
    : AbstractSheet( name, id, workbook, new ChartsheetPrivate(this, flag) )
{
    setType(Type::Chartsheet);

    if (flag == Chartsheet::F_NewFromScratch)
    {
        d_func()->drawing = std::make_shared<Drawing>(this, flag);

        DrawingAbsoluteAnchor *anchor = new DrawingAbsoluteAnchor(drawing(), DrawingAnchor::Picture);

        anchor->pos = QPoint(0, 0);
        anchor->ext = qMakePair(Coordinate(9293679ll), Coordinate(6068786ll)); //TODO: why these magic numbers?

        QSharedPointer<Chart> chart = QSharedPointer<Chart>(new Chart(this, flag));
        anchor->setObjectGraphicFrame(chart);

        d_func()->chart = chart.data();
    }
}

/*!
 * \internal
 */
Chartsheet *Chartsheet::copy(const QString &distName, int distId) const
{
    Q_D(const Chartsheet);
    Chartsheet *sheet = new Chartsheet(distName, distId, d->workbook, F_NewFromScratch);
    ChartsheetPrivate *sheet_d = sheet->d_func();

    *sheet_d->chart = *d->chart;
    sheet_d->sheetState = d->sheetState;
    sheet_d->type = d->type;
    sheet_d->headerFooter = d->headerFooter;
    sheet_d->pageMargins = d->pageMargins;
    sheet_d->pageSetup = d->pageSetup;
    sheet_d->pictureFile = d->pictureFile;
    sheet_d->sheetProtection = d->sheetProtection;
    sheet_d->extLst = d->extLst;
    sheet_d->sheetProperties = d->sheetProperties;

    return sheet;
}

std::optional<bool> Chartsheet::zoomToFit() const
{
    Q_D(const Chartsheet);
    if (d->sheetViews.isEmpty()) return {};
    return d->sheetViews.last().zoomToFit;
}

void Chartsheet::setZoomToFit(bool zoom)
{
    Q_D(Chartsheet);
    if (d->sheetViews.isEmpty()) d->sheetViews << SheetView();
    d->sheetViews.last().zoomToFit = zoom;
}

Chartsheet::~Chartsheet()
{
}

Chart *Chartsheet::chart()
{
    Q_D(Chartsheet);

    if (d->drawing && !d->drawing->anchors.isEmpty())
        return d->drawing->anchors.first()->chart().get();
    return nullptr;
}

void Chartsheet::saveToXmlFile(QIODevice *device) const
{
    Q_D(const Chartsheet);
    d->relationships->clear();

    QXmlStreamWriter writer(device);

    writer.writeStartDocument(QStringLiteral("1.0"), true);
    writer.writeDefaultNamespace(QStringLiteral("http://schemas.openxmlformats.org/spreadsheetml/2006/main"));
    writer.writeNamespace(QStringLiteral("http://schemas.openxmlformats.org/officeDocument/2006/relationships"), QStringLiteral("r"));
    writer.writeStartElement(QStringLiteral("chartsheet"));

    //1. sheetPr - CT_SheetPr
    if (d->sheetProperties.isValid()) {
        writer.writeStartElement(QLatin1String("sheetPr"));
        writeAttribute(writer, QLatin1String("published"), d->sheetProperties.published);
        writeAttribute(writer, QLatin1String("codeName"), d->sheetProperties.codeName);
        if (d->sheetProperties.tabColor.isValid()) d->sheetProperties.tabColor.write(writer, QLatin1String("tabColor"));
        writer.writeEndElement();
    }
    //2. sheetViews
    d->saveXmlSheetViews(writer, false);
    //3. sheet protection
    d->sheetProtection.write(writer, true);
    //4. customSheetViews
    //TODO: customSheetViews
    //5. pageMargins
    d->pageMargins.write(writer);
    //6. pageSetup
    d->pageSetup.writeChartsheet(writer);
    //7. headerFooter
    d->headerFooter.write(writer, QLatin1String("headerFooter"));
    //8. drawing
    d->saveXmlDrawings(writer);
    //9. drawindHF
    //TODO: drawindHF
    //10. picture
    d->saveXmlPicture(writer);
    //11. webPublishItems
    //TODO: webPublishItems
    //12. extLst
    d->extLst.write(writer, "extLst");

    writer.writeEndElement();//chartsheet
    writer.writeEndDocument();
}

bool Chartsheet::loadFromXmlFile(QIODevice *device)
{
    Q_D(Chartsheet);

    //TODO: rewrite
    QXmlStreamReader reader(device);
    while (!reader.atEnd()) {
        reader.readNextStartElement();
        if (reader.tokenType() == QXmlStreamReader::StartElement) {
            if (reader.name() == QLatin1String("drawing"))
                d->loadXmlDrawing(reader);
            else if (reader.name() == QLatin1String("drawingHF")) {
                //TODO: drawingHF
            }
            else if (reader.name() == QLatin1String("sheetPr"))
                d->sheetProperties.read(reader);
            else if (reader.name() == QLatin1String("sheetProtection")) {
                SheetProtection s;
                s.read(reader);
                d->sheetProtection = s;
            }
            else if (reader.name() == QLatin1String("pageMargins"))
                d->pageMargins.read(reader);
            else if (reader.name() == QLatin1String("pageSetup"))
                d->pageSetup.read(reader);
            else if (reader.name() == QLatin1String("headerFooter"))
                d->headerFooter.read(reader);
            else if (reader.name() == QLatin1String("picture"))
                d->loadXmlPicture(reader);
            else if (reader.name() == QLatin1String("webPublishItems")) {
                //TODO: webPublishItems
            }
            else if (reader.name() == QLatin1String("sheetViews"))
                d->loadXmlSheetViews(reader);
            else if (reader.name() == QLatin1String("customSheetViews")) {
                // TODO: customSheetViews
            }
            else if (reader.name() == QLatin1String("extLst"))
                d->extLst.read(reader);
        }
    }

    return true;
}

}
