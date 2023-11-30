#include "xlsxsheetview.h"

namespace QXlsx {

void SheetView::read(QXmlStreamReader &reader)
{
    const auto &name = reader.name();
    auto a = reader.attributes();
    parseAttributeBool(a, QLatin1String("windowProtection"), windowProtection);
    parseAttributeBool(a, QLatin1String("showFormulas"), showFormulas);
    parseAttributeBool(a, QLatin1String("showGridLines"), showGridLines);
    parseAttributeBool(a, QLatin1String("showRowColHeaders"), showRowColHeaders);
    parseAttributeBool(a, QLatin1String("showZeros"), showZeros);
    parseAttributeBool(a, QLatin1String("rightToLeft"), rightToLeft);
    parseAttributeBool(a, QLatin1String("tabSelected"), tabSelected);
    parseAttributeBool(a, QLatin1String("showRuler"), showRuler);
    parseAttributeBool(a, QLatin1String("showOutlineSymbols"), showOutlineSymbols);
    parseAttributeBool(a, QLatin1String("showWhiteSpace"), showWhiteSpace);
    parseAttributeBool(a, QLatin1String("defaultGridColor"), defaultGridColor);
    if (a.hasAttribute(QLatin1String("view"))) {
        auto s = a.value(QLatin1String("view"));
        if (s == QLatin1String("normal")) type = Type::Normal;
        else if (s == QLatin1String("pageBreakPreview")) type = Type::PageBreakPreview;
        else if (s == QLatin1String("pageLayout")) type = Type::PageLayout;
    }
    if (a.hasAttribute(QLatin1String("topLeftCell")))
        topLeftCell = CellReference(a.value(QLatin1String("topLeftCell")).toString());
    parseAttributeInt(a, QLatin1String("colorId"), colorId);
    parseAttributeInt(a, QLatin1String("zoomScale"), zoomScale);
    parseAttributeInt(a, QLatin1String("zoomScaleNormal"), zoomScaleNormal);
    parseAttributeInt(a, QLatin1String("zoomScaleSheetLayoutView"), zoomScalePageBreakView);
    parseAttributeInt(a, QLatin1String("zoomScalePageLayoutView"), zoomScalePageLayoutView);
    parseAttributeInt(a, QLatin1String("workbookViewId"), workbookViewId);

    while (!reader.atEnd()) {
        const auto token = reader.readNext();
        if (token == QXmlStreamReader::StartElement) {
            a = reader.attributes();
            if (reader.name() == QLatin1String("pane")) {
                parseAttributeDouble(a, QLatin1String("xSplit"), pane.xSplit);
                parseAttributeDouble(a, QLatin1String("ySplit"), pane.ySplit);
                if (a.hasAttribute(QLatin1String("topLeftCell")))
                    pane.topLeftCell = CellReference(a.value(QLatin1String("topLeftCell")).toString());
                if (a.hasAttribute(QLatin1String("activePane"))) {
                    ViewPane::Type t;
                    ViewPane::fromString(a.value(QLatin1String("activePane")).toString(), t);
                    pane.activePane = t;
                }
                if (a.hasAttribute(QLatin1String("state"))) {
                    ViewPane::State t;
                    ViewPane::fromString(a.value(QLatin1String("state")).toString(), t);
                    pane.paneState = t;
                }
            }
            else if (reader.name() == QLatin1String("selection")) {
                selection.read(reader);
            }
            else if (reader.name() == QLatin1String("pivotSelection")) {
                //TODO://    std::optional<CT_PivotSelection> pivotSelection;
                reader.skipCurrentElement();
            }
            else if (reader.name() == QLatin1String("extLst")) {
                extLst.read(reader);
            }
            else reader.skipCurrentElement();
        }
        else if (token == QXmlStreamReader::EndElement && reader.name() == name)
            break;
    }
}

void SheetView::write(QXmlStreamWriter &writer, const QLatin1String &name, bool worksheet) const
{
    writer.writeStartElement(name);
    if (worksheet) writeAttribute(writer, QLatin1String("windowProtection"), windowProtection);
    if (worksheet) writeAttribute(writer, QLatin1String("showFormulas"), showFormulas);
    if (worksheet) writeAttribute(writer, QLatin1String("showGridLines"), showGridLines);
    if (worksheet) writeAttribute(writer, QLatin1String("showRowColHeaders"), showRowColHeaders);
    if (worksheet) writeAttribute(writer, QLatin1String("showZeros"), showZeros);
    if (worksheet) writeAttribute(writer, QLatin1String("rightToLeft"), rightToLeft);
    writeAttribute(writer, QLatin1String("tabSelected"), tabSelected);
    if (worksheet) writeAttribute(writer, QLatin1String("showRuler"), showRuler);
    if (worksheet) writeAttribute(writer, QLatin1String("showOutlineSymbols"), showOutlineSymbols);
    if (worksheet) writeAttribute(writer, QLatin1String("showWhiteSpace"), showWhiteSpace);
    if (worksheet) writeAttribute(writer, QLatin1String("defaultGridColor"), defaultGridColor);
    if (worksheet) switch (type.value_or(Type::Normal)) {
        case Type::Normal: break;
        case Type::PageLayout: writer.writeAttribute(QLatin1String("view"), QLatin1String("pageLayout")); break;
        case Type::PageBreakPreview: writer.writeAttribute(QLatin1String("view"), QLatin1String("pageBreakPreview")); break;
    }
    if (worksheet && topLeftCell.isValid()) writer.writeAttribute(QLatin1String("topLeftCell"), topLeftCell.toString());
    if (worksheet) writeAttribute(writer, QLatin1String("colorId"), colorId);
    writeAttribute(writer, QLatin1String("zoomScale"), zoomScale);
    if (worksheet) writeAttribute(writer, QLatin1String("zoomScaleNormal"), zoomScaleNormal);
    if (worksheet) writeAttribute(writer, QLatin1String("zoomScaleSheetLayoutView"), zoomScalePageBreakView);
    if (worksheet) writeAttribute(writer, QLatin1String("zoomScalePageLayoutView"), zoomScalePageLayoutView);
    writeAttribute(writer, QLatin1String("workbookViewId"), workbookViewId);
    if (!worksheet) writeAttribute(writer, QLatin1String("zoomToFit"), zoomToFit.value_or(true));

    if (worksheet && pane.isValid()) {
        writer.writeEmptyElement(QLatin1String("pane"));
        writeAttribute(writer, QLatin1String("xSplit"), pane.xSplit);
        writeAttribute(writer, QLatin1String("ySplit"), pane.ySplit);
        writeAttribute(writer, QLatin1String("topLeftCell"), pane.topLeftCell.toString());
        if (pane.activePane.has_value())
            writeAttribute(writer, QLatin1String("activePane"), ViewPane::toString(pane.activePane.value()));
        if (pane.paneState.has_value())
            writeAttribute(writer, QLatin1String("state"), ViewPane::toString(pane.paneState.value()));
    }
    if (worksheet && selection.isValid()) selection.write(writer, QLatin1String("selection"));
    //TODO:    std::optional<CT_PivotSelection> pivotSelection;

    extLst.write(writer, QLatin1String("extLst"));
    writer.writeEndElement();
}

bool SheetView::isValid() const
{
    if (windowProtection.has_value()) return true;
    if (showFormulas.has_value()) return true;
    if (showGridLines.has_value()) return true;
    if (showRowColHeaders.has_value()) return true;
    if (showZeros.has_value()) return true;
    if (rightToLeft.has_value()) return true;
    if (tabSelected.has_value()) return true;
    if (showRuler.has_value()) return true;
    if (showOutlineSymbols.has_value()) return true;
    if (showWhiteSpace.has_value()) return true;
    if (defaultGridColor.has_value()) return true;
    if (type.has_value()) return true;
    if (topLeftCell.isValid()) return true;
    if (colorId.has_value()) return true;
    if (zoomScale.has_value()) return true;
    if (zoomScaleNormal.has_value()) return true;
    if (zoomScalePageBreakView.has_value()) return true;
    if (zoomScalePageLayoutView.has_value()) return true;
    if (workbookViewId >= 0) return true; //required
    if (selection.isValid()) return true;

    //TODO:
//    std::optional<CT_Pane> pane;
//    std::optional<CT_PivotSelection> pivotSelection;

    if (extLst.isValid()) return true;
    return false;
}

bool Selection::isValid() const
{
    if (pane.has_value()) return true;
    if (activeCell.isValid()) return true;
    if (activeCellId.has_value()) return true;
    if (!selectedRanges.isEmpty()) return true;
    return false;
}

void Selection::read(QXmlStreamReader &reader)
{
    const auto &a = reader.attributes();
    if (a.hasAttribute(QLatin1String("pane"))) {
        QString s = a.value(QLatin1String("pane")).toString();
        ViewPane::Type t; ViewPane::fromString(s, t); pane = t;
    }
    if (a.hasAttribute(QLatin1String("activeCell"))) {
        QString s = a.value(QLatin1String("activeCell")).toString();
        activeCell = CellReference::fromString(s);
    }
    parseAttributeInt(a, QLatin1String("activeCellId"), activeCellId);
    if (a.hasAttribute(QLatin1String("sqref"))) {
        QStringList ranges = a.value(QLatin1String("sqref")).toString().split(' ');
        for (auto range: ranges) selectedRanges << CellRange(range);
    }
}

void Selection::write(QXmlStreamWriter &writer, const QLatin1String &name) const
{
    if (!isValid()) return;
    writer.writeEmptyElement(name);
    if (pane.has_value()) {
        writer.writeAttribute(QLatin1String("pane"), ViewPane::toString(pane.value()));
    }
    auto ref = activeCell;
    if (!ref.isValid()) {
        if (!selectedRanges.isEmpty()) ref = selectedRanges.first().topLeft();
        else ref = CellReference(1,1);
    }
    writer.writeAttribute(QLatin1String("activeCell"), ref.toString());
    if (selectedRanges.size() > 1) {
        //find the index of range that contains activeCell and write it
        for (int idx = 0; idx < selectedRanges.size(); ++idx) {
            if (selectedRanges.at(idx).contains(ref)) {
                writeAttribute(writer, QLatin1String("activeCellId"), idx);
                break;
            }
        }
    }
    if (!selectedRanges.isEmpty()) {
        QStringList ranges;
        for (const auto &range: selectedRanges) ranges << range.toString();
        writeAttribute(writer, QLatin1String("sqref"), ranges.join(' '));
    }
}

bool ViewPane::isValid() const
{
    if (xSplit.has_value()) return true;
    if (ySplit.has_value()) return true;
    if (activePane.has_value()) return true;
    if (paneState.has_value()) return true;
    if (topLeftCell.isValid()) return true;
    return false;
}

}


