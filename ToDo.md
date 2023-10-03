
# What is done in the QXlsx fork

1. Full support of fill properties.
2. Full support of line properties.
3. Full support of CT_Title and CT_Text, including text body properties, paragraph 
   properties, character properties.
4. Full support of all chart types and series types.
5. Axes, legends, etc.

# What is yet to be done:

## Charts

1. Add printSettings
2. Add user shapes
3. Add externalData
4. Add protection
5. Add pivotSource
6. Add clrMapOvr
7. Add pivotFmts
8. Add view3D
9. Add floor and walls

### Shapes

1. Add custom geometry shapes.
2. Add Effects DAG support.
3. Add methods to add effects directly to shapes.
4. add methods to create a shape with line, fill, shape

### Fills

1. Add methods to create simple fills.

### Text 

1. Write method of exporting to html.

## Sheets

1. Add support for custom sheet views.
2. Add support for web publish items.
3. Add support for drawings in header/footer.
4. Add support for custom printers via relations.
5. Add methods to fine-tune header/footer (f.e. addPageNumber(Footer::Right))

## Worksheets

1. Overhaul rich strings support.
2. Rewrite column properties to use interval maps.

### Autofiltering

1. Add support for filtering by icons
2. Add support to filtering by color
3. To filtering by values add support of CT_Filters attributes.
4. Add support for sorting.
5. Check string filtering and add the valid help.

### SheetView

1. CT_Pane
2. CT_PivotSelection

### sheetCalcPr

### protectedRanges

### scenarios

### sortState

### dataConsolidate

### customSheetViews

### phoneticPr

### printOptions

### rowBreaks

### colBreaks

### customProperties

### cellWatches

### ignoredErrors

### smartTags

### drawingHF

### oleObjects

### controls

### webPublishItems

### tableParts

## Text

1. Add hyperlinks support


## Global

1. Add support for extension lists.
2. Rewrite all examples and tests, as they duplicate each other.
3. Add documentation to all new classes.
4. Add tests.
5. Write NumberFormat class to easily create and validate number formats.
6. Convert to implicitly shareable classes that allow it.

### Examples

1. Separate testing examples from demonstration examples.
2. Add docs to examples.