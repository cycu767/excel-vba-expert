# Charts and dependency visualizations

Evidence labels: **[DOC]** documented behavior; **[ENG]** skill engineering rule; **[INFERENCE]** derived interpretation; **[UNVERIFIED]** needs target-environment verification.

## Choose the correct model
[DOC: MS-CHART-ADD] ChartObjects.Add creates an embedded ChartObject with position/size in points. [ENG] Reach the chart through `.Chart`; distinguish it from a chart sheet in Charts. Confirm the chart type supports the requested axes, labels and series operation. A donut chart does not have the same axis contract as an XY chart.

[DOC: MS-GRAPH] The separate Graph Visual Basic reference is not Microsoft Graph REST. [ENG] A dependency graph is a node/edge model, not automatically an Excel chart. Do not import Graph methods into Excel.Chart because the names sound similar. The legacy Graph API is only routed/indexed here, not exhaustively audited.

## Stable update, not accumulation
[ENG] Give the output chart a stable, project-owned name and locate it on an explicit worksheet. Update it in place or deliberately replace only that named chart. Never delete all Shapes or ChartObjects as a convenient reset. Validate source headers, category values, series lengths and empty-data behavior before changing the chart. Decide whether the graph should show a clear empty state or preserve the previous visualization marked stale.

[DOC: MS-CHART-SOURCE, MS-NEWSERIES] SetSourceData takes a Range and orientation; NewSeries is unavailable for PivotCharts. [ENG] Use explicit source/series assignments when the default inference could guess the wrong categories. Do not force a pivot series workflow onto an ordinary chart or vice versa. See `MCharts.bas` for an existing-chart update pattern, not a universal dashboard engine.

## Presentation semantics
[ENG] Choose a chart that preserves the meaning of the data: time ordering, absolute versus percentage values, missing versus zero values, meaningful axis labels and units. Avoid automatically applying a second axis or truncated scales that distort comparison. Document whether totals, filters and hidden rows are included. Test chart output after source table growth and shrinkage, not just creation.

[DOC: MS-CHART-EXPORT] Export returns Boolean and uses an available filter. [ENG] Check its result, output existence and destination overwrite policy. An exported image alone cannot prove the source formulas or chart links are correct.

## When the requested view is a graph
[ENG] Establish node identity, directed/undirected edges, duplicate-edge policy, weights, cycle handling and expected scale. Keep graph visualization separate from the cost/allocation engine. For an HTML view, make the data handoff explicit, escape labels and avoid loading external scripts/CDNs without permission. For MSForms, explain the rendering and interaction limits instead of inventing a canvas component.

## Tests
[ENG] Zero/one/many records, dates versus text categories, missing values, filters, protected sheet, repeat refresh without duplicate chart objects, source range reduction, export failure and target-version chart support. Graph workflows also need cycle/duplicate tests and reconciled edge totals when those are domain requirements.

## Source register

- [MS-CHART-ADD: ChartObjects.Add](https://learn.microsoft.com/en-us/office/vba/api/excel.chartobjects.add)
- [MS-CHART-SOURCE: Chart.SetSourceData](https://learn.microsoft.com/en-us/office/vba/api/excel.chart.setsourcedata)
- [MS-NEWSERIES: SeriesCollection.NewSeries](https://learn.microsoft.com/en-us/office/vba/api/excel.seriescollection.newseries)
- [MS-CHART-EXPORT: Chart.Export](https://learn.microsoft.com/en-us/office/vba/api/excel.chart.export)
- [MS-GRAPH: Excel Graph Visual Basic reference](https://learn.microsoft.com/en-us/office/vba/api/overview/excel/graph-visual-basic-reference)
