---
title: Create and customize Excel charts with JavaScript
description: Learn how to create Excel charts, add series, format titles and axes, add trendlines and data tables, and export chart images with the JavaScript API.
ms.date: 06/03/2026
ms.topic: how-to
ms.localizationpriority: medium
ai-usage: ai-assisted
---

# Create and customize charts with the Excel JavaScript API

Use charts when your add-in needs to turn worksheet data into a visual summary. This article shows how to create a chart from a range, add a series, update titles and axes, control gridlines, add trendlines and a data table, and export the chart as an image.

For the full API surface, see [Chart object](/javascript/api/excel/excel.chart) and [ChartCollection object](/javascript/api/excel/excel.chartcollection).

## Create a chart from a range

Charts usually start with data that already exists in a range or table. In this example, the add-in creates a **Line** chart on the **Sample** worksheet from the range **A1:B13** and then applies a few common formatting settings.

```js
await Excel.run(async (context) => {
    const sheet = context.workbook.worksheets.getItem("Sample");
    const dataRange = sheet.getRange("A1:B13");
    const chart = sheet.charts.add(
        Excel.ChartType.line,
        dataRange,
        Excel.ChartSeriesBy.auto
    );

    chart.title.text = "Sales Data";
    chart.legend.position = Excel.ChartLegendPosition.right;
    chart.legend.format.fill.setSolidColor("white");
    chart.dataLabels.format.font.size = 15;
    chart.dataLabels.format.font.color = "black";

    await context.sync();
});
```

After the code runs, the worksheet contains a new line chart.

:::image type="content" source="../images/excel-charts-create-line.png" alt-text="New line chart in Excel.":::

## Add a data series

Use an additional series when your add-in needs to compare a new column of values with the existing chart. In this example, the add-in adds the **2016** series from the range **D2:D5** to the first chart on the worksheet.

```js
await Excel.run(async (context) => {
    const sheet = context.workbook.worksheets.getItem("Sample");
    const chart = sheet.charts.getItemAt(0);
    const dataRange = sheet.getRange("D2:D5");

    const newSeries = chart.series.add("2016");
    newSeries.setValues(dataRange);

    await context.sync();
});
```

Before you add the series, the chart looks like this.

:::image type="content" source="../images/excel-charts-data-series-before.png" alt-text="Chart in Excel before the 2016 data series is added.":::

After you add the series, the chart includes the new data.

:::image type="content" source="../images/excel-charts-data-series-after.png" alt-text="Chart in Excel after the 2016 data series is added.":::

## Set the chart title

A clear title helps users understand what the chart shows without inspecting the source data. The following example sets the title of the first chart on the worksheet to **Sales Data by Year**.

```js
await Excel.run(async (context) => {
    const sheet = context.workbook.worksheets.getItem("Sample");
    const chart = sheet.charts.getItemAt(0);

    chart.title.text = "Sales Data by Year";

    await context.sync();
});
```

:::image type="content" source="../images/excel-charts-title-set.png" alt-text="Chart with a title in Excel.":::

## Format chart axes

Column, bar, and scatter charts use a category axis and a value axis. Use the axis title to explain what the categories represent, and use the display unit to make large values easier to scan.

### Set an axis title

This example sets the category axis title of the first chart on the worksheet to **Product**.

```js
await Excel.run(async (context) => {
    const sheet = context.workbook.worksheets.getItem("Sample");
    const chart = sheet.charts.getItemAt(0);

    chart.axes.categoryAxis.title.text = "Product";

    await context.sync();
});
```

:::image type="content" source="../images/excel-charts-axis-title-set.png" alt-text="Chart with an axis title in Excel.":::

### Set the axis display unit

This example changes the value axis to use the **Hundreds** display unit.

```js
await Excel.run(async (context) => {
    const sheet = context.workbook.worksheets.getItem("Sample");
    const chart = sheet.charts.getItemAt(0);

    chart.axes.valueAxis.displayUnit = "Hundreds";

    await context.sync();
});
```

:::image type="content" source="../images/excel-charts-axis-display-unit-set.png" alt-text="Chart with the axis display unit set in Excel.":::

## Show or hide gridlines

Gridlines can help users estimate values, but they can also add visual noise. The following example hides the major gridlines on the value axis of the first chart on the worksheet. To show the gridlines again, set `chart.axes.valueAxis.majorGridlines.visible` to `true`.

```js
await Excel.run(async (context) => {
    const sheet = context.workbook.worksheets.getItem("Sample");
    const chart = sheet.charts.getItemAt(0);

    chart.axes.valueAxis.majorGridlines.visible = false;

    await context.sync();
});
```

:::image type="content" source="../images/excel-charts-gridlines-removed.png" alt-text="Chart with gridlines hidden in Excel.":::

## Add and update trendlines

Trendlines help users spot direction and smoothing in a data series. Use a moving average trendline to smooth short-term changes, or switch to a linear trendline when you want to emphasize the overall direction.

### Add a moving average trendline

This example adds a moving average trendline with a 5-period window to the first series in the first chart on the **Sample** worksheet.

```js
await Excel.run(async (context) => {
    const sheet = context.workbook.worksheets.getItem("Sample");
    const chart = sheet.charts.getItemAt(0);
    const seriesCollection = chart.series;

    seriesCollection.getItemAt(0).trendlines.add("MovingAverage").movingAveragePeriod = 5;

    await context.sync();
});
```

:::image type="content" source="../images/excel-charts-create-trendline.png" alt-text="Chart with a moving average trendline in Excel.":::

### Change a trendline to linear

This example changes the first trendline on the first series to type `Linear`.

```js
await Excel.run(async (context) => {
    const sheet = context.workbook.worksheets.getItem("Sample");
    const chart = sheet.charts.getItemAt(0);
    const seriesCollection = chart.series;
    const series = seriesCollection.getItemAt(0);

    series.trendlines.getItem(0).type = "Linear";

    await context.sync();
});
```

:::image type="content" source="../images/excel-charts-trendline-linear.png" alt-text="Chart with a linear trendline in Excel.":::

## Add and format a chart data table

Use a chart data table when users need both the visual chart and the source values in one place. Get the data table by using [`Chart.getDataTableOrNullObject`](/javascript/api/excel/excel.chart#excel-excel-chart-getdatatableornullobject-member(1)). Then use the returned [`ChartDataTable`](/javascript/api/excel/excel.chartdatatable) and [`ChartDataTableFormat`](/javascript/api/excel/excel.chartdatatableformat) objects to control visibility, borders, and font settings.

The following example adds a data table to an existing chart on the **Sample** worksheet and applies simple formatting that matches a typical business chart.

```js
await Excel.run(async (context) => {
    const chart = context.workbook.worksheets.getItem("Sample").charts.getItemAt(0);
    const chartDataTable = chart.getDataTableOrNullObject();

    chartDataTable.visible = true;
    chartDataTable.showLegendKey = true;
    chartDataTable.showHorizontalBorder = false;
    chartDataTable.showVerticalBorder = true;
    chartDataTable.showOutlineBorder = true;

    const chartDataTableFormat = chartDataTable.format;
    chartDataTableFormat.font.color = "#1F1F1F";
    chartDataTableFormat.font.name = "Calibri";
    chartDataTableFormat.border.color = "#4472C4";

    await context.sync();
});
```

:::image type="content" source="../images/excel-charts-data-table.png" alt-text="Chart with a formatted data table in Excel.":::

## Export a chart as an image

Use `Chart.getImage` when your add-in needs to reuse a chart outside Excel, such as in a web page, report, or message body. The method returns a Base64-encoded string that represents the chart as a JPEG image.

```js
await Excel.run(async (context) => {
    const chart = context.workbook.worksheets.getItem("Sheet1").charts.getItem("Chart1");
    const imageAsString = chart.getImage();

    await context.sync();

    console.log(imageAsString.value);
    // Instead of logging the string, you can save it as a file or insert it into HTML.
});
```

`Chart.getImage` accepts three optional parameters: width, height, and fitting mode.

```ts
getImage(width?: number, height?: number, fittingMode?: Excel.ImageFittingMode): OfficeExtension.ClientResult<string>;
```

These parameters control image size while keeping the chart proportionally scaled.

- `Fill`: The image's minimum height or width is the specified height or width, whichever limit is reached first during scaling. This is the default behavior.
- `Fit`: The image's maximum height or width is the specified height or width, whichever limit is reached first during scaling.
- `FitAndCenter`: The image's maximum height or width is the specified height or width, whichever limit is reached first during scaling. The resulting image is centered relative to the other dimension.

## Related articles

- [Core Excel object model concepts for Office Add-ins](excel-add-ins-core-concepts.md)
- [Get Excel worksheet ranges with the JavaScript API](excel-add-ins-ranges-get.md)
- [Work with tables using the Excel JavaScript API](excel-add-ins-tables.md)
- [Work with data labels in charts using the Excel JavaScript API](excel-add-ins-charts-data-labels.md)
