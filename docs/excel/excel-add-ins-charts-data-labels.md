---
title: Work with data labels in charts using the Excel JavaScript API
description: Code samples demonstrating chart data label tasks using the Excel JavaScript API.
ms.date: 04/14/2025
ms.localizationpriority: medium
---

# Work with data labels in charts using the Excel JavaScript API

Add data labels to Excel charts to provide a better visualization experience about important aspects of the chart. To learn more about data labels, see [Add or remove data labels in a chart](https://support.microsoft.com/office/add-or-remove-data-labels-in-a-chart-884bf2f1-2e29-454e-8b42-f467c9f4eb2d).

The following code sample sets up the sample data and **Bicycle Part Production** chart used in this article.

```javascript
async function setup() {
  await Excel.run(async (context) => {
    context.workbook.worksheets.getItemOrNullObject("Sample").delete();
    const sheet = context.workbook.worksheets.add("Sample");

    let salesTable = sheet.tables.add("A1:E1", true);
    salesTable.name = "SalesTable";

    salesTable.getHeaderRowRange().values = [["Product", "Qtr1", "Qtr2", "Qtr3", "Qtr4"]];

    salesTable.rows.add(null, [
      ["Frames", 5000, 7000, 6544, 5377],
      ["Saddles", 400, 323, 276, 1451],
      ["Brake levers", 9000, 8766, 8456, 9812],
      ["Chains", 1550, 1088, 692, 2553],
      ["Mirrors", 225, 600, 923, 344],
      ["Spokes", 6005, 7634, 4589, 8765]
    ]);

    sheet.activate();
    createChart(context);
  });
}

async function createChart(context: Excel.RequestContext) {
  const worksheet = context.workbook.worksheets.getActiveWorksheet();
  const chart = worksheet.charts.add(
    Excel.ChartType.lineMarkers,
    worksheet.getRange("A1:E7"),
    Excel.ChartSeriesBy.rows
  );
  chart.axes.categoryAxis.setCategoryNames(worksheet.getRange("B1:E1"));
  chart.name = "PartChart";

  // Place the chart below sample data.
  chart.top = 125;
  chart.left = 5;
  chart.height = 300;
  chart.width = 450;

  chart.title.text = "Bicycle Part Production";
  chart.legend.position = "Bottom";

  await context.sync();
}
```

This image shows how the chart should display after running the sample code.

:::image type="content" source="../images/excel-data-labels-starter-chart.png" alt-text="Screenshot of basic line chart with no data labels, showing six different bicycle parts being produced over four quarters.":::

## Add data labels

To add data labels to a chart, get the series of data points you want to change, and set the `hasDataLabels` property to `true`.

```javascript
async function addDataLabels() {
  await Excel.run(async (context) => {
    const worksheet = context.workbook.worksheets.getActiveWorksheet();
    const chart = worksheet.charts.getItem("PartChart");
    await context.sync();

    // Get spokes data series.
    const series = chart.series.getItemAt(5);

    // Turn on data labels and set location.
    series.hasDataLabels = true;
    series.dataLabels.position = Excel.ChartDataLabelPosition.top;
    await context.sync();
  });
}
```

:::image type="content" source="../images/excel-data-labels-chart.png" alt-text="Screenshot of chart showing data labels that display the amount for each data point.":::

## Format data label size, shape, and text

You can change attributes on data labels using the following APIs.

- Change data label shapes by setting the [geometricShapeType](/javascript/api/excel/excel.chartdatalabel)  property.
- Change height and width using the [setWidth and setHeight](/javascript/api/excel/excel.chartdatalabel) methods.
- Change the text using the [text](/javascript/api/excel/excel.chartdatalabel) property.
- Change the text formatting using the [format](/javascript/api/excel/excel.chartdatalabel) property. You can change the [border, fill, and font](/javascript/api/excel/excel.chartdatalabelformat) properties.

The following code sample shows how to set the shape type, height and width, and font formatting for the data labels.

```javascript
 await Excel.run(async (context) => {
    const worksheet = context.workbook.worksheets.getActiveWorksheet();
    const chart = worksheet.charts.getItem("PartChart");
    const series = chart.series.getItemAt(5);

    // Set geometric shape of data labels to cubes.
    series.dataLabels.geometricShapeType = Excel.GeometricShapeType.cube;
    series.points.load("count");
    await context.sync();
    let pointsLoaded = series.points.count;

    // Change height, width, and font size of all data labels.
    for (let j = 0; j < pointsLoaded; j++) {
      series.points.getItemAt(j).dataLabel.setWidth(60);
      series.points.getItemAt(j).dataLabel.setHeight(30);
      series.points.getItemAt(j).dataLabel.format.font.size = 12;
    }

    // Set text of a data label.
    series.points.getItemAt(2).dataLabel.setWidth(80);
    series.points.getItemAt(2).dataLabel.setHeight(50);
    series.points.getItemAt(2).dataLabel.text = "Spokes Qtr3: 4589 ↓";

    await context.sync();
});
```

In the following screenshot, the chart now includes *count* data labels for the **Spokes** data, with custom text at the third data point.

:::image type="content" source="../images/excel-data-labels-chart-formats.png" alt-text="Screenshot of chart with data labels set to cubes, new size, and custom text in one of the data labels.":::

You can also change the formatting of text in a data label. The following code sample shows how to use the [getSubstring](/javascript/api/excel/excel.chartdatalabel) method to get part of data label text and apply font formatting.

```javascript
await Excel.run(async (context) => {
    const worksheet = context.workbook.worksheets.getActiveWorksheet();
    const chart = worksheet.charts.getItem("PartChart");
    const series = chart.series.getItemAt(5);

    // Get the "Spokes" data label.
    let label = series.points.getItemAt(2).dataLabel;
    label.load();
    await context.sync();

    // Change border weight of this label.
    label.format.border.weight = 2;
    // Format "Qtr3" as bold and italicized. 
    label.getSubstring(7, 4).font.bold = true;
    label.getSubstring(7, 4).font.italic = true;
    // Format "Spokes" as green.
    label.getSubstring(0, 6).font.color = "green";
    // Format "4589" as red.
    label.getSubstring(12).font.color = "red";
    // Move label up by 15 points.
    label.top = label.top - 15;

    await context.sync();
 });
```

:::image type="content" source="../images/excel-data-labels-chart-substring.png" alt-text="Screenshot of chart showing data label with Spokes set to green, 4589 set to red, and Qtr3 bold and italicized.":::

## Format leader lines

Leader lines connect data labels to their respective data points and make it easier to see what they refer to in the chart. Turn leader lines on using the [showLeaderLines](/javascript/api/excel/excel.chartseries) property. You can set the format of leader lines with the [leaderLines.format](/javascript/api/excel/excel.chartleaderlines) property.

```javascript
await Excel.run(async (context) => {
    const worksheet = context.workbook.worksheets.getActiveWorksheet();
    const chart = worksheet.charts.getItem("PartChart");
    const series = chart.series.getItemAt(5);
    
    // Show leader lines.
    series.showLeaderLines = true;
    await context.sync();
    
    // Format leader lines as dotted orange lines with weight 2.
    series.dataLabels.leaderLines.format.line.lineStyle = Excel.ChartLineStyle.dot;
    series.dataLabels.leaderLines.format.line.color = "orange";
    series.dataLabels.leaderLines.format.line.weight = 2;
});
```

:::image type="content" source="../images/excel-data-labels-chart-leaderlines.png" alt-text="Screenshot of chart with orange dotted leader lines connecting data labels to their data points.":::

## Create callouts

A callout is a data label that connects to the data point using a bubble-shaped pointer. A callout has an anchor which can be moved from the data point to other locations on the chart.

The following code sample shows how to change data labels in a series to use [Excel.GeometricShapeType.wedgeRectCallout](/javascript/api/excel/excel.geometricshapetype). Note that leader lines are turned off to avoid showing two indicators to the same data label.

```javascript
 await Excel.run(async (context) => {
    const worksheet = context.workbook.worksheets.getActiveWorksheet();
    const chart = worksheet.charts.getItem("PartChart");
    const series = chart.series.getItemAt(5);

    // Change to a wedge rectangle style callout.
    series.dataLabels.geometricShapeType = Excel.GeometricShapeType.wedgeRectCallout;
    series.showLeaderLines = false;
    await context.sync();
});
```

:::image type="content" source="../images/excel-data-labels-chart-callout.png" alt-text="Screenshot of chart with data labels formatted as callouts.":::

The following code sample shows how to change the anchor location of a data label.

```javascript
await Excel.run(async (context) => {
    const worksheet = context.workbook.worksheets.getActiveWorksheet();
    const chart = worksheet.charts.getItem("PartChart");
    const series = chart.series.getItemAt(5);

    let label = series.points.getItemAt(2).dataLabel;
    let point = series.points.getItemAt(2);
    label.load();
    await context.sync();

    let anchor = label.getTailAnchor();
    anchor.load();
    await context.sync();

    anchor.top = anchor.top - 10;
    anchor.left = 40;
});
```

This screenshot demonstrates how the anchor of the third data label is adjusted by the preceding code sample.

:::image type="content" source="../images/excel-data-labels-chart-anchor-change.png" alt-text="Screenshot of chart with anchor of Spokes data label moved up and left of the original data point location.":::

## See also

- [Excel JavaScript object model in Office Add-ins](excel-add-ins-core-concepts.md)
- [Work with charts using the Excel JavaScript API](excel-add-ins-charts.md)
