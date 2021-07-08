---
title: Work with charts using the Excel JavaScript API
description: 'Code samples demonstrating chart tasks using the Excel JavaScript API.'
ms.date: 07/17/2019
localization_priority: Normal
---

# Work with charts using the Excel JavaScript API

This article provides code samples that show how to perform common tasks with charts using the Excel JavaScript API.
For the complete list of properties and methods that the `Chart` and `ChartCollection` objects support, see [Chart Object (JavaScript API for Excel)](/javascript/api/excel/excel.chart) and [Chart Collection Object (JavaScript API for Excel)](/javascript/api/excel/excel.chartcollection).

## Create a chart

The following code sample creates a chart in the worksheet named **Sample**. The chart is a **Line** chart that is based upon data in the range **A1:B13**.

```js
Excel.run(function (context) {
    var sheet = context.workbook.worksheets.getItem("Sample");
    var dataRange = sheet.getRange("A1:B13");
    var chart = sheet.charts.add("Line", dataRange, "auto");

    chart.title.text = "Sales Data";
    chart.legend.position = "right"
    chart.legend.format.fill.setSolidColor("white");
    chart.dataLabels.format.font.size = 15;
    chart.dataLabels.format.font.color = "black";

    return context.sync();
}).catch(errorHandlerFunction);
```

**New line chart**

![New line chart in Excel.](../images/excel-charts-create-line.png)


## Add a data series to a chart

The following code sample adds a data series to the first chart in the worksheet. The new data series corresponds to the column named **2016** and is based upon data in the range **D2:D5**.

```js
Excel.run(function (context) {
    var sheet = context.workbook.worksheets.getItem("Sample");
    var chart = sheet.charts.getItemAt(0);
    var dataRange = sheet.getRange("D2:D5");

    var newSeries = chart.series.add("2016");
    newSeries.setValues(dataRange);

    return context.sync();
}).catch(errorHandlerFunction);
```

**Chart before the 2016 data series is added**

![Chart in Excel before 2016 data series added.](../images/excel-charts-data-series-before.png)

**Chart after the 2016 data series is added**

![Chart in Excel after 2016 data series added.](../images/excel-charts-data-series-after.png)

## Set chart title

The following code sample sets the title of the first chart in the worksheet to **Sales Data by Year**.

```js
Excel.run(function (context) {
    var sheet = context.workbook.worksheets.getItem("Sample");

    var chart = sheet.charts.getItemAt(0);
    chart.title.text = "Sales Data by Year";

    return context.sync();
}).catch(errorHandlerFunction);
```

**Chart after title is set**

![Chart with title in Excel.](../images/excel-charts-title-set.png)

## Set properties of an axis in a chart

Charts that use the [Cartesian coordinate system](https://en.wikipedia.org/wiki/Cartesian_coordinate_system) such as column charts, bar charts, and scatter charts contain a category axis and a value axis. These examples show how to set the title and display unit of an axis in a chart.

### Set axis title

The following code sample sets the title of the category axis for the first chart in the worksheet to **Product**.

```js
Excel.run(function (context) {
    var sheet = context.workbook.worksheets.getItem("Sample");

    var chart = sheet.charts.getItemAt(0);
    chart.axes.categoryAxis.title.text = "Product";

    return context.sync();
}).catch(errorHandlerFunction);
```

**Chart after title of category axis is set**

![Chart with axis title in Excel.](../images/excel-charts-axis-title-set.png)

### Set axis display unit

The following code sample sets the display unit of the value axis for the first chart in the worksheet to **Hundreds**.

```js
Excel.run(function (context) {
    var sheet = context.workbook.worksheets.getItem("Sample");

    var chart = sheet.charts.getItemAt(0);
    chart.axes.valueAxis.displayUnit = "Hundreds";

    return context.sync();
}).catch(errorHandlerFunction);
```

**Chart after display unit of value axis is set**

![Chart with axis display unit in Excel.](../images/excel-charts-axis-display-unit-set.png)

## Set visibility of gridlines in a chart

The following code sample hides the major gridlines for the value axis of the first chart in the worksheet. You can show the major gridlines for the value axis of the chart, by setting `chart.axes.valueAxis.majorGridlines.visible` to `true`.

```js
Excel.run(function (context) {
    var sheet = context.workbook.worksheets.getItem("Sample");

    var chart = sheet.charts.getItemAt(0);
    chart.axes.valueAxis.majorGridlines.visible = false;

    return context.sync();
}).catch(errorHandlerFunction);
```

**Chart with gridlines hidden**

![Chart with gridlines hidden in Excel.](../images/excel-charts-gridlines-removed.png)

## Chart trendlines

### Add a trendline

The following code sample adds a moving average trendline to the first series in the first chart in the worksheet named **Sample**. The trendline shows a moving average over 5 periods.

```js
Excel.run(function (context) {
    var sheet = context.workbook.worksheets.getItem("Sample");

    var chart = sheet.charts.getItemAt(0);
    var seriesCollection = chart.series;
    seriesCollection.getItemAt(0).trendlines.add("MovingAverage").movingAveragePeriod = 5;

    return context.sync();
}).catch(errorHandlerFunction);
```

**Chart with moving average trendline**

![Chart with moving average trendline in Excel.](../images/excel-charts-create-trendline.png)

### Update a trendline

The following code sample sets the trendline to type `Linear` for the first series in the first chart in the worksheet named **Sample**.

```js
Excel.run(function (context) {
    var sheet = context.workbook.worksheets.getItem("Sample");

    var chart = sheet.charts.getItemAt(0);
    var seriesCollection = chart.series;
    var series = seriesCollection.getItemAt(0);
    series.trendlines.getItem(0).type = "Linear";

    return context.sync();
}).catch(errorHandlerFunction);
```

**Chart with linear trendline**

![Chart with linear trendline in Excel.](../images/excel-charts-trendline-linear.png)

## Export a chart as an image

Charts can be rendered as images outside of Excel. `Chart.getImage` returns the chart as a base64-encoded string representing the chart as a JPEG image. The following code shows how to get the image string and log it to the console.

```js
Excel.run(function (ctx) {
    var chart = ctx.workbook.worksheets.getItem("Sheet1").charts.getItem("Chart1");
    var imageAsString = chart.getImage();
    return context.sync().then(function () {
        console.log(imageAsString.value);
        // Instead of logging, your add-in may use the base64-encoded string to save the image as a file or insert it in HTML.
    });
}).catch(errorHandlerFunction);
```

`Chart.getImage` takes three optional parameters: width, height, and the fitting mode.

```typescript
getImage(width?: number, height?: number, fittingMode?: Excel.ImageFittingMode): OfficeExtension.ClientResult<string>;
```

These parameters determine the size of the image. Images are always proportionally scaled. The width and height parameters put upper or lower bounds on the scaled image. `ImageFittingMode` has three values with the following behaviors.

- `Fill`: The image's minimum height or width is the specified height or width (whichever is reached first when scaling the image). This is the default behavior when no fitting mode is specified.
- `Fit`: The image's maximum height or width is the specified height or width (whichever is reached first when scaling the image).
- `FitAndCenter`: The image's maximum height or width is the specified height or width (whichever is reached first when scaling the image). The resulting image is centered relative to the other dimension.

## See also

- [Excel JavaScript object model in Office Add-ins](excel-add-ins-core-concepts.md)
