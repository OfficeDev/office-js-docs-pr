---
title: Work with dates using the Excel JavaScript API 
description: 'Use the Moment-MSDate plug-in with the Excel JavaScript API to work with dates.' 
ms.date: 04/02/2021
ms.prod: excel
localization_priority: Normal
---

# Work with dates using the Excel JavaScript API and the Moment-MSDate plug-in

This article provides code samples that show how to work with dates using the Excel JavaScript API and the [Moment-MSDate plug-in](https://www.npmjs.com/package/moment-msdate). For the complete list of properties and methods that the `Range` object supports, see the [Excel.Range class](/javascript/api/excel/excel.range).

[!include[Excel cells and ranges note](../includes/note-excel-cells-and-ranges.md)]

## Use the Moment-MSDate plug-in to work with dates

The [Moment JavaScript library](https://momentjs.com/) provides a convenient way to use dates and timestamps. The [Moment-MSDate plug-in](https://www.npmjs.com/package/moment-msdate) converts the format of moments into one preferable for Excel. This is the same format the [NOW function](https://support.office.com/article/now-function-3337fd29-145a-4347-b2e6-20c904739c46) returns.

The following code shows how to set the range at **B4** to a moment's timestamp.

```js
Excel.run(function (context) {
    var sheet = context.workbook.worksheets.getItem("Sample");

    var now = Date.now();
    var nowMoment = moment(now);
    var nowMS = nowMoment.toOADate();

    var dateRange = sheet.getRange("B4");
    dateRange.values = [[nowMS]];

    dateRange.numberFormat = [["[$-409]m/d/yy h:mm AM/PM;@"]];

    return context.sync();
}).catch(errorHandlerFunction);
```

The following code sample demonstrates a similar technique to get the date back out of the cell and convert it to a `Moment` or other format.

```js
Excel.run(function (context) {
    var sheet = context.workbook.worksheets.getItem("Sample");

    var dateRange = sheet.getRange("B4");
    dateRange.load("values");

    return context.sync().then(function () {
        var nowMS = dateRange.values[0][0];

        // log the date as a moment
        var nowMoment = moment.fromOADate(nowMS);
        console.log(`get (moment): ${JSON.stringify(nowMoment)}`);

        // log the date as a UNIX-style timestamp
        var now = nowMoment.unix();
        console.log(`get (timestamp): ${now}`);
    });
}).catch(errorHandlerFunction);
```

Your add-in has to format the ranges to display the dates in a more human-readable form. For example, `"[$-409]m/d/yy h:mm AM/PM;@"` displays "12/3/18 3:57 PM". For more information about date and time number formats, see "Guidelines for date and time formats" in the [Review guidelines for customizing a number format](https://support.microsoft.com/office/c0a1d1fa-d3f4-4018-96b7-9c9354dd99f5) article.


## See also

- [Work with cells using the Excel JavaScript API](excel-add-ins-cells.md)
- [Excel JavaScript object model in Office Add-ins](excel-add-ins-core-concepts.md)
- [Work with multiple ranges simultaneously in Excel add-ins](excel-add-ins-multiple-ranges.md)
