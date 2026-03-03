---
title: Work with dates using the Excel JavaScript API
description: Use the Moment-MSDate plug-in with the Excel JavaScript API to work with dates.
ms.date: 03/03/2026
ms.localizationpriority: medium
---

# Work with dates using the Excel JavaScript API and the Moment-MSDate plug-in

This article provides code samples that show how to work with dates using the Excel JavaScript API and the [Moment-MSDate plug-in](https://www.npmjs.com/package/moment-msdate). For the complete list of properties and methods that the `Range` object supports, see the [Excel.Range class](/javascript/api/excel/excel.range).

[!include[Excel cells and ranges note](../includes/note-excel-cells-and-ranges.md)]

## Key points

- Excel stores dates as sequential serial numbers (OADate format), not JavaScript Date objects.
- Use the Moment-MSDate plug-in to convert between JavaScript dates and Excel's date format.
- Set `numberFormat` to display dates in human-readable formats.
- The Moment.js library provides helpful date manipulation and formatting capabilities.

## Use the Moment-MSDate plug-in to work with dates

Excel stores dates as sequential serial numbers called OADate (OLE Automation Date) format. For example, January 1, 2000 is stored as 36526. This format differs from JavaScript Date objects, making date operations challenging. The [Moment JavaScript library](https://momentjs.com/) provides a way to use dates and timestamps. The [Moment-MSDate plug-in](https://www.npmjs.com/package/moment-msdate) converts between `Moment` objects and Excel's OADate format. This is the same format the [NOW function](https://support.microsoft.com/office/3337fd29-145a-4347-b2e6-20c904739c46) returns.

### Set a date value in a cell

The following code shows how to set the range at **B4** to the current date and time using a `Moment` timestamp.

```js
await Excel.run(async (context) => {
    let sheet = context.workbook.worksheets.getItem("Sample");

    let now = Date.now();
    let nowMoment = moment(now);
    let nowMS = nowMoment.toOADate();

    let dateRange = sheet.getRange("B4");
    dateRange.values = [[nowMS]];

    dateRange.numberFormat = [["[$-409]m/d/yy h:mm AM/PM;@"]];

    await context.sync();
});
```

### Read a date value from a cell

The following code sample demonstrates how to read a date from a cell and convert it to a `Moment` or other format.

```js
await Excel.run(async (context) => {
    let sheet = context.workbook.worksheets.getItem("Sample");

    let dateRange = sheet.getRange("B4");
    dateRange.load("values");

    await context.sync();

    let nowMS = dateRange.values[0][0];

    // Log the date as a moment.
    let nowMoment = moment.fromOADate(nowMS);
    console.log(`get (moment): ${JSON.stringify(nowMoment)}`);

    // Log the date as a UNIX-style timestamp.
    let now = nowMoment.unix();
    console.log(`get (timestamp): ${now}`);
});
```

## Format dates for display

Your add-in needs to format the ranges to display dates in a human-readable form. Excel uses number format codes to control date display. For example, `"[$-409]m/d/yy h:mm AM/PM;@"` displays "12/3/18 3:57 PM". For more information about date and time number formats, see "Guidelines for date and time formats" in the [Review guidelines for customizing a number format](https://support.microsoft.com/office/c0a1d1fa-d3f4-4018-96b7-9c9354dd99f5) article.

The following code sample shows how to set a date with a specific display format.

```js
await Excel.run(async (context) => {
    let sheet = context.workbook.worksheets.getItem("Sample");

    // Set cell A1 to today's date with format "March 15, 2026"
    let dateRange = sheet.getRange("A1");
    let today = moment();
    dateRange.values = [[today.toOADate()]];
    dateRange.numberFormat = [["mmmm d, yyyy"]];

    await context.sync();
});
```

## See also

- [Review guidelines for customizing a number format](https://support.microsoft.com/office/c0a1d1fa-d3f4-4018-96b7-9c9354dd99f5)
- [Excel JavaScript object model in Office Add-ins](excel-add-ins-core-concepts.md)
- [Set and get range values, text, or formulas using the Excel JavaScript API](excel-add-ins-ranges-set-get-values.md)
- [Set range format using the Excel JavaScript API](excel-add-ins-ranges-set-format.md)
