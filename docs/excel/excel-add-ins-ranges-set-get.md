---
title: Set and get ranges using the Excel JavaScript API
description: 'Learn how to use the Excel JavaScript API to set and get ranges using.'
ms.date: 03/26/2021
localization_priority: Normal
---

# Set and get ranges using the Excel JavaScript API

[!include[Excel cells and ranges note](../includes/note-excel-cells-and-ranges.md)]

## Set the selected range

The following code sample selects the range **B2:E6** in the active worksheet.

```js
Excel.run(function (context) {
    var sheet = context.workbook.worksheets.getActiveWorksheet();
    var range = sheet.getRange("B2:E6");

    range.select();

    return context.sync();
}).catch(errorHandlerFunction);
```

### Selected range B2:E6

![Selected range in Excel](../images/excel-ranges-set-selection.png)

## Get the selected range

The following code sample gets the selected range, loads its `address` property, and writes a message to the console.

```js
Excel.run(function (context) {
    var range = context.workbook.getSelectedRange();
    range.load("address");

    return context.sync()
        .then(function () {
            console.log(`The address of the selected range is "${range.address}"`);
        });
}).catch(errorHandlerFunction);
```

## See also

- [Work with ranges using the Excel JavaScript API (advanced)](excel-add-ins-ranges-advanced.md)
- [Excel JavaScript object model in Office Add-ins](excel-add-ins-core-concepts.md)
