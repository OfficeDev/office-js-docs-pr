---
title: Clear or delete ranges using the Excel JavaScript API
description: 'Learn how to clear or delete ranges using the Excel JavaScript API.'
ms.date: 04/02/2021
ms.prod: excel
localization_priority: Normal
---

# Clear or delete ranges using the Excel JavaScript API

This article provides code samples that clear and delete ranges with the Excel JavaScript API. For the complete list of properties and methods supported by the `Range` object, see [Excel.Range class](/javascript/api/excel/excel.range).

[!include[Excel cells and ranges note](../includes/note-excel-cells-and-ranges.md)]

## Clear a range of cells

The following code sample clears all contents and formatting of cells in the range **E2:E5**.  

```js
Excel.run(function (context) {
    var sheet = context.workbook.worksheets.getItem("Sample");
    var range = sheet.getRange("E2:E5");

    range.clear();

    return context.sync();
}).catch(errorHandlerFunction);
```

### Data before range is cleared

![Data in Excel before range is cleared.](../images/excel-ranges-start.png)

### Data after range is cleared

![Data in Excel after range is cleared.](../images/excel-ranges-after-clear.png)

## Delete a range of cells

The following code sample deletes the cells in the range **B4:E4** and shifts other cells up to fill the space that was vacated by the deleted cells.

```js
Excel.run(function (context) {
    var sheet = context.workbook.worksheets.getItem("Sample");
    var range = sheet.getRange("B4:E4");

    range.delete(Excel.DeleteShiftDirection.up);

    return context.sync();
}).catch(errorHandlerFunction);
```

### Data before range is deleted

![Data in Excel before range is deleted.](../images/excel-ranges-start.png)

### Data after range is deleted

![Data in Excel after range is deleted.](../images/excel-ranges-after-delete.png)


## See also

- [Work with cells using the Excel JavaScript API](excel-add-ins-cells.md)
- [Set and get ranges using the Excel JavaScript API](excel-add-ins-ranges-set-get.md)
- [Excel JavaScript object model in Office Add-ins](excel-add-ins-core-concepts.md)
