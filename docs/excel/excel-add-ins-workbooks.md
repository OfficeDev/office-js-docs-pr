---
title: Manage Excel workbooks with the Excel JavaScript API
description: Learn how to get selections, create workbooks, manage protection and settings, control calculation, and save Excel workbooks in an add-in.
ms.date: 06/03/2026
ms.localizationpriority: medium
ms.topic: how-to
ai-usage: ai-assisted
---

# Manage Excel workbooks with the Excel JavaScript API

The `Workbook` object is the top-level object for most Excel operations. Use it when your add-in needs to read the current selection, create or copy workbooks, manage workbook settings, or save changes. This article shows common workbook tasks and the [Application](/javascript/api/excel/excel.application) APIs that support workbook-level behavior.

In an Excel add-in, `Workbook` is the main entry point for workbook data and structure. It gives you access to worksheets, tables, PivotTables, and other workbook content. If you need to work with individual sheets, see [Manage Excel worksheets with the JavaScript API](excel-add-ins-worksheets.md).

## Get the active cell or selected range

### Get the active cell

Use `getActiveCell()` when your add-in needs the user's current focus in the workbook. The method returns a [Range object](/javascript/api/excel/excel.range). The following example loads the cell address and writes it to the console.

```js
await Excel.run(async (context) => {
    const activeCell = context.workbook.getActiveCell();
    activeCell.load("address");
    await context.sync();

    console.log(`The active cell is ${activeCell.address}`);
});
```

### Get the selected range

Use `getSelectedRange()` when your add-in works with the user's current selection. If multiple ranges are selected, the method throws an `InvalidSelection` error. The next example gets the selected range and sets its fill color to yellow.

```js
await Excel.run(async (context) => {
    const range = context.workbook.getSelectedRange();
    range.format.fill.color = "yellow";
    await context.sync();
});
```

## Create or copy a workbook

### Create a new workbook

Use `Excel.createWorkbook()` to open a new workbook in a separate Excel instance. Your add-in stays open in the original workbook.

```js
Excel.createWorkbook();
```

You can also create the new workbook from an existing `.xlsx` file. Pass a Base64-encoded string to `Excel.createWorkbook()`. The following example reads a workbook file, converts it to Base64, and opens the copied workbook.

```js
const fileInput = document.getElementById("file");
const reader = new FileReader();

reader.onload = () => {
    Excel.run(async (context) => {
        const startIndex = reader.result.toString().indexOf("base64,");
        const externalWorkbook = reader.result.toString().substring(startIndex + 7);

        Excel.createWorkbook(externalWorkbook);
        await context.sync();
    });
};

reader.readAsDataURL(fileInput.files[0]);
```

### Copy worksheets from another workbook

If you need to bring existing content into the current workbook, use `insertWorksheetsFromBase64()`. Pass the source workbook as a Base64-encoded string, then specify which worksheets to insert and where to place them.

```typescript
insertWorksheetsFromBase64(base64File: string, options?: Excel.InsertWorksheetOptions): OfficeExtension.ClientResult<string[]>;
```

> [!IMPORTANT]
> `insertWorksheetsFromBase64()` is supported in **Excel on the web**, **Excel on Windows**, and **Excel on Mac**. It isn't supported on **iOS**. In **Excel on the web**, the source worksheets also can't contain PivotTable, Chart, Comment, or Slicer elements. If they do, the method returns an `UnsupportedFeature` error.

The following example reads another workbook, converts it to Base64, and inserts all its worksheets after **Sheet1** in the current workbook. Passing `[]` to `sheetNamesToInsert` inserts every worksheet from the source workbook.

```js
const fileInput = document.getElementById("file");
const reader = new FileReader();

reader.onload = () => {
    Excel.run(async (context) => {
        const startIndex = reader.result.toString().indexOf("base64,");
        const externalWorkbook = reader.result.toString().substring(startIndex + 7);
        const workbook = context.workbook;

        const options = {
            sheetNamesToInsert: [],
            positionType: Excel.WorksheetPositionType.after,
            relativeTo: "Sheet1"
        };

        workbook.insertWorksheetsFromBase64(externalWorkbook, options);
        await context.sync();
    });
};

reader.readAsDataURL(fileInput.files[0]);
```

## Protect workbook structure

Use workbook protection when your add-in needs to prevent users from adding, deleting, moving, or renaming worksheets. The `Workbook.protection` property returns a [WorkbookProtection](/javascript/api/excel/excel.workbookprotection) object.

The following example checks whether workbook structure protection is already enabled. If it isn't, the add-in turns it on.

```js
await Excel.run(async (context) => {
    const workbook = context.workbook;
    workbook.load("protection/protected");
    await context.sync();

    if (!workbook.protection.protected) {
        workbook.protection.protect();
    }
});
```

`protect()` also accepts an optional password string. Use it when you want users to enter a password before they can change workbook structure. For worksheet-level protection, see [Data protection](excel-add-ins-worksheets.md#protect-worksheet-data). For Excel's built-in protection experience, see [Protect a workbook](https://support.microsoft.com/office/7e365a4d-3e89-4616-84ca-1931257c1517).

## Access document properties

A workbook exposes Office file metadata through [document properties](https://support.microsoft.com/office/21d604c2-481e-4379-8e54-1dd4622c6b75). Use `Workbook.properties` to read or update values such as the author.

```js
await Excel.run(async (context) => {
    const properties = context.workbook.properties;
    properties.author = "Alex";
    await context.sync();
});
```

You can also define custom properties. Use the `custom` property on `DocumentProperties` to work with user-defined key-value pairs. For an example, see [Custom properties in Excel and Word](../develop/persisting-add-in-state-and-settings.md#custom-properties-in-excel-and-word).

## Access workbook settings

Workbook settings are similar to custom properties, but they're scoped to a specific workbook and add-in pair. Use settings when your add-in needs to store file-specific state, such as whether the workbook still needs review.

```js
await Excel.run(async (context) => {
    const settings = context.workbook.settings;
    settings.add("NeedsReview", true);

    const needsReview = settings.getItem("NeedsReview");
    needsReview.load("value");

    await context.sync();
    console.log(`Workbook needs review: ${needsReview.value}`);
});
```

## Access application culture settings

Use application culture settings when your add-in displays or parses locale-sensitive data. `Application.cultureInfo` returns a [CultureInfo](/javascript/api/excel/excel.cultureinfo) object with values such as the decimal separator and date format.

Some settings can be changed in the **Excel** UI. When that happens, the system settings remain available through `CultureInfo`, and local overrides are exposed through [Application](/javascript/api/excel/excel.application) properties such as `Application.decimalSeparator`.

The following example converts a number stored with a comma decimal separator to the separator defined by the current system settings.

```js
await Excel.run(async (context) => {
    const sheet = context.workbook.worksheets.getItem("Sample");
    const decimalSource = sheet.getRange("B2");

    decimalSource.load("values");
    context.application.cultureInfo.numberFormat.load("numberDecimalSeparator");
    await context.sync();

    const systemDecimalSeparator =
        context.application.cultureInfo.numberFormat.numberDecimalSeparator;
    const originalValue = decimalSource.values[0][0];
    // This assumes the input column is standardized to use "," as the decimal separator.
    const localizedValue = originalValue.replace(",", systemDecimalSeparator);

    const resultRange = sheet.getRange("C2");
    resultRange.values = [[localizedValue]];
    resultRange.format.autofitColumns();
    await context.sync();
});
```

## Control calculation behavior

When your add-in updates formulas or large data sets, calculation settings can affect both responsiveness and performance. For broader guidance, see [Performance optimization using the Excel JavaScript API](performance.md).

### Set calculation mode

Use `Application.calculationMode` to control when Excel recalculates formulas. The property supports the following values.

- `automatic`: Excel recalculates formulas whenever referenced data changes. This is the default behavior.
- `automaticExceptTables`: Excel recalculates formulas automatically, except for changes to values in tables.
- `manual`: Excel recalculates only when the user or add-in requests it.

### Set calculation type

Use `Application.calculate(calculationType)` to trigger an immediate recalculation. The `calculationType` parameter supports the following values:

- `full`: Recalculate all formulas in all open workbooks, whether or not they changed since the last recalculation.
- `fullRebuild`: Rebuild dependency chains, then recalculate all formulas in all open workbooks.
- `recalculate`: Recalculate formulas that changed, or were marked for recalculation, and the formulas that depend on them in all active workbooks.

For more information about recalculation behavior in Excel, see [Change formula recalculation, iteration, or precision](https://support.microsoft.com/office/73fc7dac-91cf-4d36-86e8-67124f6bcce4).

### Temporarily suspend calculations

Use `suspendApiCalculationUntilNextSync()` when your add-in edits large ranges and doesn't need intermediate formula results before the next `context.sync()` call.

```js
context.application.suspendApiCalculationUntilNextSync();
```

## Detect workbook activation

Use workbook activation events when your add-in needs to refresh data after the user returns to a workbook. A workbook becomes inactive when the user switches to another workbook, another application, or, in **Excel on the web**, another browser tab.

To handle activation, [register an event handler](excel-add-ins-events.md#register-an-event-handler) for the [onActivated](/javascript/api/excel/excel.workbook#excel-excel-workbook-onactivated-member) event. The handler receives a [WorkbookActivatedEventArgs](/javascript/api/excel/excel.workbookactivatedeventargs) object.

> [!IMPORTANT]
> `onActivated` doesn't fire when a workbook is opened. It fires only when the user switches focus back to a workbook that is already open.

The following example registers an activation handler and logs the workbook name when activation occurs.

```js
async function run() {
    await Excel.run(async (context) => {
        const workbook = context.workbook;
        workbook.onActivated.add(workbookActivated);
        await context.sync();
    });
}

async function workbookActivated(event) {
    await Excel.run(async (context) => {
        const workbook = context.workbook;
        workbook.load("name");
        await context.sync();

        console.log(`The workbook ${workbook.name} was activated.`);
    });
}
```

## Save or close a workbook

### Save a workbook

Use `Workbook.save()` to save the workbook to persistent storage. The optional `saveBehavior` parameter supports the following values.

- `Excel.SaveBehavior.save`: Save the workbook without prompting for a name or location. If the file hasn't been saved before, Excel uses the default location.
- `Excel.SaveBehavior.prompt`: Prompt for a name and location only if the file hasn't been saved before.

If the user is prompted to save and then cancels the operation, `save()` throws an exception.

```js
context.workbook.save(Excel.SaveBehavior.prompt);
```

### Close a workbook

Use `Workbook.close()` to close the workbook and any add-ins associated with it. The Excel application stays open. The optional `closeBehavior` parameter supports the following values.

- `Excel.CloseBehavior.save`: Save the workbook before closing it. If the file hasn't been saved before, Excel prompts for a name and location.
- `Excel.CloseBehavior.skipSave`: Close the workbook immediately without saving changes.

```js
context.workbook.close(Excel.CloseBehavior.save);
```

## See also

- [Core Excel object model concepts for Office Add-ins](excel-add-ins-core-concepts.md)
- [Manage Excel worksheets with the JavaScript API](excel-add-ins-worksheets.md)
- [Create, read, and manage tables with the Excel JavaScript API](excel-add-ins-tables.md)
- [Coauthoring in Excel add-ins](co-authoring-in-excel-add-ins.md)
- [Performance optimization using the Excel JavaScript API](performance.md)
