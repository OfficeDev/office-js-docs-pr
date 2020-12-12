---
title: Troubleshooting Excel Add-ins
description: 'Learn how to troubleshoot development errors in Excel Add-ins.'
ms.date: 12/11/2020
localization_priority: Normal
---

# Troubleshooting Excel Add-ins

This article discusses troubleshooting issues that are unique to Excel. Please use the feedback tool at the bottom of the page to suggest other issues that can be added to the article.

## API limitations when the active workbook switches

Add-ins for Excel are intended to operate on a single workbook at a time. Errors can arise when a workbook that is separate from the one running the add-in gains focus. This only happens when particular methods are in the process of being called when the focus changes.

The following APIs are affected by this workbook switch:

|Excel JavaScript API | Error thrown |
|--|--|
| `Chart.activate` | GeneralException |
| `Range.select` | GeneralException |
| `Table.clearFilters` | GeneralException |
| `Workbook.getActiveCell`  | InvalidSelection|
| `Workbook.getSelectedRange` | InvalidSelection|
| `Workbook.getSelectedRanges`  | InvalidSelection|
| `Worksheet.activate` | GeneralException |
| `Worksheet.delete`  | InvalidSelection|
| `Worksheet.gridlines` | GeneralException |
| `Worksheet.showHeadings` | GeneralException |
| `WorksheetCollection.add` | GeneralException |
| `WorksheetFreezePanes.freezeAt` | GeneralException |
| `WorksheetFreezePanes.freezeColumns` | GeneralException |
| `WorksheetFreezePanes.freezeRows` | GeneralException |
| `WorksheetFreezePanes.getLocationOrNullObject`| GeneralException |
| `WorksheetFreezePanes.unfreeze` | GeneralException |

> [!NOTE]
> This only applies to multiple Excel workbooks open on Windows or Mac.

## Coauthoring

See [Coauthoring in Excel add-ins](co-authoring-in-excel-add-ins.md) for patterns to use with events in a coauthoring environment. The article also discusses potential merge conflicts when using certain APIs, such as [`TableRowCollection.add`](/javascript/api/excel/excel.tablerowcollection#add-index--values-).

## Known Issues

### Binding events ID discrepancy

Both [BindingDataChangedEventArgs.binding](/javascript/api/excel/excel.bindingdatachangedeventargs#binding) and [BindingSelectionChangedEventArgs.binding](/javascript/api/excel/excel.bindingselectionchangedeventargs#binding) return a temporary `Binding` object that contains the ID of the `Binding` object that raised the event. Use this ID with `BindingCollection.getItem(id)` to retrieve the `Binding` object that raised the event.

The following code sample first shows how to assign an event listener to a binding, and then call the `bindingCallback` method when the `onDataChanged` event is triggered. In the `bindingCallback` method, the code sample then demonstrates how to use the ID of the temporary `Binding` object to retrieve the `Binding` object that raised the event.

```js
Excel.run(function (context) {
    // Retrieve your binding.
    var binding = context.workbook.bindings.getItemAt(0);

    return context.sync().then(function () {
        // Register an event listener to detect changes to your binding
        // and then trigger the bindingCallback method. 
        binding.onDataChanged.add(bindingCallback);

        return context.sync();
    });
});

function bindingCallback(eventArgs) {
    return Excel.run(function (context) {
        // Get the temporary binding object and load its ID. 
        var tempBindingObject = eventArgs.binding;
        tempBindingObject.load("id");

        // Use the temporary binding object's ID to retrieve the original binding object. 
        var originalBindingObject = context.workbook.bindings.getItem(tempBindingObject.id);

        // Get the address of the original binding object. 
        var bindingObject = originalBindingObject.getRange();
        bindingObject.load("address");

        return context.sync().then(function () {
            // Print out the address of the binding object. 
            console.log(bindingObject.address);
        });
    });
}
```

### Cell format `useStandardHeight` issue

The [useStandardHeight](/javascript/api/excel/excel.cellpropertiesformat#useStandardHeight) property of `CellPropertiesFormat` doesn't work properly in Excel on the web. Due to a bug in the Excel on the web UI, setting the `useStandardHeight` property to `true` will calculate height imprecisely on this platform. For example, a standard height of **14** will be modified to **14.25** in Excel on the web.

### Range `getImage` method

The Range [getImage](/javascript/api/excel/excel.range#getImage__) method isn't currently supported in Excel for Mac. See [OfficeDev/office-js Issue #235](https://github.com/OfficeDev/office-js/issues/235) for the current status.

### Range return character limit

The [Worksheet.getRange(address)](/javascript/api/excel/excel.worksheet#getRange_address_) and [Worksheet.getRanges(address)](/javascript/api/excel/excel.worksheet#getRanges_address_) methods have an address string limit of 8192 characters. When this limit is exceeded, the address string is truncated to 8192 characters.

## See also

- [Troubleshoot development errors with Office Add-ins](../testing/troubleshoot-development-errors.md)
- [Troubleshoot user errors with Office Add-ins](../testing/testing-and-troubleshooting.md)
