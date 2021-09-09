---
title: Work with Events using the Excel JavaScript API
description: 'A list of events for Excel JavaScript objects. This includes information on using event handlers and the associated patterns.' 
ms.date: 07/02/2021
ms.localizationpriority: medium
---

# Work with Events using the Excel JavaScript API

This article describes important concepts related to working with events in Excel and provides code samples that show how to register event handlers, handle events, and remove event handlers using the Excel JavaScript API.

## Events in Excel

Each time certain types of changes occur in an Excel workbook, an event notification fires. By using the Excel JavaScript API, you can register event handlers that allow your add-in to automatically run a designated function when a specific event occurs. The following events are currently supported.

| Event | Description | Supported objects |
|:---------------|:-------------|:-----------|
| `onActivated` | Occurs when an object is activated. | [**Chart**](/javascript/api/excel/excel.chart#onActivated), [**ChartCollection**](/javascript/api/excel/excel.chartcollection#onActivated), [**Shape**](/javascript/api/excel/excel.shape#onActivated), [**Worksheet**](/javascript/api/excel/excel.worksheet#onActivated), [**WorksheetCollection**](/javascript/api/excel/excel.worksheetcollection#onActivated) |
| `onActivated` | Occurs when a workbook is activated. | [**Workbook**](/javascript/api/excel/excel.workbook#onActivated) |
| `onAdded` | Occurs when an object is added to the collection. | [**ChartCollection**](/javascript/api/excel/excel.chartcollection#onAdded), [**CommentCollection**](/javascript/api/excel/excel.commentcollection#onAdded), [**TableCollection**](/javascript/api/excel/excel.tablecollection#onAdded), [**WorksheetCollection**](/javascript/api/excel/excel.worksheetcollection#onAdded) |
| `onAutoSaveSettingChanged` | Occurs when the `autoSave` setting is changed on the workbook. | [**Workbook**](/javascript/api/excel/excel.workbook#onAutoSaveSettingChanged) |
| `onCalculated` | Occurs when a worksheet has finished calculation (or all the worksheets of the collection have finished). | [**Worksheet**](/javascript/api/excel/excel.worksheet#onCalculated), [**WorksheetCollection**](/javascript/api/excel/excel.worksheetcollection#onCalculated) |
| `onChanged` | Occurs when the data of individual cells or comments has changed. | [**CommentCollection**](/javascript/api/excel/excel.commentcollection#onChanged), [**Table**](/javascript/api/excel/excel.table#onChanged), [**TableCollection**](/javascript/api/excel/excel.tablecollection#onChanged), [**Worksheet**](/javascript/api/excel/excel.worksheet#onChanged), [**WorksheetCollection**](/javascript/api/excel/excel.worksheetcollection#onChanged) |
| `onColumnSorted` | Occurs when one or more columns have been sorted. This happens as the result of a left-to-right sort operation. | [**Worksheet**](/javascript/api/excel/excel.worksheet#onColumnSorted), [**WorksheetCollection**](/javascript/api/excel/excel.worksheetcollection#onColumnSorted) |
| `onDataChanged` | Occurs when data or formatting within the binding is changed. | [**Binding**](/javascript/api/excel/excel.binding#onDataChanged) |
| `onDeactivated` | Occurs when an object is deactivated. | [**Chart**](/javascript/api/excel/excel.chart#onDeactivated), [**ChartCollection**](/javascript/api/excel/excel.chartcollection#onDeactivated), [**Shape**](/javascript/api/excel/excel.shape#onDeactivated), [**Worksheet**](/javascript/api/excel/excel.worksheet#onDeactivated), [**WorksheetCollection**](/javascript/api/excel/excel.worksheetcollection#onDeactivated) |
| `onDeleted` | Occurs when an object is deleted from the collection. | [**ChartCollection**](/javascript/api/excel/excel.chartcollection#onDeleted), [**CommentCollection**](/javascript/api/excel/excel.commentcollection#onDeleted), [**TableCollection**](/javascript/api/excel/excel.tablecollection#onDeleted), [**WorksheetCollection**](/javascript/api/excel/excel.worksheetcollection#onDeleted) |
| `onFormatChanged` | Occurs when the format is changed on a worksheet. | [**Worksheet**](/javascript/api/excel/excel.worksheet#onFormatChanged), [**WorksheetCollection**](/javascript/api/excel/excel.worksheetcollection#onFormatChanged) |
| `onFormulaChanged` | Occurs when a formula is changed. | [**Worksheet**](/javascript/api/excel/excel.worksheet#onFormulaChanged), [**WorksheetCollection**](/javascript/api/excel/excel.worksheetcollection#onFormulaChanged) |
| `onRowSorted` | Occurs when one or more rows have been sorted. This happens as the result of a top-to-bottom sort operation. | [**Worksheet**](/javascript/api/excel/excel.worksheet#onRowSorted), [**WorksheetCollection**](/javascript/api/excel/excel.worksheetcollection#onRowSorted) |
| `onSelectionChanged` | Occurs when the active cell or selected range is changed. | [**Binding**](/javascript/api/excel/excel.binding#onSelectionChanged), [**Table**](/javascript/api/excel/excel.table#onSelectionChanged), [**Workbook**](/javascript/api/excel/excel.workbook#onSelectionChanged), [**Worksheet**](/javascript/api/excel/excel.worksheet#onSelectionChanged), [**WorksheetCollection**](/javascript/api/excel/excel.worksheetcollection#onSelectionChanged) |
| `onRowHiddenChanged` | Occurs when the row-hidden state changes on a specific worksheet. | [**Worksheet**](/javascript/api/excel/excel.worksheet#onRowHiddenChanged), [**WorksheetCollection**](/javascript/api/excel/excel.worksheetcollection#onRowHiddenChanged) |
| `onSettingsChanged` | Occurs when the Settings in the document are changed. | [**SettingCollection**](/javascript/api/excel/excel.settingcollection#onSettingsChanged) |
| `onSingleClicked` | Occurs when left-clicked/tapped action occurs in the worksheet. | [**Worksheet**](/javascript/api/excel/excel.worksheet#onSingleClicked), [**WorksheetCollection**](/javascript/api/excel/excel.worksheetcollection#onSingleClicked) |

### Events in preview

> [!NOTE]
> The following events are currently available only in public preview. [!INCLUDE [Information about using preview APIs](../includes/using-excel-preview-apis.md)]

| Event | Description | Supported objects |
|:---------------|:-------------|:-----------|
| `onFiltered` | Occurs when a filter is applied to an object. | [**Table**](/javascript/api/excel/excel.table#onFiltered), [**TableCollection**](/javascript/api/excel/excel.tablecollection#onFiltered), [**Worksheet**](/javascript/api/excel/excel.worksheet#onFiltered), [**WorksheetCollection**](/javascript/api/excel/excel.worksheetcollection#onFiltered) |

### Event triggers

Events within an Excel workbook can be triggered by:

- User interaction via the Excel user interface (UI) that changes the workbook
- Office Add-in (JavaScript) code that changes the workbook
- VBA add-in (macro) code that changes the workbook

Any change that complies with default behavior of Excel will trigger the corresponding event(s) in a workbook.

### Lifecycle of an event handler

An event handler is created when an add-in registers the event handler. It is destroyed when the add-in unregisters the event handler or when the add-in is refreshed, reloaded, or closed. Event handlers do not persist as part of the Excel file, or across sessions with Excel on the web.

> [!CAUTION]
> When an object to which events are registered is deleted (e.g., a table with an `onChanged` event registered), the event handler no longer triggers but remains in memory until the add-in or Excel session refreshes or closes.

### Events and coauthoring

With [coauthoring](co-authoring-in-excel-add-ins.md), multiple people can work together and edit the same Excel workbook simultaneously. For events that can be triggered by a coauthor, such as `onChanged`, the corresponding **Event** object will contain a **source** property that indicates whether the event was triggered locally by the current user (`event.source = Local`) or was triggered by the remote coauthor (`event.source = Remote`).

## Register an event handler

The following code sample registers an event handler for the `onChanged` event in the worksheet named **Sample**. The code specifies that when data changes in that worksheet, the `handleChange` function should run.

```js
Excel.run(function (context) {
    var worksheet = context.workbook.worksheets.getItem("Sample");
    worksheet.onChanged.add(handleChange);

    return context.sync()
        .then(function () {
            console.log("Event handler successfully registered for onChanged event in the worksheet.");
        });
}).catch(errorHandlerFunction);
```

## Handle an event

As shown in the previous example, when you register an event handler, you indicate the function that should run when the specified event occurs. You can design that function to perform whatever actions your scenario requires. The following code sample shows an event handler function that simply writes information about the event to the console.

```js
function handleChange(event)
{
    return Excel.run(function(context){
        return context.sync()
            .then(function() {
                console.log("Change type of event: " + event.changeType);
                console.log("Address of event: " + event.address);
                console.log("Source of event: " + event.source);
            });
    }).catch(errorHandlerFunction);
}
```

## Remove an event handler

The following code sample registers an event handler for the `onSelectionChanged` event in the worksheet named **Sample** and defines the `handleSelectionChange` function that will run when the event occurs. It also defines the `remove()` function that can subsequently be called to remove that event handler. Note that the `RequestContext` used to create the event handler is needed to remove it. 

```js
var eventResult;

Excel.run(function (context) {
    var worksheet = context.workbook.worksheets.getItem("Sample");
    eventResult = worksheet.onSelectionChanged.add(handleSelectionChange);

    return context.sync()
        .then(function () {
            console.log("Event handler successfully registered for onSelectionChanged event in the worksheet.");
        });
}).catch(errorHandlerFunction);

function handleSelectionChange(event)
{
    return Excel.run(function(context){
        return context.sync()
            .then(function() {
                console.log("Address of current selection: " + event.address);
            });
    }).catch(errorHandlerFunction);
}

function remove() {
    return Excel.run(eventResult.context, function (context) {
        eventResult.remove();

        return context.sync()
            .then(function() {
                eventResult = null;
                console.log("Event handler successfully removed.");
            });
    }).catch(errorHandlerFunction);
}
```

## Enable and disable events

The performance of an add-in may be improved by disabling events.
For example, your app might never need to receive events, or it could ignore events while performing batch-edits of multiple entities.

Events are enabled and disabled at the [runtime](/javascript/api/excel/excel.runtime) level.
The `enableEvents` property determines if events are fired and their handlers are activated.

The following code sample shows how to toggle events on and off.

```js
Excel.run(function (context) {
    context.runtime.load("enableEvents");
    return context.sync()
        .then(function () {
            var eventBoolean = !context.runtime.enableEvents;
            context.runtime.enableEvents = eventBoolean;
            if (eventBoolean) {
                console.log("Events are currently on.");
            } else {
                console.log("Events are currently off.");
            }
        }).then(context.sync);
}).catch(errorHandlerFunction);
```

## See also

- [Excel JavaScript object model in Office Add-ins](excel-add-ins-core-concepts.md)
