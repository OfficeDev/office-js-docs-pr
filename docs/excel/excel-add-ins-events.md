---
title: Work with Events using the Excel JavaScript API
description: ''
ms.date: 10/17/2018
localization_priority: Priority
---

# Work with Events using the Excel JavaScript API 

This article describes important concepts related to working with events in Excel and provides code samples that show how to register event handlers, handle events, and remove event handlers using the Excel JavaScript API. 

## Events in Excel

Each time certain types of changes occur in an Excel workbook, an event notification fires. By using the Excel JavaScript API, you can register event handlers that allow your add-in to automatically run a designated function when a specific event occurs. The following events are currently supported.

| Event | Description | Supported objects |
|:---------------|:-------------|:-----------|
| `onAdded` | Event that occurs when an object is added. | [**ChartCollection**](https://docs.microsoft.com/javascript/api/excel/excel.chartcollection), [**WorksheetCollection**](https://docs.microsoft.com/javascript/api/excel/excel.worksheetcollection) |
| `onDeleted` | Event that occurs when an object is deleted. | [**ChartCollection**](https://docs.microsoft.com/javascript/api/excel/excel.chartcollection), [**WorksheetCollection**](https://docs.microsoft.com/javascript/api/excel/excel.worksheetcollection) |
| `onActivated` | Event that occurs when an object is activated. | [**Chart**](https://docs.microsoft.com/javascript/api/excel/excel.chart), [**ChartCollection**](https://docs.microsoft.com/javascript/api/excel/excel.chartcollection), [**WorksheetCollection**](https://docs.microsoft.com/javascript/api/excel/excel.worksheetcollection), [**Worksheet**](https://docs.microsoft.com/javascript/api/excel/excel.worksheet) |
| `onDeactivated` | Event that occurs when an object is deactivated. | [**Chart**](https://docs.microsoft.com/javascript/api/excel/excel.chart), [**ChartCollection**](https://docs.microsoft.com/javascript/api/excel/excel.chartcollection), [**WorksheetCollection**](https://docs.microsoft.com/javascript/api/excel/excel.worksheetcollection), [**Worksheet**](https://docs.microsoft.com/javascript/api/excel/excel.worksheet) |
| `onCalculated` | Event that occurs when a worksheet has finished calculation (or all the worksheets of the collection have finished). | [**WorksheetCollection**](https://docs.microsoft.com/javascript/api/excel/excel.worksheetcollection), [**Worksheet**](https://docs.microsoft.com/javascript/api/excel/excel.worksheet) |
| `onChanged` | Event that occurs when data within cells is changed. | [**Worksheet**](https://docs.microsoft.com/javascript/api/excel/excel.worksheet), [**Table**](https://docs.microsoft.com/javascript/api/excel/excel.table), [**TableCollection**](https://docs.microsoft.com/javascript/api/excel/excel.tablecollection) |
| `onDataChanged` | Event that occurs when data or formatting within the binding is changed. | [**Binding**](https://docs.microsoft.com/javascript/api/excel/excel.binding) |
| `onSelectionChanged` | Event that occurs when the active cell or selected range is changed. | [**Worksheet**](https://docs.microsoft.com/javascript/api/excel/excel.worksheet), [**Table**](https://docs.microsoft.com/javascript/api/excel/excel.table), [**Binding**](https://docs.microsoft.com/javascript/api/excel/excel.binding) |
| `onSettingsChanged` | Event that occurs when the Settings in the document are changed. | [**SettingCollection**](https://docs.microsoft.com/javascript/api/excel/excel.settingcollection) |

### Event triggers

Events within an Excel workbook can be triggered by:

- User interaction via the Excel user interface (UI) that changes the workbook
- Office Add-in (JavaScript) code that changes the workbook
- VBA add-in (macro) code that changes the workbook

Any change that complies with default behavior of Excel will trigger the corresponding event(s) in a workbook.

### Lifecycle of an event handler

An event handler is created when an add-in registers the event handler. It is destroyed when the add-in unregisters the event handler or when the add-in is refreshed, reloaded, or closed. Event handlers do not persist as part of the Excel file, or across sessions with Excel Online.

> [!CAUTION]
> When an object to which events are registered is deleted (e.g., a table with an `onChanged` event registered), the event handler no longer triggers but remains in memory until the add-in or Excel session refreshes or closes.

### Events and coauthoring

With [coauthoring](co-authoring-in-excel-add-ins.md), multiple people can work together and edit the same Excel workbook simultaneously. For events that can be triggered by a coauthor, such as `onChanged`, the corresponding **Event** object will contain a **source** property that indicates whether the event was triggered locally by the current user (`event.source = Local`) or was triggered by the remote coauthor (`event.source = Remote`).

## Register an event handler

The following code sample registers an event handler for the `onChanged` event in the worksheet named **Sample**. The code specifies that when data changes in that worksheet, the `handleDataChange` function should run.

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

The following code sample registers an event handler for the `onSelectionChanged` event in the worksheet named **Sample** and defines the `handleSelectionChange` function that will run when the event occurs. It also defines the `remove()` function that can subsequently be called to remove that event handler.

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

Events are enabled and disabled at the [runtime](https://docs.microsoft.com/javascript/api/excel/excel.runtime) level. 
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

- [Fundamental programming concepts with the Excel JavaScript API](excel-add-ins-core-concepts.md)
