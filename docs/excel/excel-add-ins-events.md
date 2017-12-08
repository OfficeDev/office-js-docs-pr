---
title: Work with Events using the Excel JavaScript API
description: ''
ms.date: 12/08/2017
---

# Work with Events using the Excel JavaScript API

This article describes important concepts about events in Excel and provides code samples that show how to register event handlers, handle events, and remove event handlers using the Excel JavaScript API. 

> [!IMPORTANT]
> The APIs described in this article are currently available only in public preview (beta) and are not intended for use in production environments. To run the code samples that this article contains, you must use a sufficiently recent build of Office and reference the beta library of the Office.js CDN: https://appsforoffice.microsoft.com/lib/beta/hosted/office.js.

## Events in Excel

Each time certain types of changes occur in an Excel workbook, an event notification fires. By using the Excel JavaScript API, you can register event handlers that allow your add-in to automatically run a designated function when a specific event occurs. The following events are currently supported.

| Event | Description |
|:---------------|:-------------|
| `onAdded` | Event that occurs when an object is added. Currently supported by the following object(s): **Worksheet**. |
| `onDeleted`  | Event that occurs when an object is deleted. Currently supported by the following object(s): **Worksheet**. |
| `onActivated` | Event that occurs when an object is activated. Currently supported by the following object(s): **Worksheet**. |
| `onDeactivated` | Event that occurs when an object is deactivated.  Currently supported by the following object(s): **Worksheet**. |
| `onDataChanged` | Event that occurs when data within cells is changed. Currently supported by the following object(s): **Worksheet**, **Table**, **TableCollection**, **Binding**. |
| `onSelectionChanged` | Event that occurs when the active cell or selected range is changed. Currently supported by the following object(s): **Worksheet**, **Table**, **Binding**. |

### Event triggers

Events within an Excel workbook can be triggered by:

- User interaction via the Excel user interface (UI) that changes the workbook
- Office add-in (JavaScript) code that changes the workbook
- VBA add-in (macro) code that changes the workbook

Any change that complies with default behavior of Excel will trigger the corresponding event(s) in a workbook.

### Lifecycle of an event handler

An event handler is created when an add-in registers the event handler and is destroyed when the add-in unregisters the event handler or when the add-in is closed. Event handlers do not persist as part of the Excel file.

## Register an event handler

The following code sample registers an event handler for the `onDataChanged` event in the worksheet named **Sample**. The code specifies that when data changes in that worksheet, the `handleDataChange` function should run.

```js
Excel.run(function (context) {
    var worksheet = context.workbook.worksheets.getItem("Sample");
    worksheet.onDataChanged.add(handleDataChange);

    return context.sync()
        .then(function () {
            console.log("Event handler successfully registered for onDataChanged event in the worksheet.");
        });
}).catch(errorHandlerFunction);
```

## Handle an event

As shown in the previous example, when you register an event handler, you indicate the function that should run when the specified event occurs. You can design that function to perform whatever actions your scenario requires. The following code sample shows an event handler function that simply writes information about the event to the console. 

```js
function handleDataChange(event)
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

The following code sample registers an event handler for the `onSelectionChanged` event in the worksheet named **Sample** and specifies the `remove()` function which can subsequently be called to remove that event handler.

```js
Excel.run(function (context) {
    var worksheet = context.workbook.worksheets.getItem("Sample");
    eventResult = worksheet.onSelectionChanged.add(handleSelectionChange);

    return context.sync()
        .then(function () {
            console.log("Event handler successfully registered for onSelectionChanged event in the worksheet.");
        });
}).catch(errorHandlerFunction);

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

## See also

- [Excel JavaScript API core concepts](excel-add-ins-core-concepts.md)
- [Excel JavaScript API Open Specification](https://github.com/OfficeDev/office-js-docs/tree/ExcelJs_OpenSpec)
- [Introduction to Excel Event Features (preview)](https://github.com/OfficeDev/office-js-docs/blob/ExcelJs_OpenSpec/Event_README.md)