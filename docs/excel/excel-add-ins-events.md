---
title: Work with Events using the Excel JavaScript API
description: A list of events for Excel JavaScript objects. This includes information on using event handlers and the associated patterns.
ms.date: 03/03/2026
ms.localizationpriority: medium
---

# Work with Events using the Excel JavaScript API

This article describes important concepts related to working with events in Excel and provides code samples that show how to register event handlers, handle events, and remove event handlers using the Excel JavaScript API.

## Key points

- Excel fires event notifications when specific changes occur in the workbook.
- Register event handlers using the `add` method on event objects like `worksheet.onChanged`.
- Remove event handlers using the `remove` method with the same `RequestContext`.
- Event handlers don't persist across sessions and are destroyed when the add-in refreshes or closes.
- Disable events temporarily using `context.runtime.enableEvents` to improve performance during batch operations.

## Events in Excel

Each time certain types of changes occur in an Excel workbook, an event notification fires. By using the Excel JavaScript API, you can register event handlers that allow your add-in to automatically run a designated function when a specific event occurs. The following events are currently supported.

| Event | Description | Supported objects |
|:---------------|:-------------|:-----------|
| `onActivated` | Occurs when an object is activated. | [**Chart**](/javascript/api/excel/excel.chart#excel-excel-chart-onactivated-member), [**ChartCollection**](/javascript/api/excel/excel.chartcollection#excel-excel-chartcollection-onactivated-member), [**Shape**](/javascript/api/excel/excel.shape#excel-excel-shape-onactivated-member), [**Worksheet**](/javascript/api/excel/excel.worksheet#excel-excel-worksheet-onactivated-member), [**WorksheetCollection**](/javascript/api/excel/excel.worksheetcollection#excel-excel-worksheetcollection-onactivated-member) |
| `onActivated` | Occurs when a workbook is activated. | [**Workbook**](/javascript/api/excel/excel.workbook#excel-excel-workbook-onactivated-member) |
| `onAdded` | Occurs when an object is added to the collection. | [**ChartCollection**](/javascript/api/excel/excel.chartcollection#excel-excel-chartcollection-onadded-member), [**CommentCollection**](/javascript/api/excel/excel.commentcollection#excel-excel-commentcollection-onadded-member), [**TableCollection**](/javascript/api/excel/excel.tablecollection#excel-excel-tablecollection-onadded-member), [**WorksheetCollection**](/javascript/api/excel/excel.worksheetcollection#excel-excel-worksheetcollection-onadded-member) |
| `onAutoSaveSettingChanged` | Occurs when the `autoSave` setting is changed on the workbook. | [**Workbook**](/javascript/api/excel/excel.workbook#excel-excel-workbook-onautosavesettingchanged-member) |
| `onCalculated` | Occurs when a worksheet has finished calculation (or all the worksheets of the collection have finished). | [**Worksheet**](/javascript/api/excel/excel.worksheet#excel-excel-worksheet-oncalculated-member), [**WorksheetCollection**](/javascript/api/excel/excel.worksheetcollection#excel-excel-worksheetcollection-oncalculated-member) |
| `onChanged` | Occurs when the data of individual cells or comments has changed. | [**CommentCollection**](/javascript/api/excel/excel.commentcollection#excel-excel-commentcollection-onchanged-member), [**Table**](/javascript/api/excel/excel.table#excel-excel-table-onchanged-member), [**TableCollection**](/javascript/api/excel/excel.tablecollection#excel-excel-tablecollection-onchanged-member), [**Worksheet**](/javascript/api/excel/excel.worksheet#excel-excel-worksheet-onchanged-member), [**WorksheetCollection**](/javascript/api/excel/excel.worksheetcollection#excel-excel-worksheetcollection-onchanged-member) |
| `onColumnSorted` | Occurs when one or more columns have been sorted. This happens as the result of a left-to-right sort operation. | [**Worksheet**](/javascript/api/excel/excel.worksheet#excel-excel-worksheet-oncolumnsorted-member), [**WorksheetCollection**](/javascript/api/excel/excel.worksheetcollection#excel-excel-worksheetcollection-oncolumnsorted-member) |
| `onDataChanged` | Occurs when data or formatting within the binding is changed. | [**Binding**](/javascript/api/excel/excel.binding#excel-excel-binding-ondatachanged-member) |
| `onDeactivated` | Occurs when an object is deactivated. | [**Chart**](/javascript/api/excel/excel.chart#excel-excel-chart-ondeactivated-member), [**ChartCollection**](/javascript/api/excel/excel.chartcollection#excel-excel-chartcollection-ondeactivated-member), [**Shape**](/javascript/api/excel/excel.shape#excel-excel-shape-ondeactivated-member), [**Worksheet**](/javascript/api/excel/excel.worksheet#excel-excel-worksheet-ondeactivated-member), [**WorksheetCollection**](/javascript/api/excel/excel.worksheetcollection#excel-excel-worksheetcollection-ondeactivated-member) |
| `onDeleted` | Occurs when an object is deleted from the collection. | [**ChartCollection**](/javascript/api/excel/excel.chartcollection#excel-excel-chartcollection-ondeleted-member), [**CommentCollection**](/javascript/api/excel/excel.commentcollection#excel-excel-commentcollection-ondeleted-member), [**TableCollection**](/javascript/api/excel/excel.tablecollection#excel-excel-tablecollection-ondeleted-member), [**WorksheetCollection**](/javascript/api/excel/excel.worksheetcollection#excel-excel-worksheetcollection-ondeleted-member) |
| `onFormatChanged` | Occurs when the format is changed on a worksheet. | [**Worksheet**](/javascript/api/excel/excel.worksheet#excel-excel-worksheet-onformatchanged-member), [**WorksheetCollection**](/javascript/api/excel/excel.worksheetcollection#excel-excel-worksheetcollection-onformatchanged-member) |
| `onFormulaChanged` | Occurs when a formula is changed. | [**Worksheet**](/javascript/api/excel/excel.worksheet#excel-excel-worksheet-onformulachanged-member), [**WorksheetCollection**](/javascript/api/excel/excel.worksheetcollection#excel-excel-worksheetcollection-onformulachanged-member) |
| `onMoved` | Occurs when a worksheet is moved within a workbook. | [**WorksheetCollection**](/javascript/api/excel/excel.worksheetcollection#excel-excel-worksheetcollection-onmoved-member) |
| `onNameChanged` | Occurs when the worksheet name is changed. | [**Worksheet**](/javascript/api/excel/excel.worksheet#excel-excel-worksheet-onnamechanged-member), [**WorksheetCollection**](/javascript/api/excel/excel.worksheetcollection#excel-excel-worksheetcollection-onnamechanged-member )|
| `onProtectionChanged` | Occurs when the worksheet protection state is changed. | [**Worksheet**](/javascript/api/excel/excel.worksheet#excel-excel-worksheet-onprotectionchanged-member), [**WorksheetCollection**](/javascript/api/excel/excel.worksheetcollection#excel-excel-worksheetcollection-onprotectionchanged-member) |
| `onRowHiddenChanged` | Occurs when the row-hidden state changes on a specific worksheet. This event doesn't fire when rows are hidden or shown by [advanced filters](https://support.microsoft.com/office/4c9222fe-8529-4cd7-a898-3f16abdff32b). See [Known limitations](#known-limitations). | [**Worksheet**](/javascript/api/excel/excel.worksheet#excel-excel-worksheet-onrowhiddenchanged-member), [**WorksheetCollection**](/javascript/api/excel/excel.worksheetcollection#excel-excel-worksheetcollection-onrowhiddenchanged-member) |
| `onRowSorted` | Occurs when one or more rows have been sorted. This happens as the result of a top-to-bottom sort operation. | [**Worksheet**](/javascript/api/excel/excel.worksheet#excel-excel-worksheet-onrowsorted-member), [**WorksheetCollection**](/javascript/api/excel/excel.worksheetcollection#excel-excel-worksheetcollection-onrowsorted-member) |
| `onSelectionChanged` | Occurs when the active cell or selected range is changed. | [**Binding**](/javascript/api/excel/excel.binding#excel-excel-binding-onselectionchanged-member), [**Table**](/javascript/api/excel/excel.table#excel-excel-table-onselectionchanged-member), [**Workbook**](/javascript/api/excel/excel.workbook#excel-excel-workbook-onselectionchanged-member), [**Worksheet**](/javascript/api/excel/excel.worksheet#excel-excel-worksheet-onselectionchanged-member), [**WorksheetCollection**](/javascript/api/excel/excel.worksheetcollection#excel-excel-worksheetcollection-onselectionchanged-member) |
| `onSettingsChanged` | Occurs when the Settings in the document are changed. | [**SettingCollection**](/javascript/api/excel/excel.settingcollection#excel-excel-settingcollection-onsettingschanged-member) |
| `onSingleClicked` | Occurs when left-clicked/tapped action occurs in the worksheet. | [**Worksheet**](/javascript/api/excel/excel.worksheet#excel-excel-worksheet-onsingleclicked-member), [**WorksheetCollection**](/javascript/api/excel/excel.worksheetcollection#excel-excel-worksheetcollection-onsingleclicked-member) |
| `onVisibilityChanged` | Occurs when the worksheet visibility is changed. | [**Worksheet**](/javascript/api/excel/excel.worksheet#excel-excel-worksheet-onvisibilitychanged-member), [**WorksheetCollection**](/javascript/api/excel/excel.worksheetcollection#excel-excel-worksheetcollection-onvisibilitychanged-member) |

### Events in preview

> [!NOTE]
> The following events are currently available only in public preview. [!INCLUDE [Information about using preview APIs](../includes/using-excel-preview-apis.md)]

| Event | Description | Supported objects |
|:---------------|:-------------|:-----------|
| `onFiltered` | Occurs when a filter is applied to an object. | [**Table**](/javascript/api/excel/excel.table#excel-excel-table-onfiltered-member), [**TableCollection**](/javascript/api/excel/excel.tablecollection#excel-excel-tablecollection-onfiltered-member), [**Worksheet**](/javascript/api/excel/excel.worksheet#excel-excel-worksheet-onfiltered-member), [**WorksheetCollection**](/javascript/api/excel/excel.worksheetcollection#excel-excel-worksheetcollection-onfiltered-member) |

### Event triggers

Events within an Excel workbook can be triggered by:

- User interaction via the Excel user interface (UI) that changes the workbook
- Office Add-in (JavaScript) code that changes the workbook
- VBA add-in (macro) code that changes the workbook

Any change that complies with default behavior of Excel will trigger the corresponding events in a workbook.

### Lifecycle of an event handler

An event handler is created when an add-in registers the event handler. It is destroyed when the add-in unregisters the event handler or when the add-in is refreshed, reloaded, or closed. Event handlers do not persist as part of the Excel file, or across sessions with Excel on the web.

> [!CAUTION]
> When an object to which events are registered is deleted (e.g., a table with an `onChanged` event registered), the event handler no longer triggers but remains in memory until the add-in or Excel session refreshes or closes.

### Events and coauthoring

With [coauthoring](co-authoring-in-excel-add-ins.md), multiple people can work together and edit the same Excel workbook simultaneously. For events that can be triggered by a coauthor, such as `onChanged`, the corresponding **Event** object will contain a **source** property that indicates whether the event was triggered locally by the current user (`event.source == Local`) or was triggered by the remote coauthor (`event.source == Remote`).

## Register an event handler

The following code sample registers an event handler for the `onChanged` event in the worksheet named **Sample**. The code specifies that when data changes in that worksheet, the `handleChange` function should run.

```js
await Excel.run(async (context) => {
    const worksheet = context.workbook.worksheets.getItem("Sample");
    worksheet.onChanged.add(handleChange);

    await context.sync();
    console.log("Event handler successfully registered for onChanged event in the worksheet.");
}).catch(errorHandlerFunction);
```

## Handle an event

As shown in the previous example, when you register an event handler, you indicate the function that should run when the specified event occurs. You can design that function to perform whatever actions your scenario requires. The following code sample shows an event handler function that simply writes information about the event to the console.

```js
async function handleChange(event) {
    await Excel.run(async (context) => {
        await context.sync();        
        console.log("Change type of event: " + event.changeType);
        console.log("Address of event: " + event.address);
        console.log("Source of event: " + event.source);       
    }).catch(errorHandlerFunction);
}
```

## Remove an event handler

The following code sample registers an event handler for the `onSelectionChanged` event in the worksheet named **Sample** and defines the `handleSelectionChange` function that will run when the event occurs. It also defines the `remove()` function that can subsequently be called to remove that event handler. Note that the `RequestContext` used to create the event handler is needed to remove it.

```js
let eventResult;

async function run() {
  await Excel.run(async (context) => {
    const worksheet = context.workbook.worksheets.getItem("Sample");
    eventResult = worksheet.onSelectionChanged.add(handleSelectionChange);

    await context.sync();
    console.log("Event handler successfully registered for onSelectionChanged event in the worksheet.");
  });
}

async function handleSelectionChange(event) {
  await Excel.run(async (context) => {
    await context.sync();
    console.log("Address of current selection: " + event.address);
  });
}

async function remove() {
  // The `RequestContext` used to create the event handler is needed to remove it.
  // In this example, `eventContext` is being used to keep track of that context.
  await Excel.run(eventResult.context, async (context) => {
    eventResult.remove();
    await context.sync();
    
    eventResult = null;
    console.log("Event handler successfully removed.");
  });
}
```

## Enable and disable events

The performance of an add-in may be improved by disabling events.
For example, your app might never need to receive events, or it could ignore events while performing batch-edits of multiple entities.

Events are enabled and disabled at the [runtime](/javascript/api/excel/excel.runtime) level.
The `enableEvents` property determines if events are fired and their handlers are activated.

The following code sample shows how to toggle events on and off.

```js
await Excel.run(async (context) => {
    context.runtime.load("enableEvents");
    await context.sync();

    let eventBoolean = !context.runtime.enableEvents;
    context.runtime.enableEvents = eventBoolean;
    if (eventBoolean) {
        console.log("Events are currently on.");
    } else {
        console.log("Events are currently off.");
    }
    
    await context.sync();
});
```

## Known limitations

### `onRowHiddenChanged` doesn't fire for advanced filters

The `onRowHiddenChanged` event doesn't fire when rows are hidden or shown as a result of applying an [advanced filter](https://support.microsoft.com/office/4c9222fe-8529-4cd7-a898-3f16abdff32b). It fires correctly for standard worksheet filters and manual row hide or show operations.

As a workaround, poll the `rowHidden` property on a set of rows at a regular interval to detect visibility changes caused by advanced filters. The following code sample demonstrates this approach.

```js
// Global variables to manage the polling mechanism.
let pollTimer = null;
const rowHiddenSnapshot = new Map(); // Stores previous row visibility states.

async function startTracking() {
    // Prevent multiple timers from running simultaneously.
    if (pollTimer) {
        console.log("Already tracking.");
        return;
    }

    console.log("Starting row visibility tracking...");
    // Poll every 500ms - adjust this interval based on your performance needs.
    pollTimer = window.setInterval(() => {
        pollRowHiddenDiff();
    }, 500);
}

function stopTracking() {
    // Clean up the polling mechanism.
    if (pollTimer) {
        clearInterval(pollTimer);
        pollTimer = null;
        rowHiddenSnapshot.clear(); // Clear the snapshot to free memory.
        console.log("Stopped tracking.");
    }
}

async function pollRowHiddenDiff() {
    try {
        await Excel.run(async (context) => {
            // Get the active worksheet and its used range.
            const sheet = context.workbook.worksheets.getActiveWorksheet();
            const usedRange = sheet.getUsedRange();
            usedRange.load("rowIndex,rowCount"); // Load properties we need.
            await context.sync();

            // Get the starting row (0-based) and limit the number of rows to check.
            const startRow = usedRange.rowIndex;
            const maxRows = Math.min(usedRange.rowCount, 200); // Limit to 200 rows for performance.

            // Build an array of row ranges to check.
            const rows = [];
            for (let i = 0; i < maxRows; i++) {
                // Convert to 1-based row number for Excel range addressing.
                const row1Based = startRow + i + 1;
                // Create a range for the entire row (e.g., "5:5" for row 5).
                const row = sheet.getRange(`${row1Based}:${row1Based}`);
                row.load("rowHidden"); // Load the rowHidden property.
                rows.push(row);
            }

            // Sync to get all the rowHidden values.
            await context.sync();

            // Arrays to track which rows changed visibility.
            const hidden = [];
            const unhidden = [];

            // Compare current visibility state with previous snapshot.
            for (let i = 0; i < rows.length; i++) {
                const rowIndex = startRow + i; // 0-based row index for internal tracking.
                const current = rows[i].rowHidden; // Current visibility state.
                const previous = rowHiddenSnapshot.get(rowIndex); // Previous state from snapshot.

                // If this is the first time we're checking this row, just store its state.
                if (previous === undefined) {
                    rowHiddenSnapshot.set(rowIndex, current);
                    continue;
                }

                // If the visibility state changed, record it and update the snapshot.
                if (previous !== current) {
                    rowHiddenSnapshot.set(rowIndex, current);
                    // Convert to 1-based row numbers for user-friendly logging.
                    if (current) {
                        hidden.push(rowIndex + 1); // Row became hidden.
                    } else {
                        unhidden.push(rowIndex + 1); // Row became visible.
                    }
                }
            }

            // Log any detected changes.
            if (hidden.length || unhidden.length) {
                console.log("Row visibility changed:", { hidden, unhidden });
            }
        });
    } catch (error) {
        // Handle errors gracefully to prevent polling from stopping.
        console.error("Error during row visibility polling:", error);
        
        // If the error is severe (e.g., worksheet deleted), consider stopping tracking.
        if (error.code === "ItemNotFound") {
            console.log("Worksheet no longer exists, stopping tracking.");
            stopTracking();
        }
    }
}
```

## See also

- [Excel JavaScript object model in Office Add-ins](excel-add-ins-core-concepts.md)
- [Work with worksheets using the Excel JavaScript API](excel-add-ins-worksheets.md)
- [Work with tables using the Excel JavaScript API](excel-add-ins-tables.md)
- [Coauthoring in Excel add-ins](co-authoring-in-excel-add-ins.md)
