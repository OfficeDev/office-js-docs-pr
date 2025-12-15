---
title: Coauthoring in Excel add-ins
description: Learn how to handle coauthoring scenarios in Excel add-ins to prevent stale data, synchronization conflicts, and merge issues.
ms.date: 12/15/2025
ms.localizationpriority: medium
---


# Coauthoring in Excel add-ins  

When multiple users and add-ins work in the same Excel workbook, changes made by one user can create unexpected behavior in another user's add-in instance. Cached values for your add-in can become stale when coauthors modify the workbook. This stale data results in your add-in displaying incorrect data, making decisions based on outdated information, or creating merge conflicts.

## Why add-in developers need to handle coauthoring

Excel supports [coauthoring](https://support.microsoft.com/office/7152aa8b-b791-414c-a3bb-3024e46fb104) in workbooks stored on OneDrive, OneDrive for Business, or SharePoint Online. When AutoSave is enabled, changes synchronize in real-time. **Your add-in code doesn't automatically know when coauthors modify the workbook**.

You need to handle coauthoring if your add-in:

- Caches workbook values in JavaScript variables (risk of stale data).
- Stores state in hidden worksheets (risk of lost synchronization).
- Adds rows to tables using `TableRowCollection.add` (risk of merge conflicts).
- Shows UI in response to data changes (risk of unexpected dialogs for all users).

If your add-in only reads data once at startup or rarely runs in shared workbooks, coauthoring support is lower priority but should still be monitored.

> [!IMPORTANT]
> In Excel for Microsoft 365, AutoSave synchronizes changes in real-time. When you turn on AutoSave, coauthoring problems become more frequent and noticeable. Test your add-in with AutoSave enabled to identify potential problems. Users can toggle AutoSave via the switch in the upper left of the Excel window.

## Excel synchronizes workbook content, not your add-in's memory

Excel automatically synchronizes workbook content (such as cell values, formatting, table data) across all coauthors. However, **Excel doesn't synchronize your add-in's JavaScript variables, objects, or in-memory state**. Each user runs their own instance of your add-in with separate memory.

### Problem: Stale data from cached variables

In the following code, User A's `cachedValue` variable never updates automatically. If your add-in logic uses `cachedValue` for calculations, displays, or decisions, it's working with outdated information.

```js
// User A's add-in reads a value.
const range = context.workbook.worksheets.getActiveWorksheet().getRange("A1");
range.load("values");
await context.sync();

const cachedValue = range.values[0][0]; // Stores "Contoso".
console.log(cachedValue); // "Contoso"

// Meanwhile, User B (coauthor) changes A1 to "Fabrikam".
// User B's change synchronizes to the workbook.

// User A's add-in still has the old value.
console.log(cachedValue); // Still "Contoso" - STALE!

// The workbook has the new value.
range.load("values");
await context.sync();
console.log(range.values[0][0]); // "Fabrikam" - CURRENT
```

Each coauthor has their own separate add-in instance. When you copy workbook values to JavaScript variables, those copies don't stay synchronized with the workbook. You must explicitly refresh values or use events to detect changes.

### Solution: Use events to detect coauthor changes

To keep your add-in's state synchronized when coauthors modify the workbook, use Excel events. Events notify your add-in when workbook content changes, so you can refresh cached data or update your UI.

| Scenario | Event to use | Reason |
|----------|-----------|--------|
| Hidden worksheet stores settings | `BindingDataChanged` | Detect when coauthors change configuration |
| Dashboard displays cell values | `BindingDataChanged` | Keep display synchronized with workbook |
| Monitor specific range for changes | `WorksheetChanged` | More flexible for complex change detection |
| Track any worksheet modification | `WorksheetChanged` | Broader change awareness |

#### Example: Keeping a dashboard synchronized

Scenario: Your add-in displays a dashboard showing data from cells A1:C10. Without event handling, the dashboard shows stale data when coauthors update those cells.

The following code uses the `BindingDataChanged` event ([BindingDataChanged](/javascript/api/office/office.bindingdatachangedeventargs)). This event activates whenever any user (local or coauthor) modifies the bound range. The event handler refreshes the cached data, so all users see current information.

```js
let cachedData = null;

// Initial load.
async function loadDashboard() {
  await Excel.run(async (context) => {
    const range = context.workbook.worksheets.getActiveWorksheet().getRange("A1:C10");
    range.load("values");
    await context.sync();
    
    cachedData = range.values;
    updateDashboardDisplay(cachedData);
  });
}

// Set up event to detect changes from coauthors.
async function setupCoauthoringSupport() {
  await Excel.run(async (context) => {
    const sheet = context.workbook.worksheets.getActiveWorksheet();
    const range = sheet.getRange("A1:C10");
    
    // Create a binding to enable change detection.
    const binding = context.workbook.bindings.add(range, Excel.BindingType.range, "DashboardData");
    await context.sync();
    
    // Register event handler for data changes.
    binding.onDataChanged.add(handleDataChange);
    await context.sync();
  });
}

// This activates when coauthors change the bound range.
async function handleDataChange(event) {
  await Excel.run(async (context) => {
    const binding = context.workbook.bindings.getItem("DashboardData");
    const range = binding.getRange();
    range.load("values");
    await context.sync();
    
    // Update cached data and refresh display.
    cachedData = range.values;
    updateDashboardDisplay(cachedData);
  });
}

function updateDashboardDisplay(data) {
  // Update your UI with the current data.
  console.log("Dashboard refreshed with current data");
}
```

### Don't show UI in event handlers

When coauthoring is active, your event handlers run for **all users** when **any user** makes a change. This behavior creates a critical design constraint.

**❌ Don't do this**:

```js
binding.onDataChanged.add(async (event) => {
  // This is a bad pattern. It shows a dialog to all users when any user changes data.
  Office.context.ui.displayDialogAsync("https://contoso.com/validation.html");
});
```

When User B changes a cell, User A unexpectedly sees a validation dialog, even though User A didn't make any changes. This experience is confusing and disruptive.

**✅ Do this instead**:

```js
let cachedData = null;

binding.onDataChanged.add(async (event) => {
  await Excel.run(async (context) => {
    const range = event.binding.getRange();
    range.load("values");
    await context.sync();
    
    // Update internal state silently.
    cachedData = range.values;
    
    // Update displayed values without dialogs or alerts.
    document.getElementById("dashboard").textContent = JSON.stringify(cachedData);
  });
});

// Only show UI in response to explicit user actions.
document.getElementById("showData").onclick = async () => {
  // Now it's OK to show UI - user clicked a button.
  Office.context.ui.displayDialogAsync("https://contoso.com/validation.html");
};
```

Use events to update your add-in's internal state and passive displays. Only show dialogs, alerts, or modal UI in response to explicit user actions, such as button clicks or menu selections.

## Avoid table row conflicts in coauthoring scenarios

When your add-in uses [`TableRowCollection.add`](/javascript/api/excel/excel.tablerowcollection#excel-excel-tablerowcollection-add-member(1)) while coauthors are editing the same table or nearby cells, Excel detects a merge conflict. Users see a yellow bar prompting them to refresh, and recent changes might be lost.

The `TableRowCollection.add` API changes the table structure in a way that conflicts with simultaneous edits. When User A's add-in adds a row while User B is editing cell B5, Excel can't safely merge both changes.

### Use Range.values to add rows

Instead of using the Table API, set values in the range directly below the table. Excel automatically expands the table without creating conflicts.

**❌ Don't do this (it causes conflicts)**:

```js
const table = context.workbook.tables.getItem("SalesData");
table.rows.add(null, [["Product", 100, "=B2*1.2"]]);
// This is a bad pattern. This code causes coauthoring conflicts.
```

**✅ Use this approach**:

```js
await Excel.run(async (context) => {
  const table = context.workbook.tables.getItem("SalesData");
  const tableRange = table.getRange();
  tableRange.load("rowCount, address");
  await context.sync();
  
  // Get the range directly below the table.
  const sheet = context.workbook.worksheets.getActiveWorksheet();
  const newRowRange = table.getDataBodyRange().getRowsBelow(1);
  
  // Set values - table automatically expands without conflicts.
  newRowRange.values = [["Product", 100, "=B2*1.2"]];
  await context.sync();
});
```

### Additional requirements

For the `Range.values` approach to work reliably:

1. **No data validation rules below the table**: Remove [data validation rules](https://support.microsoft.com/office/29fecbcc-d1b9-42c1-9d76-eff3ce5f7249) from cells below the table, or apply validation to entire columns instead of specific cell ranges.

1. **Handle existing data below the table**: If users have data below your table, insert a blank row first.

   ```js
   // Insert empty row to push existing data down.
   let insertRange = table.getDataBodyRange().getRowsBelow(1);
   insertRange.insert(Excel.InsertShiftDirection.down);
   await context.sync();
   
   // Now set your data.
   insertRange = table.getDataBodyRange().getRowsBelow(1);
   insertRange.values = [["Product", 100, "=B2*1.2"]];
   ```

1. **Can't add truly empty rows**: Tables only auto-expand when you set actual data. If you need an empty row, use a workaround.
   - Put temporary data (like a space character) in a hidden column.
   - Use placeholder data that users can clear.

## Troubleshooting common coauthoring issues

| Symptom | Likely cause | Fix |
|---------|--------------|-----|
| Add-in displays outdated values | Values cached in JavaScript variables | Implement event handlers to refresh on changes |
| Yellow "refresh" bar appears frequently | Using `TableRowCollection.add` | Switch to `Range.values` for adding rows |
| Dialogs pop up unexpectedly | Showing UI in event handlers | Only show UI from user-initiated actions |
| Settings don't sync between users | Hidden worksheet not monitored for changes | Add `BindingDataChanged` event on settings range |
| Changes lost during coauthoring | Merge conflict from table modifications | Follow table row best practices |

## See also

- [About coauthoring in Excel (VBA)](/office/vba/excel/concepts/about-coauthoring-in-excel)
- [How AutoSave impacts add-ins and macros (VBA)](/office/vba/library-reference/concepts/how-autosave-impacts-addins-and-macros)
- [Working with Events using the Excel JavaScript API](../excel/excel-add-ins-events.md)
- [Excel JavaScript API performance optimization](../excel/performance.md)
- [Bindings in Excel add-ins](../excel/excel-add-ins-ranges-advanced.md#work-with-bindings-using-the-excel-javascript-api)
