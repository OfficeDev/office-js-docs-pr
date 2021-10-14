---
title: Excel JavaScript API online-only requirement set
description: 'Details about the ExcelApiOnline requirement set.'
ms.date: 10/13/2021
ms.prod: excel
ms.localizationpriority: medium
---

# Excel JavaScript API online-only requirement set

The `ExcelApiOnline` requirement set is a special requirement set that includes features that are only available for Excel on the web. APIs in this requirement set are considered to be production APIs (not subject to undocumented behavioral or structural changes) for the Excel on the web application. `ExcelApiOnline` APIs are considered to be "preview" APIs for other platforms (Windows, Mac, iOS) and may not be supported by any of those platforms.

When APIs in the `ExcelApiOnline` requirement set are supported across all platforms, they will added to the next released requirement set (`ExcelApi 1.[NEXT]`). Once that new requirement is public, those APIs will be removed from `ExcelApiOnline`. Think of this as a similar promotion process to an API moving from preview to release.

> [!IMPORTANT]
> `ExcelApiOnline` is a superset of the latest numbered requirement set.

> [!IMPORTANT]
> `ExcelApiOnline 1.1` is the only version of the online-only APIs. This is because Excel on the web will always have a single version available to users that is the latest version.

The following table provides a concise summary of the APIs, while the subsequent [API list](#api-list) table gives a detailed list of the current `ExcelApiOnline` APIs.

| Feature area | Description | Relevant objects |
|:--- |:--- |:--- |
| Linked workbooks | Manage links between workbooks, including support for refreshing and breaking workbook links. | [LinkedWorkbook](/javascript/api/excel/excel.linkedworkbook), [LinkedWorkbookCollection](/javascript/api/excel/excel.linkedworkbookcollection) |
| Named sheet views | Gives programmatic control of per-user worksheet views. | [NamedSheetView](/javascript/api/excel/excel.namedsheetview), [NamedSheetViewCollection](/javascript/api/excel/excel.namedsheetviewcollection) |

## Recommended usage

Because `ExcelApiOnline` APIs are only supported by Excel on the web, your add-in should check if the requirement set is supported before calling these APIs. This avoids calling an online-only API on a different platform.

```js
if (Office.context.requirements.isSetSupported("ExcelApiOnline", "1.1")) {
   // Any API exclusive to the ExcelApiOnline requirement set.
}
```

Once the API is in a cross-platform requirement set, you should remove or edit the `isSetSupported` check. This will enable your add-in's feature on other platforms. Be sure to test the feature on those platforms when making this change.

> [!IMPORTANT]
> Your manifest cannot specify `ExcelApiOnline 1.1` as an activation requirement. It is not a valid value to use in the [Set element](../manifest/set.md).

## API list

The following table lists the Excel JavaScript APIs currently included in the `ExcelApiOnline` requirement set. For a complete list of all Excel JavaScript APIs (including `ExcelApiOnline` APIs and previously released APIs), see [all Excel JavaScript APIs](/javascript/api/excel?view=excel-js-online&preserve-view=true).

| Class | Fields | Description |
|:---|:---|:---|
|[AutoFilter](/javascript/api/excel/excel.autofilter)|[clearColumnCriteria(columnIndex: number)](/javascript/api/excel/excel.autofilter#clearColumnCriteria_columnIndex_)|Clears the column filter criteria of the AutoFilter.|
|[LinkedWorkbook](/javascript/api/excel/excel.linkedworkbook)|[breakLinks()](/javascript/api/excel/excel.linkedworkbook#breakLinks__)|Makes a request to break the links pointing to the linked workbook.|
||[id](/javascript/api/excel/excel.linkedworkbook#id)|The original URL pointing to the linked workbook.|
||[refresh()](/javascript/api/excel/excel.linkedworkbook#refresh__)|Makes a request to refresh the data retrieved from the linked workbook.|
|[LinkedWorkbookCollection](/javascript/api/excel/excel.linkedworkbookcollection)|[breakAllLinks()](/javascript/api/excel/excel.linkedworkbookcollection#breakAllLinks__)|Breaks all the links to the linked workbooks.|
||[getItem(key: string)](/javascript/api/excel/excel.linkedworkbookcollection#getItem_key_)|Gets information about a linked workbook by its URL.|
||[getItemOrNullObject(key: string)](/javascript/api/excel/excel.linkedworkbookcollection#getItemOrNullObject_key_)|Gets information about a linked workbook by its URL.|
||[items](/javascript/api/excel/excel.linkedworkbookcollection#items)|Gets the loaded child items in this collection.|
||[refreshAll()](/javascript/api/excel/excel.linkedworkbookcollection#refreshAll__)|Makes a request to refresh all the workbook links.|
||[workbookLinksRefreshMode](/javascript/api/excel/excel.linkedworkbookcollection#workbookLinksRefreshMode)|Represents the update mode of the workbook links.|
|[NamedSheetView](/javascript/api/excel/excel.namedsheetview)|[activate()](/javascript/api/excel/excel.namedsheetview#activate__)|Activates this sheet view.|
||[delete()](/javascript/api/excel/excel.namedsheetview#delete__)|Removes the sheet view from the worksheet.|
||[duplicate(name?: string)](/javascript/api/excel/excel.namedsheetview#duplicate_name_)|Creates a copy of this sheet view.|
||[name](/javascript/api/excel/excel.namedsheetview#name)|Gets or sets the name of the sheet view.|
|[NamedSheetViewCollection](/javascript/api/excel/excel.namedsheetviewcollection)|[add(name: string)](/javascript/api/excel/excel.namedsheetviewcollection#add_name_)|Creates a new sheet view with the given name.|
||[enterTemporary()](/javascript/api/excel/excel.namedsheetviewcollection#enterTemporary__)|Creates and activates a new temporary sheet view.|
||[exit()](/javascript/api/excel/excel.namedsheetviewcollection#exit__)|Exits the currently active sheet view.|
||[getActive()](/javascript/api/excel/excel.namedsheetviewcollection#getActive__)|Gets the worksheet's currently active sheet view.|
||[getCount()](/javascript/api/excel/excel.namedsheetviewcollection#getCount__)|Gets the number of sheet views in this worksheet.|
||[getItem(key: string)](/javascript/api/excel/excel.namedsheetviewcollection#getItem_key_)|Gets a sheet view using its name.|
||[getItemAt(index: number)](/javascript/api/excel/excel.namedsheetviewcollection#getItemAt_index_)|Gets a sheet view by its index in the collection.|
||[items](/javascript/api/excel/excel.namedsheetviewcollection#items)|Gets the loaded child items in this collection.|
|[TableRowCollection](/javascript/api/excel/excel.tablerowcollection)|[deleteRows(rows: number[] \| TableRow[])](/javascript/api/excel/excel.tablerowcollection#deleteRows_rows_)|Delete multiple rows from a table.|
||[deleteRowsAt(index: number, count?: number)](/javascript/api/excel/excel.tablerowcollection#deleteRowsAt_index__count_)|Delete a specified number of rows from a table, starting at a given index.|
|[Workbook](/javascript/api/excel/excel.workbook)|[linkedWorkbooks](/javascript/api/excel/excel.workbook#linkedWorkbooks)|Returns a collection of linked workbooks.|
|[Worksheet](/javascript/api/excel/excel.worksheet)|[namedSheetViews](/javascript/api/excel/excel.worksheet#namedSheetViews)|Returns a collection of sheet views that are present in the worksheet.|
||[onNameChanged](/javascript/api/excel/excel.worksheet#onNameChanged)|Occurs when the worksheet name is changed.|
||[onVisibilityChanged](/javascript/api/excel/excel.worksheet#onVisibilityChanged)|Occurs when the worksheet visibility is changed.|
|[WorksheetCollection](/javascript/api/excel/excel.worksheetcollection)|[onMoved](/javascript/api/excel/excel.worksheetcollection#onMoved)|Occurs when a worksheet is moved by a user within a workbook.|
||[onNameChanged](/javascript/api/excel/excel.worksheetcollection#onNameChanged)|Occurs when the worksheet name is changed in the worksheet collection.|
||[onVisibilityChanged](/javascript/api/excel/excel.worksheetcollection#onVisibilityChanged)|Occurs when the worksheet visibility is changed in the worksheet collection.|
|[WorksheetMovedEventArgs](/javascript/api/excel/excel.worksheetmovedeventargs)|[positionAfter](/javascript/api/excel/excel.worksheetmovedeventargs#positionAfter)|Gets the new position of the worksheet, after the move.|
||[positionBefore](/javascript/api/excel/excel.worksheetmovedeventargs#positionBefore)|Gets the previous position of the worksheet, prior to the move.|
||[source](/javascript/api/excel/excel.worksheetmovedeventargs#source)|The source of the event.|
||[type](/javascript/api/excel/excel.worksheetmovedeventargs#type)|Gets the type of the event.|
||[worksheetId](/javascript/api/excel/excel.worksheetmovedeventargs#worksheetId)|Gets the ID of the worksheet that was moved.|
|[WorksheetNameChangedEventArgs](/javascript/api/excel/excel.worksheetnamechangedeventargs)|[nameAfter](/javascript/api/excel/excel.worksheetnamechangedeventargs#nameAfter)|Gets the new name of the worksheet, after the name change.|
||[nameBefore](/javascript/api/excel/excel.worksheetnamechangedeventargs#nameBefore)|Gets the previous name of the worksheet, before the name changed.|
||[source](/javascript/api/excel/excel.worksheetnamechangedeventargs#source)|The source of the event.|
||[type](/javascript/api/excel/excel.worksheetnamechangedeventargs#type)|Gets the type of the event.|
||[worksheetId](/javascript/api/excel/excel.worksheetnamechangedeventargs#worksheetId)|Gets the ID of the worksheet with the new name.|
|[WorksheetVisibilityChangedEventArgs](/javascript/api/excel/excel.worksheetvisibilitychangedeventargs)|[source](/javascript/api/excel/excel.worksheetvisibilitychangedeventargs#source)|The source of the event.|
||[type](/javascript/api/excel/excel.worksheetvisibilitychangedeventargs#type)|Gets the type of the event.|
||[visibilityAfter](/javascript/api/excel/excel.worksheetvisibilitychangedeventargs#visibilityAfter)|Gets the new visibility setting of the worksheet, after the visibility change.|
||[visibilityBefore](/javascript/api/excel/excel.worksheetvisibilitychangedeventargs#visibilityBefore)|Gets the previous visibility setting of the worksheet, before the visibility change.|
||[worksheetId](/javascript/api/excel/excel.worksheetvisibilitychangedeventargs#worksheetId)|Gets the ID of the worksheet whose visibility has changed.|

## See also

- [Excel JavaScript API Reference Documentation](/javascript/api/excel?view=excel-js-online&preserve-view=true)
- [Excel JavaScript preview APIs](excel-preview-apis.md)
- [Excel JavaScript API requirement sets](excel-api-requirement-sets.md)
