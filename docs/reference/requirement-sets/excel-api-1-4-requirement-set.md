---
title: Excel JavaScript API requirement set 1.44
description: 'Details about the ExcelApi 1.4 requirement set'
ms.date: 07/15/2019
ms.prod: excel
localization_priority: Normal
---

# What's new in Excel JavaScript API 1.4

The following are the new additions to the Excel JavaScript APIs in requirement set 1.4.

## Named item add and new properties

New properties:

* `comment`
* `scope` - Worksheet or workbook scoped items.
* `worksheet` - Returns the worksheet on which the named item is scoped to.

New methods:

* `add(name: string, reference: Range or string, comment: string)` - Adds a new name to the collection of the given scope.
* `addFormulaLocal(name: string, formula: string, comment: string)` - Adds a new name to the collection of the given scope using the user's locale for the formula.

## Settings API in the Excel namespace

The [Setting](/javascript/api/excel/excel.setting) object represents a key:value pair for a setting persisted to the document. The functionality of `Excel.Setting` is equivalent to `Office.Settings`, but uses the batched API syntax, rather than the Common API's callback model.

APIs include `getItem()` to get setting entry via the key and `add()` to add the specified key:value setting pair to the workbook.

## Others

* Set the table column name.
* Add a table column to the end of the table.
* Add multiple rows to a table at a time.
* `range.getColumnsAfter(count: number)` and `range.getColumnsBefore(count: number)` to get a certain number of columns to the right/left of the current Range object.
* The [get item or null object function](../../excel/excel-add-ins-advanced-concepts.md#ornullobject-methods): This functionality allows getting object using a key. If the object does not exist, the returned object's `isNullObject` property will be true. This allows developers to check if an object exists or not without having to handle it through exception handling. The `*OrNullObject` method is available on most collection objects.

```javascript
worksheet.getItemOrNullObject("itemName")
```

## API list

| Class | Fields | Description |
|:---|:---|:---|
|[BindingCollection](/javascript/api/excel/excel.bindingcollection)|[getCount()](/javascript/api/excel/excel.bindingcollection#getcount--)|Gets the number of bindings in the collection.|
||[getItemOrNullObject(id: string)](/javascript/api/excel/excel.bindingcollection#getitemornullobject-id-)|Gets a binding object by ID. If the binding object does not exist, will return a null object.|
|[ChartCollection](/javascript/api/excel/excel.chartcollection)|[getCount()](/javascript/api/excel/excel.chartcollection#getcount--)|Returns the number of charts in the worksheet.|
||[getItemOrNullObject(name: string)](/javascript/api/excel/excel.chartcollection#getitemornullobject-name-)|Gets a chart using its name. If there are multiple charts with the same name, the first one will be returned.|
|[ChartPointsCollection](/javascript/api/excel/excel.chartpointscollection)|[getCount()](/javascript/api/excel/excel.chartpointscollection#getcount--)|Returns the number of chart points in the series.|
|[ChartSeriesCollection](/javascript/api/excel/excel.chartseriescollection)|[getCount()](/javascript/api/excel/excel.chartseriescollection#getcount--)|Returns the number of series in the collection.|
|[NamedItem](/javascript/api/excel/excel.nameditem)|[comment](/javascript/api/excel/excel.nameditem#comment)|Represents the comment associated with this name.|
||[delete()](/javascript/api/excel/excel.nameditem#delete--)|Deletes the given name.|
||[getRangeOrNullObject()](/javascript/api/excel/excel.nameditem#getrangeornullobject--)|Returns the range object that is associated with the name. Returns a null object if the named item's type is not a range.|
||[scope](/javascript/api/excel/excel.nameditem#scope)|Indicates whether the name is scoped to the workbook or to a specific worksheet. Possible values are: Worksheet, Workbook. Read-only.|
||[worksheet](/javascript/api/excel/excel.nameditem#worksheet)|Returns the worksheet on which the named item is scoped to. Throws an error if the item is scoped to the workbook instead.|
||[worksheetOrNullObject](/javascript/api/excel/excel.nameditem#worksheetornullobject)|Returns the worksheet on which the named item is scoped to. Returns a null object if the item is scoped to the workbook instead.|
|[NamedItemCollection](/javascript/api/excel/excel.nameditemcollection)|[add(name: string, reference: Range \| string, comment?: string)](/javascript/api/excel/excel.nameditemcollection#add-name--reference--comment-)|Adds a new name to the collection of the given scope.|
||[addFormulaLocal(name: string, formula: string, comment?: string)](/javascript/api/excel/excel.nameditemcollection#addformulalocal-name--formula--comment-)|Adds a new name to the collection of the given scope using the user's locale for the formula.|
||[getCount()](/javascript/api/excel/excel.nameditemcollection#getcount--)|Gets the number of named items in the collection.|
||[getItemOrNullObject(name: string)](/javascript/api/excel/excel.nameditemcollection#getitemornullobject-name-)|Gets a NamedItem object using its name. If the nameditem object does not exist, will return a null object.|
|[NamedItemCollectionLoadOptions](/javascript/api/excel/excel.nameditemcollectionloadoptions)|[comment](/javascript/api/excel/excel.nameditemcollectionloadoptions#comment)|For EACH ITEM in the collection: Represents the comment associated with this name.|
||[scope](/javascript/api/excel/excel.nameditemcollectionloadoptions#scope)|For EACH ITEM in the collection: Indicates whether the name is scoped to the workbook or to a specific worksheet. Possible values are: Worksheet, Workbook. Read-only.|
||[worksheet](/javascript/api/excel/excel.nameditemcollectionloadoptions#worksheet)|For EACH ITEM in the collection: Returns the worksheet on which the named item is scoped to. Throws an error if the item is scoped to the workbook instead.|
||[worksheetOrNullObject](/javascript/api/excel/excel.nameditemcollectionloadoptions#worksheetornullobject)|For EACH ITEM in the collection: Returns the worksheet on which the named item is scoped to. Returns a null object if the item is scoped to the workbook instead.|
|[NamedItemData](/javascript/api/excel/excel.nameditemdata)|[comment](/javascript/api/excel/excel.nameditemdata#comment)|Represents the comment associated with this name.|
||[scope](/javascript/api/excel/excel.nameditemdata#scope)|Indicates whether the name is scoped to the workbook or to a specific worksheet. Possible values are: Worksheet, Workbook. Read-only.|
|[NamedItemLoadOptions](/javascript/api/excel/excel.nameditemloadoptions)|[comment](/javascript/api/excel/excel.nameditemloadoptions#comment)|Represents the comment associated with this name.|
||[scope](/javascript/api/excel/excel.nameditemloadoptions#scope)|Indicates whether the name is scoped to the workbook or to a specific worksheet. Possible values are: Worksheet, Workbook. Read-only.|
||[worksheet](/javascript/api/excel/excel.nameditemloadoptions#worksheet)|Returns the worksheet on which the named item is scoped to. Throws an error if the item is scoped to the workbook instead.|
||[worksheetOrNullObject](/javascript/api/excel/excel.nameditemloadoptions#worksheetornullobject)|Returns the worksheet on which the named item is scoped to. Returns a null object if the item is scoped to the workbook instead.|
|[NamedItemUpdateData](/javascript/api/excel/excel.nameditemupdatedata)|[comment](/javascript/api/excel/excel.nameditemupdatedata#comment)|Represents the comment associated with this name.|
|[PivotTableCollection](/javascript/api/excel/excel.pivottablecollection)|[getCount()](/javascript/api/excel/excel.pivottablecollection#getcount--)|Gets the number of pivot tables in the collection.|
||[getItemOrNullObject(name: string)](/javascript/api/excel/excel.pivottablecollection#getitemornullobject-name-)|Gets a PivotTable by name. If the PivotTable does not exist, will return a null object.|
|[Range](/javascript/api/excel/excel.range)|[getIntersectionOrNullObject(anotherRange: Range \| string)](/javascript/api/excel/excel.range#getintersectionornullobject-anotherrange-)|Gets the range object that represents the rectangular intersection of the given ranges. If no intersection is found, will return a null object.|
||[getUsedRangeOrNullObject(valuesOnly?: boolean)](/javascript/api/excel/excel.range#getusedrangeornullobject-valuesonly-)|Returns the used range of the given range object. If there are no used cells within the range, this function will return a null object.|
|[RangeViewCollection](/javascript/api/excel/excel.rangeviewcollection)|[getCount()](/javascript/api/excel/excel.rangeviewcollection#getcount--)|Gets the number of RangeView objects in the collection.|
|[Setting](/javascript/api/excel/excel.setting)|[delete()](/javascript/api/excel/excel.setting#delete--)|Deletes the setting.|
||[](/javascript/api/excel/excel.setting#datejsonprefix)||
||[](/javascript/api/excel/excel.setting#datejsonsuffix)||
||[](/javascript/api/excel/excel.setting#replacestringdatewithdate)||
||[key](/javascript/api/excel/excel.setting#key)|Returns the key that represents the id of the Setting. Read-only.|
||[set(properties: Excel.Setting)](/javascript/api/excel/excel.setting#set-properties-)|Sets multiple properties on the object at the same time, based on an existing loaded object.|
||[set(properties: Interfaces.SettingUpdateData, options?: OfficeExtension.UpdateOptions)](/javascript/api/excel/excel.setting#set-properties--options-)|Sets multiple properties of an object at the same time. You can pass either a plain object with the appropriate properties, or another API object of the same type.|
||[value](/javascript/api/excel/excel.setting#value)|Represents the value stored for this setting.|
|[SettingCollection](/javascript/api/excel/excel.settingcollection)|[add(key: string, value: string \| number \| boolean \| Date \| Array<any> \| any)](/javascript/api/excel/excel.settingcollection#add-key--value-)|Sets or adds the specified setting to the workbook.|
||[getCount()](/javascript/api/excel/excel.settingcollection#getcount--)|Gets the number of Settings in the collection.|
||[getItem(key: string)](/javascript/api/excel/excel.settingcollection#getitem-key-)|Gets a Setting entry via the key.|
||[getItemOrNullObject(key: string)](/javascript/api/excel/excel.settingcollection#getitemornullobject-key-)|Gets a Setting entry via the key. If the Setting does not exist, will return a null object.|
||[items](/javascript/api/excel/excel.settingcollection#items)|Gets the loaded child items in this collection.|
||[onSettingsChanged](/javascript/api/excel/excel.settingcollection#onsettingschanged)|Occurs when the Settings in the document are changed.|
|[SettingCollectionLoadOptions](/javascript/api/excel/excel.settingcollectionloadoptions)|[$all](/javascript/api/excel/excel.settingcollectionloadoptions#$all)||
||[key](/javascript/api/excel/excel.settingcollectionloadoptions#key)|For EACH ITEM in the collection: Returns the key that represents the id of the Setting. Read-only.|
||[value](/javascript/api/excel/excel.settingcollectionloadoptions#value)|For EACH ITEM in the collection: Represents the value stored for this setting.|
|[SettingData](/javascript/api/excel/excel.settingdata)|[key](/javascript/api/excel/excel.settingdata#key)|Returns the key that represents the id of the Setting. Read-only.|
||[value](/javascript/api/excel/excel.settingdata#value)|Represents the value stored for this setting.|
|[SettingLoadOptions](/javascript/api/excel/excel.settingloadoptions)|[$all](/javascript/api/excel/excel.settingloadoptions#$all)||
||[key](/javascript/api/excel/excel.settingloadoptions#key)|Returns the key that represents the id of the Setting. Read-only.|
||[value](/javascript/api/excel/excel.settingloadoptions#value)|Represents the value stored for this setting.|
|[SettingUpdateData](/javascript/api/excel/excel.settingupdatedata)|[value](/javascript/api/excel/excel.settingupdatedata#value)|Represents the value stored for this setting.|
|[SettingsChangedEventArgs](/javascript/api/excel/excel.settingschangedeventargs)|[settings](/javascript/api/excel/excel.settingschangedeventargs#settings)|Gets the Setting object that represents the binding that raised the SettingsChanged event|
|[TableCollection](/javascript/api/excel/excel.tablecollection)|[getCount()](/javascript/api/excel/excel.tablecollection#getcount--)|Gets the number of tables in the collection.|
||[getItemOrNullObject(key: string)](/javascript/api/excel/excel.tablecollection#getitemornullobject-key-)|Gets a table by Name or ID. If the table does not exist, will return a null object.|
|[TableColumnCollection](/javascript/api/excel/excel.tablecolumncollection)|[getCount()](/javascript/api/excel/excel.tablecolumncollection#getcount--)|Gets the number of columns in the table.|
||[getItemOrNullObject(key: number \| string)](/javascript/api/excel/excel.tablecolumncollection#getitemornullobject-key-)|Gets a column object by Name or ID. If the column does not exist, will return a null object.|
|[TableRowCollection](/javascript/api/excel/excel.tablerowcollection)|[getCount()](/javascript/api/excel/excel.tablerowcollection#getcount--)|Gets the number of rows in the table.|
|[Workbook](/javascript/api/excel/excel.workbook)|[settings](/javascript/api/excel/excel.workbook#settings)|Represents a collection of Settings associated with the workbook. Read-only.|
|[WorkbookData](/javascript/api/excel/excel.workbookdata)|[settings](/javascript/api/excel/excel.workbookdata#settings)|Represents a collection of Settings associated with the workbook. Read-only.|
|[Worksheet](/javascript/api/excel/excel.worksheet)|[getUsedRangeOrNullObject(valuesOnly?: boolean)](/javascript/api/excel/excel.worksheet#getusedrangeornullobject-valuesonly-)|The used range is the smallest range that encompasses any cells that have a value or formatting assigned to them. If the entire worksheet is blank, this function will return a null object.|
||[names](/javascript/api/excel/excel.worksheet#names)|Collection of names scoped to the current worksheet. Read-only.|
|[WorksheetCollection](/javascript/api/excel/excel.worksheetcollection)|[getCount(visibleOnly?: boolean)](/javascript/api/excel/excel.worksheetcollection#getcount-visibleonly-)|Gets the number of worksheets in the collection.|
||[getItemOrNullObject(key: string)](/javascript/api/excel/excel.worksheetcollection#getitemornullobject-key-)|Gets a worksheet object using its Name or ID. If the worksheet does not exist, will return a null object.|
|[WorksheetData](/javascript/api/excel/excel.worksheetdata)|[names](/javascript/api/excel/excel.worksheetdata#names)|Collection of names scoped to the current worksheet. Read-only.|

## See also

- [Excel JavaScript API Reference Documentation](/javascript/api/excel&view=excel-js-1.4)
- [Excel JavaScript API requirement sets](./excel-api-requirement-sets.md)
