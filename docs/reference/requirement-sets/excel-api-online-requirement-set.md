---
title: Excel JavaScript API online-only requirement set
description: 'Details about the ExcelApiOnline requirement set'
ms.date: 05/06/2020
ms.prod: excel
localization_priority: Normal
---

# Excel JavaScript API online-only requirement set

The `ExcelApiOnline` requirement set is a special requirement set that includes features that are only available for Excel on the web. APIs in this requirement set are considered to be production APIs (not subject to undocumented behavioral or structural changes) for the Excel on the web application. `ExcelApiOnline` are considered to be "preview" APIs for other platforms (Windows, Mac, iOS) and may not be supported by any of those platforms.

When APIs in the `ExcelApiOnline` requirement set are supported across all platforms, they will added to the next released requirement set (`ExcelApi 1.[NEXT]`). Once that new requirement is public, those APIs will be removed from `ExcelApiOnline`. Think of this as a similar promotion process as an API moving from preview to release.

> [!IMPORTANT]
> `ExcelApiOnline` is superset of the latest numbered requirement set.

> [!IMPORTANT]
> `ExcelApiOnline 1.1` is the only version of the online-only APIs. This is because Excel on the web will always have a single version available to users that is the latest version.

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

The following APIs are currently available for Excel on the web as part of the `ExcelApiOnline 1.1` requirement set.

| Class | Fields | Description |
|:---|:---|:---|
|[ChartAxisTitle](/javascript/api/excel/excel.chartaxistitle)|[textOrientation](/javascript/api/excel/excel.chartaxistitle#textorientation)|Specifies the angle to which the text is oriented for the chart axis title. The value should either be an integer from -90 to 90 or the integer 180 for vertically-oriented text.|
|[PivotTableScopedCollection](/javascript/api/excel/excel.pivottablescopedcollection)|[getCount()](/javascript/api/excel/excel.pivottablescopedcollection#getcount--)|Gets the number of PivotTables in the collection.|
||[getFirst()](/javascript/api/excel/excel.pivottablescopedcollection#getfirst--)|Gets the first PivotTable in the collection. The PivotTables in the collection are sorted top to bottom and left to right, such that top-left table is the first PivotTable in the collection.|
||[getItem(key: string)](/javascript/api/excel/excel.pivottablescopedcollection#getitem-key-)|Gets a PivotTable by name.|
||[getItemOrNullObject(name: string)](/javascript/api/excel/excel.pivottablescopedcollection#getitemornullobject-name-)|Gets a PivotTable by name. If the PivotTable does not exist, will return a null object.|
||[items](/javascript/api/excel/excel.pivottablescopedcollection#items)|Gets the loaded child items in this collection.|
|[Range](/javascript/api/excel/excel.range)|[getPivotTables(fullyContained?: boolean)](/javascript/api/excel/excel.range#getpivottables-fullycontained-)|Gets a scoped collection of PivotTables that overlap with the range.|

## See also

- [Excel JavaScript API Reference Documentation](/javascript/api/excel?view=excel-js-online)
- [Excel JavaScript preview APIs](./excel-preview-apis.md)
- [Excel JavaScript API requirement sets](./excel-api-requirement-sets.md)