---
title: Excel JavaScript API requirement set 1.5
description: 'Details about the ExcelApi 1.5 requirement set'
ms.date: 07/25/2019
ms.prod: excel
localization_priority: Normal
---

# What's new in Excel JavaScript API 1.5

ExcelApi 1.5 adds Custom XML parts. These are accessible through the [custom XML parts collection](/javascript/api/excel/excel.workbook#customxmlparts) in the workbook object.

## Custom XML part

* Get custom XML parts using their ID.
* Get a new scoped collection of custom XML parts whose namespaces match the given namespace.
* Get an XML string associated with a part.
* Provide the ID and namespace of a part.
* Add a new custom XML part to the workbook.
* Set an entire XML part.
* Delete a custom XML part.
* Delete an attribute with the given name from the element identified by xpath.
* Query the XML content by xpath.
* Insert, update, and delete attributes.

## API list

To see a complete list of all APIs supported by this requirement set (including previously released APIs), [click here to see a version-specific of the API reference documentation](/javascript/api/excel?view=excel-js-1.5).

| Class | Fields | Description |
|:---|:---|:---|
|[CustomXmlPart](/javascript/api/excel/excel.customxmlpart)|[delete()](/javascript/api/excel/excel.customxmlpart#delete--)|Deletes the custom XML part.|
||[getXml()](/javascript/api/excel/excel.customxmlpart#getxml--)|Gets the custom XML part's full XML content.|
||[id](/javascript/api/excel/excel.customxmlpart#id)|The custom XML part's ID. Read-only.|
||[namespaceUri](/javascript/api/excel/excel.customxmlpart#namespaceuri)|The custom XML part's namespace URI. Read-only.|
||[setXml(xml: string)](/javascript/api/excel/excel.customxmlpart#setxml-xml-)|Sets the custom XML part's full XML content.|
|[CustomXmlPartCollection](/javascript/api/excel/excel.customxmlpartcollection)|[add(xml: string)](/javascript/api/excel/excel.customxmlpartcollection#add-xml-)|Adds a new custom XML part to the workbook.|
||[getByNamespace(namespaceUri: string)](/javascript/api/excel/excel.customxmlpartcollection#getbynamespace-namespaceuri-)|Gets a new scoped collection of custom XML parts whose namespaces match the given namespace.|
||[getCount()](/javascript/api/excel/excel.customxmlpartcollection#getcount--)|Gets the number of CustomXml parts in the collection.|
||[getItem(id: string)](/javascript/api/excel/excel.customxmlpartcollection#getitem-id-)|Gets a custom XML part based on its ID.|
||[getItemOrNullObject(id: string)](/javascript/api/excel/excel.customxmlpartcollection#getitemornullobject-id-)|Gets a custom XML part based on its ID.|
||[items](/javascript/api/excel/excel.customxmlpartcollection#items)|Gets the loaded child items in this collection.|
|[CustomXmlPartCollectionLoadOptions](/javascript/api/excel/excel.customxmlpartcollectionloadoptions)|[$all](/javascript/api/excel/excel.customxmlpartcollectionloadoptions#$all)||
||[id](/javascript/api/excel/excel.customxmlpartcollectionloadoptions#id)|For EACH ITEM in the collection: The custom XML part's ID. Read-only.|
||[namespaceUri](/javascript/api/excel/excel.customxmlpartcollectionloadoptions#namespaceuri)|For EACH ITEM in the collection: The custom XML part's namespace URI. Read-only.|
|[CustomXmlPartData](/javascript/api/excel/excel.customxmlpartdata)|[id](/javascript/api/excel/excel.customxmlpartdata#id)|The custom XML part's ID. Read-only.|
||[namespaceUri](/javascript/api/excel/excel.customxmlpartdata#namespaceuri)|The custom XML part's namespace URI. Read-only.|
|[CustomXmlPartLoadOptions](/javascript/api/excel/excel.customxmlpartloadoptions)|[$all](/javascript/api/excel/excel.customxmlpartloadoptions#$all)||
||[id](/javascript/api/excel/excel.customxmlpartloadoptions#id)|The custom XML part's ID. Read-only.|
||[namespaceUri](/javascript/api/excel/excel.customxmlpartloadoptions#namespaceuri)|The custom XML part's namespace URI. Read-only.|
|[CustomXmlPartScopedCollection](/javascript/api/excel/excel.customxmlpartscopedcollection)|[getCount()](/javascript/api/excel/excel.customxmlpartscopedcollection#getcount--)|Gets the number of CustomXML parts in this collection.|
||[getItem(id: string)](/javascript/api/excel/excel.customxmlpartscopedcollection#getitem-id-)|Gets a custom XML part based on its ID.|
||[getItemOrNullObject(id: string)](/javascript/api/excel/excel.customxmlpartscopedcollection#getitemornullobject-id-)|Gets a custom XML part based on its ID.|
||[getOnlyItem()](/javascript/api/excel/excel.customxmlpartscopedcollection#getonlyitem--)|If the collection contains exactly one item, this method returns it.|
||[getOnlyItemOrNullObject()](/javascript/api/excel/excel.customxmlpartscopedcollection#getonlyitemornullobject--)|If the collection contains exactly one item, this method returns it.|
||[items](/javascript/api/excel/excel.customxmlpartscopedcollection#items)|Gets the loaded child items in this collection.|
|[CustomXmlPartScopedCollectionLoadOptions](/javascript/api/excel/excel.customxmlpartscopedcollectionloadoptions)|[$all](/javascript/api/excel/excel.customxmlpartscopedcollectionloadoptions#$all)||
||[id](/javascript/api/excel/excel.customxmlpartscopedcollectionloadoptions#id)|For EACH ITEM in the collection: The custom XML part's ID. Read-only.|
||[namespaceUri](/javascript/api/excel/excel.customxmlpartscopedcollectionloadoptions#namespaceuri)|For EACH ITEM in the collection: The custom XML part's namespace URI. Read-only.|
|[PivotTable](/javascript/api/excel/excel.pivottable)|[id](/javascript/api/excel/excel.pivottable#id)|Id of the PivotTable. Read-only.|
|[PivotTableCollectionLoadOptions](/javascript/api/excel/excel.pivottablecollectionloadoptions)|[id](/javascript/api/excel/excel.pivottablecollectionloadoptions#id)|For EACH ITEM in the collection: Id of the PivotTable. Read-only.|
|[PivotTableData](/javascript/api/excel/excel.pivottabledata)|[id](/javascript/api/excel/excel.pivottabledata#id)|Id of the PivotTable. Read-only.|
|[PivotTableLoadOptions](/javascript/api/excel/excel.pivottableloadoptions)|[id](/javascript/api/excel/excel.pivottableloadoptions#id)|Id of the PivotTable. Read-only.|
|[Runtime](/javascript/api/excel/excel.runtime)|[set(properties: Excel.Runtime)](/javascript/api/excel/excel.runtime#set-properties-)|Sets multiple properties on the object at the same time, based on an existing loaded object.|
||[set(properties: Interfaces.RuntimeUpdateData, options?: OfficeExtension.UpdateOptions)](/javascript/api/excel/excel.runtime#set-properties--options-)|Sets multiple properties of an object at the same time. You can pass either a plain object with the appropriate properties, or another API object of the same type.|
|[RuntimeLoadOptions](/javascript/api/excel/excel.runtimeloadoptions)|[$all](/javascript/api/excel/excel.runtimeloadoptions#$all)||
|[Workbook](/javascript/api/excel/excel.workbook)|[customXmlParts](/javascript/api/excel/excel.workbook#customxmlparts)|Represents the collection of custom XML parts contained by this workbook. Read-only.|
|[WorkbookData](/javascript/api/excel/excel.workbookdata)|[customXmlParts](/javascript/api/excel/excel.workbookdata#customxmlparts)|Represents the collection of custom XML parts contained by this workbook. Read-only.|
|[Worksheet](/javascript/api/excel/excel.worksheet)|[getNext(visibleOnly?: boolean)](/javascript/api/excel/excel.worksheet#getnext-visibleonly-)|Gets the worksheet that follows this one. If there are no worksheets following this one, this method will throw an error.|
||[getNextOrNullObject(visibleOnly?: boolean)](/javascript/api/excel/excel.worksheet#getnextornullobject-visibleonly-)|Gets the worksheet that follows this one. If there are no worksheets following this one, this method will return a null object.|
||[getPrevious(visibleOnly?: boolean)](/javascript/api/excel/excel.worksheet#getprevious-visibleonly-)|Gets the worksheet that precedes this one. If there are no previous worksheets, this method will throw an error.|
||[getPreviousOrNullObject(visibleOnly?: boolean)](/javascript/api/excel/excel.worksheet#getpreviousornullobject-visibleonly-)|Gets the worksheet that precedes this one. If there are no previous worksheets, this method will return a null objet.|
|[WorksheetCollection](/javascript/api/excel/excel.worksheetcollection)|[getFirst(visibleOnly?: boolean)](/javascript/api/excel/excel.worksheetcollection#getfirst-visibleonly-)|Gets the first worksheet in the collection.|
||[getLast(visibleOnly?: boolean)](/javascript/api/excel/excel.worksheetcollection#getlast-visibleonly-)|Gets the last worksheet in the collection.|

## See also

- [Excel JavaScript API Reference Documentation](/javascript/api/excel&view=excel-js-1.5)
- [Excel JavaScript API requirement sets](./excel-api-requirement-sets.md)
