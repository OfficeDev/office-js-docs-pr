
# Settings object 

Represents custom settings for a task pane or content add-in that are stored in the host document as name/value pairs.

|||
|:-----|:-----|
|**Hosts:**|Access, Excel, PowerPoint, Word|
|**Available in [Requirement set](../../docs/overview/specify-office-hosts-and-api-requirements.md)**|Settings|
|**Last changed in**|1.1|

```
Office.context.document.settings
```


## Members


**Methods**

|||
|:-----|:-----|
|Name|Description|
|[addHandlerAsync](../../reference/shared/settings.addhandlerasync.md)|Adds an event handler for the  **settingsChanged** event.|
|[get](../../reference/shared/settings.get.md)|Retrieves the specified setting.|
|[refreshAsync](../../reference/shared/settings.refreshasync.md)|Reads all settings persisted in the document and refreshes the add-in's copy of those settings held in memory.|
|[remove](../../reference/shared/settings.remove.md)|Removes the specified setting.|
|[removeHandlerAsync](../../reference/shared/settings.removehandlerasync.md)|Removes an event handler for the  **settingsChanged** event.|
|[saveAsync](../../reference/shared/settings.saveasync.md)|Saves the settings.|
|[set](../../reference/shared/settings.set.md)|Sets or creates the specified setting.|

**Events**


|**Name**|**Description**|
|:-----|:-----|
|[settingsChanged](../../reference/shared/settings.settingschangedevent.md)|Occurs when a setting is changed.|

## Remarks

The settings created by using the methods of the  **Settings** object are saved per add-in and per document. That is, they are available only to the add-in that created them, and only from the document in which they are saved.

The name of a setting is a  **string**, while the value can be a  **string**,  **number**,  **boolean**,  **null**,  **object**, or  **array**.

The  **Settings** object is automatically loaded as part of the [Document](../../reference/shared/document.md) object, and is available by calling the [settings](../../reference/shared/document.settings.md) property of that object when the add-in is activated. The developer is responsible for calling the [saveAsync](../../reference/shared/settings.saveasync.md) method after adding or deleting settings to save the settings in the document.


## Support details

A capital Y in the following matrix indicates that this object is supported in the corresponding Office host application. An empty cell indicates that the Office host application doesn't support this object.

For more information about Office host application and server requirements, see [Requirements for running Office Add-ins](../../docs/overview/requirements-for-running-office-add-ins.md).


|
|
|****|**Office for Windows desktop**|**Office Online(in browser)**|**Office for iPad**|
|:-----|:-----|:-----|:-----|
|**Access**||√||
|**Excel**|√|√|√|
|**PowerPoint**|√|√|√|
|**Word**|√||√|

|||
|:-----|:-----|
|**Available in requirement sets**|Settings|
|**Add-in types**|Content, task pane|
|**Library**|Office.js|
|**Namespace**|Office|

## Support history


|**Version**|**Changes**|
|:-----|:-----|
|1.1|Added support for Excel, PowerPoint, and Word in Office for iPad.|
|1.1|For methods **addHandlerAsync** and **removeHandlerAsync**, added support  to add and remove event handlers for the event in content add-ins for Access. For methods **get**, **refreshAsync**, **remove**, **saveAsync**, and **set**, added support for custom settings in content add-ins for Access.|
|1.0|Introduced|
