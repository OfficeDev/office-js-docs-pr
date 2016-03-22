
# officeTheme.controlBackgroundColor property
Gets the Office theme control background color.

 **Important:** This API currently works only in Excel, Outlook, PowerPoint, and Word in [Office 2016 Preview](https://products.office.com/en-us/office-2016-preview) on Windows desktop.



|||
|:-----|:-----|
|**Hosts:**|Excel, Outlook, PowerPoint, Word|
|**Available in [Requirement set](../../docs/overview/specify-office-hosts-and-api-requirements.md)**|Not in a set|
|**Added in**|1.3|

```
var controlBackgroundColor = Office.context.officeTheme.controlBackgroundColor;
```


## Return value

A hex color triplet.


## Remarks

The colors returned correspond to the values of the Office theme selected by the user with  **File** > **Office Account** > **Office Theme** UI, which is applied across all Office host applications.


## Support details


A capital Y in the following matrix indicates that this method is supported in the corresponding Office host application. An empty cell indicates that the Office host application doesn't support this method.

For more information about Office host application and server requirements, see [Requirements for running Office Add-ins](../../docs/overview/requirements-for-running-office-add-ins.md).


**Supported hosts, by platform**


||**Office for Windows desktop**|**Office Online (in browser)**|**Office for iPad**|**OWA for Devices**|
|:-----|:-----|:-----|:-----|:-----|
|**Excel**|Y||||
|**Outlook**|Y||||
|**PowerPoint**|Y||||
|**Word**|Y||||

|||
|:-----|:-----|
|**Minimum permission level**|[Restricted](../../docs/develop/requesting-permissions-for-api-use-in-content-and-task-pane-add-ins.md)|
|**Add-in types**|Content, task pane, Outlook|
|**Library**|Office.js|
|**Namespace**|Office|

## Support history



****


|**Version**|**Changes**|
|:-----|:-----|
|1.3|Introduced|
