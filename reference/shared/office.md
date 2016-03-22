

# Office object
Represents an instance of an add-in, which provides access to the top-level objects of the API.

|||
|:-----|:-----|
|**Hosts:**|Access, Excel, Outlook, PowerPoint, Project, Word|
|**Last changed in**|1.1|

```js
Office
```


## Members


**Properties**

|||
|:-----|:-----|
|Name|Description|
|[context](../../reference/shared/office.context.md)|Gets the Context object that represents the runtime environment of the add-in and provides access to the top-level objects of the API.|
|[cast.item](../../reference/shared/office.cast.item.md)|Provides IntelliSense in Visual Studio specific to compose or read mode messages and appointments. <br/><br/><blockquote>**Note**  Only applicable at design time when developing Outlook add-ins in Visual Studio. </blockquote>|

**Methods**

|||
|:-----|:-----|
|Name|Description|
|[select](../../reference/shared/office.select.md)|Creates a promise to return a binding based on the selector string passed in.|
|[useShortNamespace](../../reference/shared/office.useshortnamespace.md)|Toggles on and off the  **Office** alias for the full **Microsoft.Office.WebExtension** namespace.|

**Events**

|||
|:-----|:-----|
|Name|Description|
|[initialize](../../reference/shared/office.initialize.md)|Occurs when the runtime environment is loaded and the add-in is ready to start interacting with the application and hosted document.|

## Remarks

The  **Office** object enables the developer to implement a callback function for the Initialize event and provides access to the [Context](../../reference/shared/asyncresult.context.md) object.


## Support details


A capital Y in the following matrix indicates that this object is supported in the corresponding Office host application. An empty cell indicates that the Office host application doesn't support this object.

For more information about Office host application and server requirements, see [Requirements for running Office Add-ins](../../docs/overview/requirements-for-running-office-add-ins.md).


||**Office for Windows desktop**|**Office Online (in browser)**|**Office for iPad**|**OWA for Devices**|**Outlook for Mac**|
|:-----|:-----|:-----|:-----|:-----|:-----|
|**Access**||Y||||
|**Excel**|Y|Y|Y|||
|**Outlook**|Y|Y||Y|Y|
|**PowerPoint**|Y|Y|Y|||
|**Project**|Y|||||
|**Word**|Y|Y|Y|||

|||
|:-----|:-----|
|**Add-in types**|Content, Outlook, task pane|
|**Library**|Office.js|
|**Namespace**|Office|

## Support history


|**Version**|**Changes**|
|:-----|:-----|
|1.1|Added support for Excel, PowerPoint, and Word in Office for iPad.|
|1.1|<ul><li>For <a href="6c4b2c16-d4fb-4ecf-b72c-1e33b205daaf.htm">context</a>, added support for getting the runtime context in content add-ins for Access.</p></li><li><p>For <a href="23aeb136-da1f-4127-a798-99dc27bc4dae.htm">select</a>, added support for selecting table bindings in content add-ins for Access.</li><li>For <a href="9a4d5c7d-fcc4-4e8f-bef2-f2a8d8b4ae00.htm">useShortNamespace</a>, added support for content add-ins for Access.</li><li>For <a href="727adf79-a0b5-48d2-99c7-6642c2c334fc.htm">initialize</a>, added support for initialization in content add-ins for Access.</li></ul>|
|1.0|Introduced|

