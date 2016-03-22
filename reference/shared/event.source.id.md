

# event.source.id
Gets the id of the control that triggered calling this function.

****

|||
|:-----|:-----|
|**Hosts:**Outlook|**Add-in type:** Outlook|
|**Available in [requirement sets](../../docs/overview/specify-office-hosts-and-api-requirements.md)**|Mailbox|
|**Last changed in Mailbox**|1.3|
|**Applicable Outlook modes**|Read and Compose|



```js
event.source.id;
```


## Return value

The id of the control that triggered calling this function. The id comes from the manifest.


## Support details


A capital Y in the following table indicates that this property is supported in the corresponding Outlook host application. An empty cell indicates that the Outlook host application doesn't support this property.

For more information about Office host application and server requirements, see [Requirements for running Office Add-ins](../../docs/overview/requirements-for-running-office-add-ins.md).

 **Important:** Add-in commands and the APIs associated with them currently work only in Outlook in [Office 2016 Preview](https://products.office.com/en-us/office-2016-preview) on Windows desktop.


**Supported hosts, by platform**

| |**Office for Windows desktop**|**Office Online (in browser)**|**OWA for Devices**|
|:-----|:-----|:-----|:-----|
|**Outlook**|Y|||

|||
|:-----|:-----|
|**Available in requirement sets**|Mailbox|
|**Minimum permission level**|[ReadWriteItem](../../docs/outlook/understanding-outlook-add-in-permissions.md)|
|**Add-in types**|Outlook|
|**Library**|Office.js|
|**Namespace**|Office|

## Support history




|**Version**|**Changes**|
|:-----|:-----|
|1.3|Introduced|
