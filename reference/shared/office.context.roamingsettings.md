
# Context.roamingSettings property
Gets an object that represents the custom settings or state of a Outlook add-in saved to a user's mailbox.

|||
|:-----|:-----|
|**Hosts:**|Outlook|
|**Available in [Requirement set](../../docs/overview/specify-office-hosts-and-api-requirements.md)**|Mailbox|
|**Last changed in**|1.0|

```
var appSettings = office.context.roamingSettings;
```


## Return value

A [RoamingSettings](http://msdn.microsoft.com/library/cf21bb08-7274-4ad6-ae9e-b2c12f92abc9%28Office.15%29.aspx) object.


## Remarks

The  **RoamingSettings** object lets you store and access data for a Outlook add-in that is stored in a user's mailbox, so that is available to that add-in when it is running from any host client application used to access that mailbox.


## Support details


A capital Y in the following matrix indicates that this method is supported in the corresponding Office host application. An empty cell indicates that the Office host application doesn't support this method.

For more information about Office host application and server requirements, see [Requirements for running Office Add-ins](../../docs/overview/requirements-for-running-office-add-ins.md).


||**Office for Windows desktop**|**Office Online (in browser)**|**Outlook for Mac**|
|:-----|:-----|:-----|:-----|
|**Outlook**|Y|Y|Y|

|||
|:-----|:-----|
|**Available in requirement sets**|Mailbox|
|**Minimum permission level**|[Restricted](../../docs/develop/requesting-permissions-for-api-use-in-content-and-task-pane-add-ins.md)|
|**Add-in types**|Outlook|
|**Library**|Office.js|
|**Namespace**|Office|

## Support history



****


|**Version**|**Changes**|
|:-----|:-----|
|1.0|Introduced|
