
# Context object
Represents the runtime environment of the add-in and provides access to key objects of the API.

|||
|:-----|:-----|
|**Hosts:**|Access, Excel, Outlook, PowerPoint, Project, Word|
|**Last changed in**|1.1|

```
Office.context
```


## Members


**Properties**

|||
|:-----|:-----|
|Name|Description|
|[commerceAllowed](../../reference/shared/office.context.commerceallowed.md)|Gets whether the add-in is running on a platform that allows links to external payment systems.|
|[contentLanguage](../../reference/shared/office.context.contentlanguage.md)|Gets the locale (language) for data as it is stored in the document or item.|
|[displayLanguage](../../reference/shared/office.context.displaylanguage.md)|Gets the locale (language) for the UI of the hosting application.|
|[document](../../reference/shared/office.context.document.md)|Gets an object that represents the document the content or task pane add-in is interacting with.|
|[mailbox](../../reference/shared/office.context.mailbox.md)|Gets the  **mailbox** object that provides access to members of the API that are specifically for Outlook add-ins.|
|[officeTheme](../../reference/shared/office.context.officetheme.md)|Provides access to the properties for Office theme colors|
|[roamingSettings](../../reference/shared/office.context.roamingsettings.md)|Gets an object that represents the saved custom settings of the add-in.|
|[touchEnabled](../../reference/shared/office.context.touchenabled.md)|Gets whether the add-in is running in an Office host application that is touch enabled.|

## Remarks

The  **Context** object provides access to key objects in the JavaScript API for Office.


## Support details



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
|1.1|Added  **commerceAllowed** and **touchEnabledAdded** properties (Excel, PowerPoint and Word on Office for iPad only).|
|1.1|Added support for add-ins with Excel and Word on Office for iPad.|
|1.1|For [contentLanguage](../../reference/shared/office.context.contentlanguage.md), [displayLanguage](../../reference/shared/office.context.displaylanguage.md), and [document](../../reference/shared/office.context.document.md), added support for content add-ins for Access.|
|1.0|Introduced|
