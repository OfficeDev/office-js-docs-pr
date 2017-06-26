
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

| Member | Type | Description | 
|:-------|:-----|:------------|
|commerceAllowed |bool|Returns True if developers can display sell or upgrade UI in the add-in on that platform; otherwise returns False.|
|contentLanguage | string | Gets the locale (language) specified by the user for editing the document or item.|
|displayLanguage|string|Gets the locale (language) specified by the user for the UI of the Office host application.|
|document| [Document object](office.context.document.md)|Gets an object that represents the document the content or task pane add-in is interacting with.|
|host|string|Contains the host in which the add-in (web application) is running in. Possible values are: Word, Excel, PowerPoint, Outlook, OneNote, Project, Access|
|officeTheme|[OfficeTheme object](office.context.officetheme.md)|Provides access to the properties for Office theme colors.|
|platform|string|Provides the platform on which the add-in is running. Possible values are: PC, OfficeOnline, Mac, iOS|
|requirements|object|Offers `requirements.isSetSupported()` method to check if the specified requirement set is supported by the host Office application. <br/> `isSetSupported(name: string, minVersion?: number): boolean;` <br> @param name - Set name. e.g.: "MatrixBindings". <br/>|
|roamingSettings| [RoamingSettings object](office.context.roamingsettings.md)|Gets an object that represents the saved custom settings of the add-in.|
|touchEnabled|bool|Gets whether the add-in is running in an Office host application that is touch enabled.|
|ui|[Ui object](office.context.ui.md)|Provides objects and methods that you can use to create and manipulate UI components, such as dialog boxes.|
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
