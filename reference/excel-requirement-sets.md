# Excel add-in requirement sets

Requirement sets are named groups of API members. Office Add-ins use requirement sets specified in the manifest or using a runtime check to determine if an Office host supports APIs needed by the add-in. For more information, see [Specify Office hosts and API requirements](../docs/overview/specify-office-hosts-and-api-requirements.md).

To get a broad view of where add-ins are supported by Office host, see the [Office Add-in host and platform availability](https://dev.office.com/add-in-availability) page.


Excel Add-ins run across multiple versions of Office including Office for Windows Desktop, Office Online, Office for the Mac, and Office for the iPad.


The following table lists the names of requirement sets, the methods in each set, the Office host applications that support that requirement set, and the version number of the API.

For information about requirement sets for Outlook, see [Understanding Outlook API requirement sets](./outlook/tutorial-api-requirement-sets.md).

|  Set version  |  Office Online  |  Office 2016 for Windows  |  Office 2016 for iPad  |  Office 2016 for Mac  |
|:-----|-----|:-----|:-----|:-----|
| ExcelApi 1.3  | 7403.1000 | 7403.1000| 7403.1000 | 7403.1000|
| ExcelApi 1.2  | 7403.1000 | 7403.1000| 7403.1000 | 7403.1000|
| ExcelAPI 1.1  | 7403.1000 | 7403.1000| 7403.1000 | 7403.1000|

## Methods that aren't part of a requirement set


The following methods in the JavaScript API for Office aren't part of a requirement set. If your add-in requires any of these methods, use the  **Methods** and **Method** elements in the add-in's manifest to declare that they are required, or perform the runtime check using an if statement. For more information, see [Specify Office hosts and API requirements](../docs/overview/specify-office-hosts-and-api-requirements.md).



|**Method name**|**Office host support**|
|:-----|:-----|
|Bindings.addFromPromptAsync|Access web apps, Excel, and Excel Online|
|Document.getFilePropertiesAsync|Excel, Excel Online, Word, and PowerPoint|
|Document.getProjectFieldAsync|Project Standard 2013 and Project Professional 2013|
|Document.getResourceFieldAsync|Project Standard 2013 and Project Professional 2013|
|Document.getSelectedResourceAsync|Project Standard 2013 and Project Professional 2013|
|Document.getSelectedTaskAsync|Project Standard 2013 and Project Professional 2013|
|Document.getSelectedViewAsync|PowerPoint and PowerPoint Online|
|Document.getTaskAsync|Project Standard 2013 and Project Professional 2013|
|Document.getTaskFieldAsync|Project Standard 2013 and Project Professional 2013|
|Document.goToByIdAsync|Excel, Excel Online, Word, and PowerPoint|
|Settings.addHandlerAsync|Access web apps, Excel, Excel Online, Word, and PowerPoint|
|Settings.refreshAsync|Access web apps, Excel, Excel Online, Word, PowerPoint, and PowerPoint Online|
|Settings.removeHandlerAsync|Access web apps, Excel, Excel Online, Word, and PowerPoint|
|TableBinding.clearFormatsAsync|Excel, Excel Online|
|TableBinding.setFormatsAsync|Excel, Excel Online|
|TableBinding.setTableOptionsAsync|Excel, Excel Online|

## Additional resources



- [Specify Office hosts and API requirements](../docs/overview/specify-office-hosts-and-api-requirements.md)
