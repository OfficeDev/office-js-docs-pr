
# Shared API


The Shared API subset of the JavaScript API for Office includes objects, methods, properties and events that you can use in all three types of Office Add-ins: content, task pane, and Outlook add-ins.


## Objects





|**Object**|**Description**|
|:-----|:-----|
|[AsyncResult](../../reference/shared/asyncresult.md)|An object which encapsulates the result of an asynchronous request, including status and error information if the request failed.|
|[Context](../../reference/shared/asyncresult.context.md)|Represents the runtime environment of the add-in and provides access to key objects of the API.|
|[Error](../../reference/shared/error.md)|Provides specific information about an error that occurred during an asynchronous data operation.|
|[Office](../../reference/shared/office.md)|Represents an instance of an add-in, which provides access to the top-level objects of the API.|


|**Member**|**Description**|
|:-----|:-----|
|[event.completed](../../reference/shared/event.completed.md)|The callback that the add-in invokes to let Outlook know that the operation is done.|
|[event.source.id](../../reference/shared/event.source.id.md)|Gets the id of the control that triggered calling this function.|

## Supported host applications


|||
|:-----|:-----|
|**Supported hosts**|
<ul><li><p>Access</p></li><li><p>Excel</p></li><li><p>Outlook</p></li><li><p>PowerPoint</p></li><li><p>Project</p></li><li><p>Word</p></li></ul>|
|**Library**|Office.js|
|**Namespace**|Office|

## Additional resources



- [JavaScript API for Office](../../reference/javascript-api-for-office.md)
    
