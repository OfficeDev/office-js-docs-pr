
# Document object
An abstract class that represents the document the add-in is interacting with.

|||
|:-----|:-----|
|**Hosts:**|Access, Excel, PowerPoint, Project, Word|
|**Added in**|1.0|
|**Last changed in**|1.1|

```
Office.context.document
```


## Members


**Properties**


|**Name**|**Description**|**Support notes**|
|:-----|:-----|:-----|
|[bindings](../../reference/shared/document.bindings.md)|Gets an object that provides access to the bindings defined in the document.|In 1.1, added support for content add-ins for Access.|
|[customXmlParts](../../reference/shared/document.customxmlparts.md)|Gets an object that represents the custom XML parts in the document.||
|[mode](../../reference/shared/document.mode.md)|Gets the mode the document is in.|In 1.1, added support for content add-ins for Access.|
|[settings](../../reference/shared/document.settings.md)|Gets an object that represents the saved custom settings of the content or task pane add-in for the current document.|In 1.1, added support for content add-ins for Access.|
|[url](../../reference/shared/document.url.md)|Gets the URL of the document that the host application currently has open.|In 1.1, added support for content add-ins for Access.|

**Methods**


|**Name**|**Description**|**Support notes**|
|:-----|:-----|:-----|
|[addHandlerAsync](../../reference/shared/document.addhandlerasync.md)|Adds an event handler for a  **Document** object event.||
|[getActiveViewAsync](../../reference/shared/document.getactiveviewasync.md)|Returns the current view of the presentation.|In 1.1, added to support [add-ins for PowerPoint](../../docs/powerpoint/powerpoint-add-ins.md).|
|[getFileAsync](../../reference/shared/document.getfileasync.md)|Returns the entire document file in slices of up to 4194304 bytes (4MB).|In 1.1, added support getting file as PDF in add-ins for PowerPoint and Word.|
|[getFilePropertiesAsync](../../reference/shared/document.getfilepropertiesasync.md)|Gets file properties of the current document.In this release, can get only the URL of the document.|In 1.1, added to get the document's URL in add-ins for Excel, Word, and PowerPoint.|
|[getSelectedDataAsync](../../reference/shared/document.getselecteddataasync.md)|Reads the data contained in the current selection of the document.|In 1.1, added support for getting the id, title, and index for the selected range of slides in add-ins for PowerPoint.|
|[goToByIdAsync](../../reference/shared/document.gotobyidasync.md)|Goes to the specified object or location in the document.|In 1.1, added to support navigation within the document in add-ins for Excel and PowerPoint.|
|[removeHandlerAsync](../../reference/shared/document.removehandlerasync.md)|Removes an event handler for a  **Document** object event.||
|[setSelectedDataAsync](../../reference/shared/document.setselecteddataasync.md)|Writes data to the current selection in the document.|In 1.1, added support for [setting formatting on the selected table when writing data in add-ins for Excel](../../docs/excel/format-tables-in-add-ins-for-excel.md).|

**Events**


|**Name**|**Description**|**Support notes**||
|:-----|:-----|:-----|:-----|
|[ActiveViewChanged](../../reference/shared/document.activeviewchanged.md)|Occurs when the user changes the current view of the document.|In 1.1, added to support add-ins for PowerPoint.||
|[SelectionChanged](../../reference/shared/document.selectionchanged.event.md)|Occurs when the selection in the document is changed.|||

## Remarks

You don't instantiate the  **Document** object directly in your script. To call members of the **Document** object to interact with the current document or worksheet, use `Office.context.document` in your script.


## Example

The following example uses the  **getSelectedDataAsync** method of the **Document** object to retrieve the user's current selection as text, and then display it in the add-in's page.


```js

// Display the user's current selection.
function showSelection() {
    Office.context.document.getSelectedDataAsync(
        "text",                        // coercionType
        {valueFormat: "unformatted",   // valueFormat
        filterType: "all"},            // filterType
        function (result) {            // callback
            var dataValue; 
            dataValue = result.value;
            write('Selected data is: ' + dataValue);
        });
}
// Function that writes to a div with id='message' on the page.
function write(message){
    document.getElementById('message').innerText += message; 
}
```




## Support details


Support for each API member of the  **Document** object differs across Office host applications. See the "Support details" section of each member's topic for host support information.

For more information about Office host application and server requirements, see [Requirements for running Office Add-ins](../../docs/overview/requirements-for-running-office-add-ins.md).


|||
|:-----|:-----|
|**Added in**|1.0|
|**Last changed in**|1.1|
|**Add-in types**|Content, task pane|
|**Library**|Office.js|
|**Namespace**|Office|
