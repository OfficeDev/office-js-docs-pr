

# ProjectDocument object
An abstract class that represents the project document (the active project) with which the Office Add-in interacts.

|||
|:-----|:-----|
|**Hosts:**|Project|
|**Added in**|1.0|

```js
Office.context.document
```


## Members


**Methods**


|**Name**|**Description**|
|:-----|:-----|
|[addHandlerAsync method](../../reference/shared/projectdocument.addhandlerasync.md)|Asynchronously adds an event handler for an event in a  **ProjectDocument** object.|
|[getMaxResourceIndexAsync method](../../reference/shared/projectdocument.getmaxresourceindexasync.md)|Asynchronously gets the maximum index of the collection of resources in the current project.|
|[getMaxTaskIndexAsync method](../../reference/shared/projectdocument.getmaxtaskindexasync.md)|Asynchronously gets the maximum index of the collection of tasks in the current project.|
|[getProjectFieldAsync method](../../reference/shared/projectdocument.getprojectfieldasync.md)|Asynchronously gets the value of the specified field in the active project.|
|[getResourceByIndexAsync method](../../reference/shared/projectdocument.getresourcebyindexasync.md)|Asynchronously gets the GUID of the resource that has the specified index in the resource collection.|
|[getResourceFieldAsync method](../../reference/shared/projectdocument.getresourcefieldasync.md)|Asynchronously gets the value of the specified field for the specified resource.|
|[getSelectedDataAsync method](../../reference/shared/projectdocument.getselecteddataasync.md)|Asynchronously gets the data that is contained in the current selection of one or more cells in the Gantt chart.|
|[getSelectedResourceAsync method](../../reference/shared/projectdocument.getselectedresourceasync.md)|Asynchronously gets the GUID of the selected resource.|
|[getSelectedTaskAsync method](../../reference/shared/projectdocument.getselectedtaskasync.md)|Asynchronously gets the GUID of the selected task.|
|[getSelectedViewAsync method](../../reference/shared/projectdocument.getselectedviewasync.md)|Asynchronously gets the view type and name of the active view.|
|[getTaskAsync method](../../reference/shared/projectdocument.gettaskasync.md)|Asynchronously gets the task name, the resources that are assigned to the task, and the ID of the task in the synchronized SharePoint task list.|
|[getTaskByIndexAsync method](../../reference/shared/projectdocument.gettaskbyindexasync.md)|Asynchronously gets the GUID of the task that has the specified index in the task collection.|
|[getTaskFieldAsync method](../../reference/shared/projectdocument.gettaskfieldasync.md)|Asynchronously gets the value of the specified field for the specified task.|
|[getWSSUrlAsync method](../../reference/shared/projectdocument.getwssurlasync.md)|Asynchronously gets the URL of the synchronized SharePoint task list.|
|[removeHandlerAsync method](../../reference/shared/projectdocument.removehandlerasync.md)|Asynchronously removes an event handler for an event in a  **ProjectDocument** object.|
|[setResourceFieldAsync method](../../reference/shared/projectdocument.setresourcefieldasync.md)|Asynchronously sets the value of the specified field for the specified resource.|
|[setTaskFieldAsync method](../../reference/shared/projectdocument.settaskfieldasync.md)|Asynchronously sets the value of the specified field for the specified task.|

**Events**


|**Name**|**Description**|
|:-----|:-----|
|[ResourceSelectionChanged event](../../reference/shared/projectdocument.resourceselectionchanged.event.md)|Occurs when the resource selection changes in the active project.|
|[TaskSelectionChanged event](../../reference/shared/projectdocument.taskselectionchanged.event.md)|Occurs when the task selection changes in the active project.|
|[ViewSelectionChanged event](../../reference/shared/projectdocument.viewselectionchanged.event.md)|Occurs when the active view changes in the active project.|

## Remarks

Do not directly call or instantiate the  **ProjectDocument** object in your script.


## Example

The following example initializes the add-in and then gets properties of the [Document](../../reference/shared/document.md) object that are available in the context of a Project document. A Project document is the opened, active project. To access members of the **ProjectDocument** object, use the **Office.context.document** object as shown in the code examples for **ProjectDocument** methods and events.

The example assumes your add-in has a reference to the jQuery library and that the following page control is defined in the content div in the page body:




```HTML
<span id="message"></span>
```




```js
(function () {
    "use strict";

    // The initialize function must be run each time a new page is loaded.
    Office.initialize = function (reason) {
        $(document).ready(function () {

            // Get information about the document.
            showDocumentProperties();
        });
    };

    // Get the document mode and the URL of the active project.
    function showDocumentProperties() {
        var output = String.format(
            'The document mode is {0}.<br/>The URL of the active project is {1}.',
            Office.context.document.mode,
            Office.context.document.url);
        $('#message').html(output);
    }
})();
```


## Support details


A capital Y in the following matrix indicates that this object is supported in the corresponding Office host application. An empty cell indicates that the Office host application doesn't support this object.

For more information about Office host application and server requirements, see [Requirements for running Office Add-ins](../../docs/overview/requirements-for-running-office-add-ins.md).


||**Office for Windows desktop**|**Office Online (in browser)**|
|:-----|:-----|:-----|
|**Project**|Y||

|||
|:-----|:-----|
|**Add-in types**|Task pane|
|**Library**|Office.js|
|**Namespace**|Office|

## Support history


|**Version**|**Changes**|
|:-----|:-----|
|1.0|Introduced|

## See also



#### Other resources


[Task pane add-ins for Project](../../docs/project/project-add-ins.md)
[Document object](../../reference/shared/document.md)

