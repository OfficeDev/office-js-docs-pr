

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
|[addHandlerAsync method](https://dev.office.com/reference/add-ins/shared/projectdocument.addhandlerasync)|Asynchronously adds an event handler for an event in a  **ProjectDocument** object.|
|[getMaxResourceIndexAsync method](https://dev.office.com/reference/add-ins/shared/projectdocument.getmaxresourceindexasync)|Asynchronously gets the maximum index of the collection of resources in the current project.|
|[getMaxTaskIndexAsync method](https://dev.office.com/reference/add-ins/shared/projectdocument.getmaxtaskindexasync)|Asynchronously gets the maximum index of the collection of tasks in the current project.|
|[getProjectFieldAsync method](https://dev.office.com/reference/add-ins/shared/projectdocument.getprojectfieldasync)|Asynchronously gets the value of the specified field in the active project.|
|[getResourceByIndexAsync method](https://dev.office.com/reference/add-ins/shared/projectdocument.getresourcebyindexasync)|Asynchronously gets the GUID of the resource that has the specified index in the resource collection.|
|[getResourceFieldAsync method](https://dev.office.com/reference/add-ins/shared/projectdocument.getresourcefieldasync)|Asynchronously gets the value of the specified field for the specified resource.|
|[getSelectedDataAsync method](https://dev.office.com/reference/add-ins/shared/projectdocument.getselecteddataasync)|Asynchronously gets the data that is contained in the current selection of one or more cells in the Gantt chart.|
|[getSelectedResourceAsync method](https://dev.office.com/reference/add-ins/shared/projectdocument.getselectedresourceasync)|Asynchronously gets the GUID of the selected resource.|
|[getSelectedTaskAsync method](https://dev.office.com/reference/add-ins/shared/projectdocument.getselectedtaskasync)|Asynchronously gets the GUID of the selected task.|
|[getSelectedViewAsync method](https://dev.office.com/reference/add-ins/shared/projectdocument.getselectedviewasync)|Asynchronously gets the view type and name of the active view.|
|[getTaskAsync method](https://dev.office.com/reference/add-ins/shared/projectdocument.gettaskasync)|Asynchronously gets the task name, the resources that are assigned to the task, and the ID of the task in the synchronized SharePoint task list.|
|[getTaskByIndexAsync method](https://dev.office.com/reference/add-ins/shared/projectdocument.gettaskbyindexasync)|Asynchronously gets the GUID of the task that has the specified index in the task collection.|
|[getTaskFieldAsync method](https://dev.office.com/reference/add-ins/shared/projectdocument.gettaskfieldasync)|Asynchronously gets the value of the specified field for the specified task.|
|[getWSSUrlAsync method](https://dev.office.com/reference/add-ins/shared/projectdocument.getwssurlasync)|Asynchronously gets the URL of the synchronized SharePoint task list.|
|[removeHandlerAsync method](https://dev.office.com/reference/add-ins/shared/projectdocument.removehandlerasync)|Asynchronously removes an event handler for an event in a  **ProjectDocument** object.|
|[setResourceFieldAsync method](https://dev.office.com/reference/add-ins/shared/projectdocument.setresourcefieldasync)|Asynchronously sets the value of the specified field for the specified resource.|
|[setTaskFieldAsync method](https://dev.office.com/reference/add-ins/shared/projectdocument.settaskfieldasync)|Asynchronously sets the value of the specified field for the specified task.|

**Events**


|**Name**|**Description**|
|:-----|:-----|
|[ResourceSelectionChanged event](https://dev.office.com/reference/add-ins/shared/projectdocument.resourceselectionchanged.event)|Occurs when the resource selection changes in the active project.|
|[TaskSelectionChanged event](https://dev.office.com/reference/add-ins/shared/projectdocument.taskselectionchanged.event)|Occurs when the task selection changes in the active project.|
|[ViewSelectionChanged event](https://dev.office.com/reference/add-ins/shared/projectdocument.viewselectionchanged.event)|Occurs when the active view changes in the active project.|

## Remarks

Do not directly call or instantiate the  **ProjectDocument** object in your script.


## Example

The following example initializes the add-in and then gets properties of the [Document](https://dev.office.com/reference/add-ins/shared/document) object that are available in the context of a Project document. A Project document is the opened, active project. To access members of the **ProjectDocument** object, use the **Office.context.document** object as shown in the code examples for **ProjectDocument** methods and events.

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
[Document object](https://dev.office.com/reference/add-ins/shared/document)

