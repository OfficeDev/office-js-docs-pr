
# ProjectDocument.addHandlerAsync method
Asynchronously adds an event handler for a change event in a [ProjectDocument](../../reference/shared/projectdocument.projectdocument.md) object.

|||
|:-----|:-----|
|**Hosts:**|Project|
|**Available in [Requirement set](../../docs/overview/specify-office-hosts-and-api-requirements.md)**|Selection|
|**Added in**|1.0|

```
Office.context.document.addHandlerAsync(eventType, handler[, options][, callback]);
```


## Parameters



|**Name**|**Type**|**Description**|
|:-----|:-----|:-----|
| _eventType_|[EventType](../../reference/shared/eventtype-enumeration.md)|The type of event to add, as an [EventType](../../reference/shared/eventtype-enumeration.md) constant or its corresponding text value. Required.The following table shows valid  _eventType_ arguments for a [ProjectDocument](../../reference/shared/projectdocument.projectdocument.md) object. <table><tr><td>**Enumeration**</td><td>**Text value**</td></tr><tr><td>[Office.EventType.ResourceSelectionChanged](../../reference/shared/projectdocument.resourceselectionchanged.event.md)</td><td>resourceSelectionChanged</td></tr><tr><td>[Office.EventType.TaskSelectionChanged](../../reference/shared/projectdocument.taskselectionchanged.event.md)</td><td>taskSelectionChanged</td></tr><tr><td>[Office.EventType.ViewSelectionChanged](../../reference/shared/projectdocument.viewselectionchanged.event.md)</td><td>viewSelectionChanged</td></tr></table>|
| _handler_|**function**|The name of the event handler. Required.|
| _options_|**object**|Specifies any of the following [optional parameters](../../docs/develop/asynchronous-programming-in-office-add-ins.md#passing-optional-parameters-to-asynchronous-methods).|
| _asyncContext_|**array**,  **boolean**,  **null**,  **number**,  **object**, **string**, or  **undefined**|A user-defined item of any type that is returned in the  **AsyncResult** object without being altered.|
| _callback_|**object**|A function that is invoked when the callback returns, whose only parameter is of type  **AsyncResult**.|

## Callback Value

When the  _callback_ function executes, it receives an [AsyncResult](../../reference/shared/asyncresult.md) object that you can access from the parameter in the callback function.

For the  **addHandlerAsync** method, the returned [AsyncResult](../../reference/shared/asyncresult.md) object contains the following properties:


****


|**Name**|**Description**|
|:-----|:-----|
|[asyncContext](../../reference/shared/asyncresult.asynccontext.md)|The data passed in the optional  _asyncContext_ parameter, if the parameter was used.|
|[error](../../reference/shared/asyncresult.error.md)|Information about the error, if the  **status** property equals **failed**.|
|[status](../../reference/shared/asyncresult.status.md)|The  **succeeded** or **failed** status of the asynchronous call.|
|[value](../../reference/shared/asyncresult.value.md)|**addHandlerAsync** always returns **undefined**.|

## Example

The following code example uses  **addHandlerAsync** to add an event handler for the [ViewSelectionChanged](../../reference/shared/projectdocument.viewselectionchanged.event.md) event.

When the active view changes, the handler checks the view type. It enables a button if the view is a resource view and disables the button if it isn't a resource view. Choosing the button gets the GUID of the selected resource and displays it in the add-in.

The example assumes your add-in has a reference to the jQuery library and that the following page controls are defined in the content div in the page body.




```HTML
<input id="get-info" type="button" value="Get info" disabled="disabled" /><br />
<span id="message"></span>
```




```js
(function () {
    "use strict";

    // The initialize function must be run each time a new page is loaded.
    Office.initialize = function (reason) {
        $(document).ready(function () {

            // After the DOM is loaded, add-in-specific code can run.
            // Add a ViewSelectionChanged event handler.
            Office.context.document.addHandlerAsync(
                Office.EventType.ViewSelectionChanged,
                getActiveView);
            $('#get-info').click(getResourceGuid);

            // This example calls the handler on page load to get the active view
            // of the default page.
            getActiveView();
        });
    };

    // Activate the button based on the active view type of the document.
    // This is the ViewSelectionChanged event handler.
    function getActiveView() {
        Office.context.document.getSelectedViewAsync(
            function (result) {
                if (result.status === Office.AsyncResultStatus.Failed) {
                    onError(result.error);
                }
                else {
                    var viewType = result.value.viewType;
                    if (viewType == 6 ||   // ResourceForm
                        viewType == 7 ||   // ResourceSheet
                        viewType == 8 ||   // ResourceGraph
                        viewType == 15) {  // ResourceUsage
                        $('#get-info').removeAttr('disabled');
                    }
                    else {
                        $('#get-info').attr('disabled', 'disabled');
                    }
                    var output = String.format(
                        'View name: {0}<br/>View type: {1}',
                        result.value.viewName, viewType);
                    $('#message').html(output);
                }
            }
        );
    }

    // Get the GUID of the currently selected resource and display it in the add-in.
    function getResourceGuid() {
        Office.context.document.getSelectedResourceAsync(
            function (result) {
                if (result.status === Office.AsyncResultStatus.Failed) {
                    onError(result.error);
                }
                else {
                    $('#message').html('Resource GUID: ' + result.value);
                }
            }
        );
    }

    function onError(error) {
        $('#message').html(error.name + ' ' + error.code + ': ' + error.message);
    }
})();
```

For a complete code sample that shows how to use a [TaskSelectionChanged](../../reference/shared/projectdocument.taskselectionchanged.event.md) event handler in a Project add-in, see [Create your first task pane add-in for Project by using a text editor](../../docs/project/create-your-first-task-pane-add-in-for-project-by-using-a-text-editor.md).


## Support details


A capital Y in the following matrix indicates that this method is supported in the corresponding Office host application. An empty cell indicates that the Office host application doesn't support this method.

For more information about Office host application and server requirements, see [Requirements for running Office Add-ins](../../docs/overview/requirements-for-running-office-add-ins.md).


||**Office for Windows desktop**|**Office Online (in browser)**|
|:-----|:-----|:-----|
|**Project**|Y||

|||
|:-----|:-----|
|**Available in requirement sets**||
|**Minimum permission level**|[ReadWriteDocument](../../docs/develop/requesting-permissions-for-api-use-in-content-and-task-pane-add-ins.md)|
|**Add-in types**|Task pane|
|**Library**|Office.js|
|**Namespace**|Office|

## Support history



****


|**Version**|**Changes**|
|:-----|:-----|
|1.0|Introduced|

## See also



#### Other resources


[TaskSelectionChanged event](../../reference/shared/projectdocument.taskselectionchanged.event.md)

[removeHandlerAsync method](../../reference/shared/projectdocument.addhandlerasync.md)

[ProjectDocument object](../../reference/shared/projectdocument.projectdocument.md)
