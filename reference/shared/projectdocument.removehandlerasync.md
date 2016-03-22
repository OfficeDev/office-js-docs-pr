

# ProjectDocument.removeHandlerAsync method
Asynchronously removes an event handler for the task selection changed event in a [ProjectDocument](../../reference/shared/projectdocument.projectdocument.md) object.

|||
|:-----|:-----|
|**Hosts:**|Project|
|**Available in [Requirement set](../../docs/overview/specify-office-hosts-and-api-requirements.md)**|Selection|
|**Added in**|1.0|

```js
Office.context.document.removeHandlerAsync(eventType[, options][, callback]);
```


## Parameters
|**Name**|**Type**|**Description**|**Support notes**|
|:-----|:-----|:-----|:-----|
|_eventType_|[EventType](../../reference/shared/eventtype-enumeration.md)|The type of event to remove, as an [EventType](../../reference/shared/eventtype-enumeration.md) constant or its corresponding text value. Required.<br/><br/>The following table shows valid eventType arguments for a [ProjectDocument](../../reference/shared/projectdocument.projectdocument.md) object.<br/><br/><table><tr><th>Enumeration</th><th>Text value</th></tr><tr><td><a href="https://msdn.microsoft.com/en-us/library/office/fp179836.aspx">Office.EventType.ResourceSelectionChanged</a></td><td>resourceSelectionChanged</td></tr><tr><td><a href="https://msdn.microsoft.com/en-us/library/office/fp179816.aspx">Office.EventType.TaskSelectionChanged</a></td><td>taskSelectionChanged</td></tr><tr><td><a href="https://msdn.microsoft.com/en-us/library/office/fp179839.aspx">Office.EventType.ViewSelectionChanged</a></td><td>viewSelectionChanged</td></tr></table>||
|_options_|**object**|Specifies any of the following [optional parameters](../../docs/develop/asynchronous-programming-in-office-add-ins.md#passing-optional-parameters-to-asynchronous-methods).||
|_asyncContext_|**array**,  **boolean**,  **null**,  **number**,  **object**, **string**, or  **undefined**|A user-defined item of any type that is returned in the  **AsyncResult** object without being altered.||
|_callback_|**object**|A function that is invoked when the callback returns, whose only parameter is of type  **AsyncResult**.||


## Callback Value

When the  _callback_ function executes, it receives an [AsyncResult](../../reference/shared/asyncresult.md) object that you can access from the parameter in the callback function.

For the  **removeHandlerAsync** method, the returned [AsyncResult](../../reference/shared/asyncresult.md) object contains the following properties.


|**Name**|**Description**|
|:-----|:-----|
|[asyncContext](../../reference/shared/asyncresult.asynccontext.md)|The data passed in the optional  _asyncContext_ parameter, if the parameter was used.|
|[error](../../reference/shared/asyncresult.error.md)|Information about the error, if the  **status** property equals **failed**.|
|[status](../../reference/shared/asyncresult.status.md)|The  **succeeded** or **failed** status of the asynchronous call.|
|[value](../../reference/shared/asyncresult.value.md)|**removeHandlerAsync** always returns **undefined**.|

## Example

The following code example uses [addHandlerAsync](../../reference/shared/projectdocument.addhandlerasync.md) to add an event handler for the [ResourceSelectionChanged](../../reference/shared/projectdocument.resourceselectionchanged.event.md) event and **removeHandlerAsync** to remove the handler.

When a resource is selected in a resource view, the handler displays the resource GUID. When the handler is removed, the GUID is not displayed.

The example assumes your add-in has a reference to the jQuery library and that the following page control is defined in the content div in the page body.




```HTML
<input id="remove-handler" type="button" value="Remove handler" /><br />
<span id="message"></span>
```




```js
(function () {
    "use strict";

    // The initialize function must be run each time a new page is loaded.
    Office.initialize = function (reason) {
        $(document).ready(function () {

            // After the DOM is loaded, add-in-specific code can run.
            Office.context.document.addHandlerAsync(
                Office.EventType.ResourceSelectionChanged,
                getResourceGuid);
            $('#remove-handler').click(removeEventHandler);
        });
    };

    // Remove the event handler.
    function removeEventHandler() {
        Office.context.document.removeHandlerAsync(
            Office.EventType.ResourceSelectionChanged,
            {handler:getResourceGuid,
            asyncContext:'The handler is removed.'},
            function (result) {
                if (result.status === Office.AsyncResultStatus.Failed) {
                    onError(result.error);
                }
                else {
                    $('#remove-handler').attr('disabled', 'disabled');
                    $('#message').html(result.asyncContext);
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


## Support details


A capital Y in the following matrix indicates that this method is supported in the corresponding Office host application. An empty cell indicates that the Office host application doesn't support this method.

For more information about Office host application and server requirements, see [Requirements for running Office Add-ins](../../docs/overview/requirements-for-running-office-add-ins.md).


||**Office for Windows desktop**|**Office Online (in browser)**|
|:-----|:-----|:-----|
|**Project**|Y||

|||
|:-----|:-----|
|**Available in requirement sets**|Selection|
|**Minimum permission level**|[ReadWriteDocument](../../docs/develop/requesting-permissions-for-api-use-in-content-and-task-pane-add-ins.md)|
|**Add-in types**|Task pane|
|**Library**|Office.js|
|**Namespace**|Office|

## Support history

|**Version**|**Changes**|
|:-----|:-----|
|1.0|Introduced|

## See also



#### Other resources


[addHandlerAsync method](../../reference/shared/projectdocument.addhandlerasync.md)
[EventType enumeration](../../reference/shared/eventtype-enumeration.md)
[ProjectDocument object](../../reference/shared/projectdocument.projectdocument.md)

