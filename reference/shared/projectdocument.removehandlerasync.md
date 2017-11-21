

# ProjectDocument.removeHandlerAsync method
Asynchronously removes an event handler for the task selection changed event in a [ProjectDocument](https://dev.office.com/reference/add-ins/shared/projectdocument.projectdocument) object.

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
|_eventType_|[EventType](https://dev.office.com/reference/add-ins/shared/eventtype-enumeration)|The type of event to remove, as an [EventType](https://dev.office.com/reference/add-ins/shared/eventtype-enumeration) constant or its corresponding text value. Required. See [eventType value](#eventtype-value). ||
|_options_|**object**|Specifies any of the following [optional parameters](../../docs/develop/asynchronous-programming-in-office-add-ins.md#passing-optional-parameters-to-asynchronous-methods).||
|_asyncContext_|**array**, **boolean**, **null**, **number**, **object**, **string**, or **undefined**|A user-defined item of any type that is returned in the  **AsyncResult** object without being altered.||
|_callback_|**object**|A function that is invoked when the callback returns, whose only parameter is of type **AsyncResult**.||

## eventType value

The following table shows valid eventType arguments for a [ProjectDocument](https://dev.office.com/reference/add-ins/shared/projectdocument.projectdocument) object.

|**Enumeration**|**Text value**|
|:-----|:-----|
|[Office.EventType.ResourceSelectionChanged](https://msdn.microsoft.com/en-us/library/office/fp179836.aspx)|resourceSelectionChanged|
|[Office.EventType.TaskSelectionChanged](https://msdn.microsoft.com/en-us/library/office/fp179816.aspx)|taskSelectionChanged|
|[Office.EventType.ViewSelectionChanged](https://msdn.microsoft.com/en-us/library/office/fp179839.aspx)|viewSelectionChanged|


## Callback value

When the _callback_ function executes, it receives an [AsyncResult](https://dev.office.com/reference/add-ins/shared/asyncresult) object that you can access from the parameter in the callback function.

For the  **removeHandlerAsync** method, the returned [AsyncResult](https://dev.office.com/reference/add-ins/shared/asyncresult) object contains the following properties.

|**Name**|**Description**|
|:-----|:-----|
|[asyncContext](https://dev.office.com/reference/add-ins/shared/asyncresult.asynccontext)|The data passed in the optional  _asyncContext_ parameter, if the parameter was used.|
|[error](https://dev.office.com/reference/add-ins/shared/asyncresult.error)|Information about the error, if the  **status** property equals **failed**.|
|[status](https://dev.office.com/reference/add-ins/shared/asyncresult.status)|The  **succeeded** or **failed** status of the asynchronous call.|
|[value](https://dev.office.com/reference/add-ins/shared/asyncresult.value)|**removeHandlerAsync** always returns **undefined**.|

## Example

The following code example uses [addHandlerAsync](https://dev.office.com/reference/add-ins/shared/projectdocument.addhandlerasync) to add an event handler for the [ResourceSelectionChanged](https://dev.office.com/reference/add-ins/shared/projectdocument.resourceselectionchanged.event) event and **removeHandlerAsync** to remove the handler.

When a resource is selected in a resource view, the handler displays the resource GUID. When the handler is removed, the GUID is not displayed.

The example assumes that your add-in has a reference to the jQuery library and that the following page control is defined in the content div in the page body.

```HTML
<input id="remove-handler" type="button" value="Remove handler" /><br />
<span id="message"></span>
```

<br/>

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

<br/>

## Support details

A capital Y in the following matrix indicates that this method is supported in the corresponding Office host application. An empty cell indicates that the Office host application doesn't support this method.

For more information about Office host application and server requirements, see [Requirements for running Office Add-ins](../../docs/overview/requirements-for-running-office-add-ins.md).


||**Office for Windows desktop**|**Office Online (in browser)**|
|:-----|:---:|:-----|
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

- [addHandlerAsync method](https://dev.office.com/reference/add-ins/shared/projectdocument.addhandlerasync)

- [EventType enumeration](https://dev.office.com/reference/add-ins/shared/eventtype-enumeration)

- [ProjectDocument object](https://dev.office.com/reference/add-ins/shared/projectdocument.projectdocument)

