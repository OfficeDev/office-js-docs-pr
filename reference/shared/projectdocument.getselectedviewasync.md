

# ProjectDocument.getSelectedViewAsync method
Asynchronously gets the type and name of the active view in the document.

|||
|:-----|:-----|
|**Hosts:**|Project|
|**Available in [Requirement set](../../docs/overview/specify-office-hosts-and-api-requirements.md)**|Selection|
|**Added in**|1.0|

```js
Office.context.document.getSelectedViewAsync([options,] [callback]);
```


## Parameters



|**Name**|**Type**|**Description**|**Support notes**|
|:-----|:-----|:-----|:-----|
| _options_|**object**|Specifies any of the following [optional parameters](../../docs/develop/asynchronous-programming-in-office-add-ins.md#passing-optional-parameters-to-asynchronous-methods).||
| _asyncContext_|**array**,  **boolean**,  **null**,  **number**,  **object**, **string**, or  **undefined**|A user-defined item of any type that is returned in the  **AsyncResult** object without being altered.||
| _callback_|**object**|A function that is invoked when the callback returns, whose only parameter is of type  **AsyncResult**.||

## Callback Value

When the  _callback_ function executes, it receives an [AsyncResult](https://dev.office.com/reference/add-ins/shared/asyncresult) object that you can access from the parameter in the callback function.

For the  **getSelectedViewAsync** method, the returned [AsyncResult](https://dev.office.com/reference/add-ins/shared/asyncresult) object contains the following properties.


****


|**Name**|**Description**|
|:-----|:-----|
|[asyncContext](https://dev.office.com/reference/add-ins/shared/asyncresult.asynccontext)|The data passed in the optional  _asyncContext_ parameter, if the parameter was used.|
|[error](https://dev.office.com/reference/add-ins/shared/asyncresult.error)|Information about the error, if the  **status** property equals **failed**.|
|[status](https://dev.office.com/reference/add-ins/shared/asyncresult.status)|The  **succeeded** or **failed** status of the asynchronous call.|
|[value](https://dev.office.com/reference/add-ins/shared/asyncresult.value)|Contains the following properties:<br/><br/><div>* **viewName** - The name of the view, as a [ProjectViewTypes](https://dev.office.com/reference/add-ins/shared/projectviewtypes-enumeration) constant.<br/>* **viewType** - The type of view, as the integer value of a [ProjectViewTypes](https://dev.office.com/reference/add-ins/shared/projectviewtypes-enumeration) constant.</div>|

## Example

The following code example calls adds a [ViewSelectionChanged](https://dev.office.com/reference/add-ins/shared/projectdocument.viewselectionchanged.event) event handler that calls **getSelectedViewAsync** to get the name and type of the active view in the document.

The example assumes your add-in has a reference to the jQuery library and that the following page control is defined in the content div in the page body.




```HTML
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
                Office.EventType.ViewSelectionChanged,
                getActiveView);
            getActiveView();
        });
    };

    // Get the active view's name and type.
    function getActiveView() {
        Office.context.document.getSelectedViewAsync(
            function (result) {
                if (result.status === Office.AsyncResultStatus.Failed) {
                    onError(result.error);
                }
                else {
                    var output = String.format(
                        'View name: {0}<br/>View type: {1}',
                        result.value.viewName, viewType);
                    $('#message').html(output);
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
|**Minimum permission level**|[ReadDocument](../../docs/develop/requesting-permissions-for-api-use-in-content-and-task-pane-add-ins.md)|
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


[ProjectViewTypes enumeration](https://dev.office.com/reference/add-ins/shared/projectviewtypes-enumeration)
[AsyncResult object](https://dev.office.com/reference/add-ins/shared/asyncresult)
[ViewSelectionChanged event](https://dev.office.com/reference/add-ins/shared/projectdocument.viewselectionchanged.event)
[ProjectDocument object](https://dev.office.com/reference/add-ins/shared/projectdocument.projectdocument)
