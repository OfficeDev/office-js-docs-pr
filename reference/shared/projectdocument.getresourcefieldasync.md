
# ProjectDocument.getResourceFieldAsync method
Asynchronously gets the value of the specified field for the specified resource in a resource view.

|||
|:-----|:-----|
|**Hosts:**|Project|
|**Available in [Requirement set](../../docs/overview/specify-office-hosts-and-api-requirements.md)**|Selection|
|**Added in**|1.0|

```
Office.context.document.getResourceFieldAsync(resourceId, fieldId[, options][, callback]);
```


## Parameters



|**Name**|**Type**|**Description**|**Support notes**|
|:-----|:-----|:-----|:-----|
| _resourceId_|**string**|The GUID of the resource. Required.||
| _fieldId_|[ProjectResourceFields](../../reference/shared/projectresourcefields-enumeration.md)|The ID of the target field. Required.||
| _options_|**object**|Specifies any of the following [optional parameters](../../docs/develop/asynchronous-programming-in-office-add-ins.md#passing-optional-parameters-to-asynchronous-methods).||
| _asyncContext_|**array**,  **boolean**,  **null**,  **number**,  **object**, **string**, or  **undefined**|A user-defined item of any type that is returned in the  **AsyncResult** object without being altered.||
| _callback_|**object**|A function that is invoked when the callback returns, whose only parameter is of type  **AsyncResult**.||

## Callback Value

When the  _callback_ function executes, it receives an [AsyncResult](../../reference/shared/asyncresult.md) object that you can access from the parameter in the callback function.

For the  **getResourceFieldAsync** method, the returned [AsyncResult](../../reference/shared/asyncresult.md) object contains the following properties.


****


|**Name**|**Description**|
|:-----|:-----|
|[asyncContext](../../reference/shared/asyncresult.asynccontext.md)|The data passed in the optional  _asyncContext_ parameter, if the parameter was used.|
|[error](../../reference/shared/asyncresult.error.md)|Information about the error, if the  **status** property equals **failed**.|
|[status](../../reference/shared/asyncresult.status.md)|The  **succeeded** or **failed** status of the asynchronous call.|
|[value](../../reference/shared/asyncresult.value.md)|Contains the  **fieldValue** property, which represents the value of the specified field.|

## Remarks

First call the [getSelectedResourceAsync](../../reference/shared/projectdocument.getselectedtaskasync.md) method to get the resource GUID, and then pass it as the _resourceId_ argument to **getResourceFieldAsync**. If the active view is not a resource view (for example a Resource Usage or Resource Sheet view), or if no resource is selected in a resource view, [getSelectedResourceAsync](../../reference/shared/projectdocument.getselectedtaskasync.md) returns a 5001 error (Internal Error). See [addHandlerAsync method](../../reference/shared/projectdocument.addhandlerasync.md) for an example that uses the [ViewSelectionChanged](../../reference/shared/projectdocument.viewselectionchanged.event.md) event and the [getSelectedViewAsync](../../reference/shared/projectdocument.getselectedviewasync.md) method to activate a button based on the active view type.


## Example

The following code example calls [getSelectedResourceAsync](../../reference/shared/projectdocument.getselectedtaskasync.md) to get the GUID of the resource that's currently selected in a resource view. Then it gets three resource field values by calling **getResourceFieldAsync** recursively.

The example assumes your add-in has a reference to the jQuery library and that the following page controls are defined in the content div in the page body.




```HTML
<input id="get-info" type="button" value="Get info" /><br />
<span id="message"></span>
```




```js
(function () {
    "use strict";

    // The initialize function must be run each time a new page is loaded.
    Office.initialize = function (reason) {
        $(document).ready(function () {

            // After the DOM is loaded, add-in-specific code can run.
            $('#get-info').click(getResourceInfo);
        });
    };

    // Get the GUID of the resource and then get the resource fields.
    function getResourceInfo() {
        getResourceGuid().then(
            function (data) {
                getResourceFields(data);
            }
        );
    }

    // Get the GUID of the selected resource.
    function getResourceGuid() {
        var defer = $.Deferred();
        Office.context.document.getSelectedResourceAsync(
            function (result) {
                if (result.status === Office.AsyncResultStatus.Failed) {
                    onError(result.error);
                }
                else {
                    defer.resolve(result.value);
                }
            }
        );
        return defer.promise();
    }

    // Get the specified fields for the selected resource.
    function getResourceFields(resourceGuid) {
        var targetFields =
            [Office.ProjectResourceFields.Name, Office.ProjectResourceFields.Units, Office.ProjectResourceFields.BaseCalendar];
        var fieldValues = ['Name: ', 'Units: ', 'Base calendar: '];
        var index = 0; 
        getField();

        // Get each field, and then display the field values in the add-in.
        function getField() {
            if (index == targetFields.length) {
                var output = '';
                for (var i = 0; i < fieldValues.length; i++) {
                    output += fieldValues[i] + '<br />';
                }
                $('#message').html(output);
            }

            // If the call is successful, get the field value and then get the next field.
            else {
                Office.context.document.getResourceFieldAsync(
                    resourceGuid,
                    targetFields[index],
                    function (result) {
                        if (result.status === Office.AsyncResultStatus.Succeeded) {
                            fieldValues[index] += result.value.fieldValue;
                            getField(index++);
                        }
                        else {
                            onError(result.error);
                        }
                    }
                );
            }
        }
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


[getSelectedResourceAsync method](../../reference/shared/projectdocument.getselectedresourceasync.md)

[ProjectResourceFields enumeration](../../reference/shared/projectresourcefields-enumeration.md)

[AsyncResult object](../../reference/shared/asyncresult.md)

[ProjectDocument object](../../reference/shared/projectdocument.projectdocument.md)
