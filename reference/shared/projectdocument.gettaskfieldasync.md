

# ProjectDocument.getTaskFieldAsync method
Asynchronously gets the value of the specified field for the specified task.

|||
|:-----|:-----|
|**Hosts:**|Project|
|**Available in [Requirement set](../../docs/overview/specify-office-hosts-and-api-requirements.md)**|Selection|
|**Added in**|1.0|

```js
Office.context.document.getTaskFieldAsync(taskId, fieldId[, options][, callback]);
```


## Parameters



|**Name**|**Type**|**Description**|**Support notes**|
|:-----|:-----|:-----|:-----|
| _taskId_|**string**|The GUID of the task. Required.||
| _fieldId_|[ProjectTaskFields](../../reference/shared/projecttaskfields-enumeration.md)|The ID of the target field. Required.||
| _options_|**object**|Specifies any of the following [optional parameters](../../docs/develop/asynchronous-programming-in-office-add-ins.md#passing-optional-parameters-to-asynchronous-methods).||
| _asyncContext_|**array**,  **boolean**,  **null**,  **number**,  **object**, **string**, or  **undefined**|A user-defined item of any type that is returned in the  **AsyncResult** object without being altered.||
| _callback_|**object**|A function that is invoked when the callback returns, whose only parameter is of type  **AsyncResult**.||

## Callback Value

When the  _callback_ function executes, it receives an [AsyncResult](../../reference/shared/asyncresult.md) object that you can access from the parameter in the callback function.

For the  **getTaskFieldAsync** method, the returned [AsyncResult](../../reference/shared/asyncresult.md) object contains following properties.



|**Name**|**Description**|
|:-----|:-----|
|[asyncContext](../../reference/shared/asyncresult.asynccontext.md)|The data passed in the optional  _asyncContext_ parameter, if the parameter was used.|
|[error](../../reference/shared/asyncresult.error.md)|Information about the error, if the  **status** property equals **failed**.|
|[status](../../reference/shared/asyncresult.status.md)|The  **succeeded** or **failed** status of the asynchronous call.|
|[value](../../reference/shared/asyncresult.value.md)|Contains the  **fieldValue** property, which represents the value of the specified field.|

## Remarks

First call the [getSelectedTaskAsync](../../reference/shared/projectdocument.getselectedtaskasync.md) method to get the task GUID, and then pass it as the _taskId_ argument to **getTaskFieldAsync**. If the active view is not a task view (for example a Gantt Chart or Task Usage view), or if no task is selected in a task view, [getSelectedTaskAsync](../../reference/shared/projectdocument.getselectedtaskasync.md) returns a 5001 error (Internal Error). See [addHandlerAsync method](../../reference/shared/projectdocument.addhandlerasync.md) for an example that uses the [ViewSelectionChanged](../../reference/shared/projectdocument.viewselectionchanged.event.md) event and the [getSelectedViewAsync](../../reference/shared/projectdocument.getselectedviewasync.md) method to activate a button based on the active view type.


## Example

The following code example calls [getSelectedTaskAsync](../../reference/shared/projectdocument.getselectedtaskasync.md) to get the GUID of the task that's currently selected in a task view. Then it gets two task field values by calling **getTaskFieldAsync** recursively.

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
            $('#get-info').click(getTaskInfo);
        });
    };

    // Get the GUID of the task, and then get the task fields.
    function getTaskInfo() {
        getTaskGuid().then(
            function (data) {
                getTaskFields(data);
            }
        );
    }

    // Get the GUID of the selected task.
    function getTaskGuid() {
        var defer = $.Deferred();
        Office.context.document.getSelectedTaskAsync(
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

    // Get the specified fields for the selected task.
    function getTaskFields(taskGuid) {
        var output = '';
        var targetFields = [Office.ProjectTaskFields.Priority, Office.ProjectTaskFields.PercentComplete];
        var fieldValues = ['Priority: ', '% Complete: '];
        var index = 0;
        getField();

        // Get each field, and then display the field values in the add-in.
        function getField() {
            if (index == targetFields.length) {
                for (var i = 0; i < fieldValues.length; i++) {
                    output += fieldValues[i] + '<br />';
                }
                $('#message').html(output);
            }

            // Get the field value. If the call is successful, then get the next field.
            else {
                Office.context.document.getTaskFieldAsync(
                    taskGuid,
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
|**Available in requirement sets**||
|**Minimum permission level**|[ReadDocument](../../docs/develop/requesting-permissions-for-api-use-in-content-and-task-pane-add-ins.md)|
|**Add-in types**|Task pane|
|**Library**|Office.js|
|**Namespace**|Office|

## Support history



|**Version**|**Changes**|
|:-----|:-----|
|1.0|Introduced|

## See also



#### Other resources


[getSelectedTaskAsync method](../../reference/shared/projectdocument.getselectedresourceasync.md)
[AsyncResult object](../../reference/shared/asyncresult.md)
[ProjectTaskFields enumeration](../../reference/shared/projecttaskfields-enumeration.md)
[ProjectDocument object](../../reference/shared/projectdocument.projectdocument.md)
