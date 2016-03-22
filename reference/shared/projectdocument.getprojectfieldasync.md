
# ProjectDocument.getProjectFieldAsync method
Asynchronously gets the value of the specified field in the active project.

|||
|:-----|:-----|
|**Hosts:**|Project|
|**Available in [Requirement set](../../docs/overview/specify-office-hosts-and-api-requirements.md)**|Selection|
|**Added in**|1.0|

```
Office.context.document.getProjectFieldAsync(fieldId[, options][, callback]);
```


## Parameters



|**Name**|**Type**|**Description**|
|:-----|:-----|:-----|:-----|
| _fieldId_|[ProjectProjectFields](../../reference/shared/projectprojectfields-enumeration.md)|The ID of the target field. Required.|
| _options_|**object**|Specifies any of the following [optional parameters](../../docs/develop/asynchronous-programming-in-office-add-ins.md#passing-optional-parameters-to-asynchronous-methods).|
| _asyncContext_|**array**,  **boolean**,  **null**,  **number**,  **object**, **string**, or  **undefined**|A user-defined item of any type that is returned in the  **AsyncResult** object without being altered.|
| _callback_|**object**|A function that is invoked when the callback returns, whose only parameter is of type  **AsyncResult**.|

## Callback Value

When the  _callback_ function executes, it receives an [AsyncResult](../../reference/shared/asyncresult.md) object that you can access from the parameter in the callback function.

For the  **getProjectFieldAsync** method, the returned [AsyncResult](../../reference/shared/asyncresult.md) object contains the following properties.


****


|**Name**|**Description**|
|:-----|:-----|
|[asyncContext](../../reference/shared/asyncresult.asynccontext.md)|The data passed in the optional  _asyncContext_ parameter, if the parameter was used.|
|[error](../../reference/shared/asyncresult.error.md)|Information about the error, if the  **status** property equals **failed**.|
|[status](../../reference/shared/asyncresult.status.md)|The  **succeeded** or **failed** status of the asynchronous call.|
|[value](../../reference/shared/asyncresult.value.md)|Contains the  **fieldValue** property, which represents the value of the specified field.|

## Example

The following code example gets the values of three specified fields for the active project, and then displays the values in the add-in.

The example calls  **getProjectFieldAsync** recursively, after the previous call returns successfully. It also tracks the calls to determine when all calls are sent.

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

            // Get information for the active project.
            getProjectInformation();
        });
    };

    // Get the specified fields for the active project.
    function getProjectInformation() {
        var fields =
            [Office.ProjectProjectFields.Start, Office.ProjectProjectFields.Finish, Office.ProjectProjectFields.GUID];
        var fieldValues = ['Start: ', 'Finish: ', 'GUID: '];
        var index = 0; 
        getField();

        // Get each field, and then display the field values in the add-in.
        function getField() {
            if (index == fields.length) {
                var output = '';
                for (var i = 0; i < fieldValues.length; i++) {
                    output += fieldValues[i] + '<br />';
                }
                $('#message').html(output);
            }
            else {
                Office.context.document.getProjectFieldAsync(
                    fields[index],
                    function (result) {

                        // If the call is successful, get the field value and then get the next field.
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



****


|**Version**|**Changes**|
|:-----|:-----|
|1.0|Introduced|

## See also



#### Other resources


[ProjectProjectFields enumeration](../../reference/shared/projectprojectfields-enumeration.md)

[AsyncResult object](../../reference/shared/asyncresult.md)

[ProjectDocument object](../../reference/shared/projectdocument.projectdocument.md)
