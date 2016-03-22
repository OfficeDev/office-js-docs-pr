

# ProjectDocument.setResourceFieldAsync method
Asynchronously sets the value of the specified field for the specified resource.
 **Important:** This API works only in Project 2016 on Windows desktop.

|||
|:-----|:-----|
|**Hosts:**|Project|
|**Available in [Requirement set](../../docs/overview/specify-office-hosts-and-api-requirements.md)**|Selection|
|**Added in**|1.1|

```js
Office.context.document.setResourceFieldAsync(resourceId, fieldId, fieldValue[, options][, callback]);
```


## Parameters

_resourceId_<br/>
&nbsp;&nbsp;&nbsp;&nbsp;The GUID of the resource. Required.
    
_fieldId_<br/>
&nbsp;&nbsp;&nbsp;&nbsp;The ID of the target field, as a [ProjectResourceFields](../../reference/shared/projectresourcefields-enumeration.md) constant or its corresponding integer value. Required.
    
_fieldValue_<br/>
&nbsp;&nbsp;&nbsp;&nbsp;The value for the target field, as  **string**,  **number**,  **boolean**, or  **object**. Required.
    
_options_<br/>
&nbsp;&nbsp;&nbsp;&nbsp;The following [optional parameter](../../docs/develop/asynchronous-programming-in-office-add-ins.md#passing-optional-parameters-to-asynchronous-methods):

&nbsp;&nbsp;&nbsp;&nbsp;_asyncContext_<br/>
&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;Type: **array, boolean, null, number, object, string,** or **undefined**<br/></br>&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;A user-defined item of any type that is returned in the [AsyncResult](../../reference/shared/asyncresult.md) object without being altered. Optional.</br></br>&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;For example, you can pass the _asyncContext_ argument by using the format `{asyncContext: 'Some text'}` or `{asyncContext: <object>}`.


_callback_<br/>
&nbsp;&nbsp;&nbsp;&nbsp;Type:  **function**

&nbsp;&nbsp;&nbsp;&nbsp;A function that is invoked when the method call returns, where the only parameter is of type [AsyncResult](../../reference/shared/asyncresult.md). Optional.

    

## Callback Value

When the  _callback_ function executes, it receives an [AsyncResult](../../reference/shared/asyncresult.md) object that you can access from the parameter in the callback function.

For the  **setResourceFieldAsync** method, the returned [AsyncResult](../../reference/shared/asyncresult.md) object contains following properties.


|**Name**|**Description**|
|:-----|:-----|
|[asyncContext](../../reference/shared/asyncresult.asynccontext.md)|The data passed in the optional  _asyncContext_ parameter, if the parameter was used.|
|[error](../../reference/shared/asyncresult.error.md)|Information about the error, if the  **status** property equals **failed**.|
|[status](../../reference/shared/asyncresult.status.md)|The  **succeeded** or **failed** status of the asynchronous call.|
|[value](../../reference/shared/asyncresult.value.md)|This method does not return a value.|

## Remarks

First call the [getSelectedResourceAsync](../../reference/shared/projectdocument.getselectedtaskasync.md) or [getResourceByIndexAsync](../../reference/shared/projectdocument.getresourcebyindexasync.md) method to get the resource GUID, and then pass the GUID as the _resourceId_ argument to **setResourceFieldAsync**. Only a single field for a single resource can be updated in each asynchronous call.


## Example

The following code example calls [getSelectedResourceAsync](../../reference/shared/projectdocument.getselectedtaskasync.md) to get the GUID of the resource that's currently selected in a resource view. Then it sets two resource field values by calling **setResourceFieldAsync** recursively.

The [getSelectedTaskAsync](../../reference/shared/projectdocument.getselectedtaskasync.md) method used in the example requires that a task view (for example, Task Usage) is the active view and that a task is selected. See the [addHandlerAsync](../../reference/shared/projectdocument.addhandlerasync.md) method for an example that activates a button based on the active view type.

The example assumes your add-in has a reference to the jQuery library and that the following page controls are defined in the content div in the page body.




```HTML
<input id="set-info" type="button" value="Set info" /><br />
<span id="message"></span>
```




```js
(function () {
    "use strict";

    // The initialize function must be run each time a new page is loaded.
    Office.initialize = function (reason) {
        $(document).ready(function () {

            // After the DOM is loaded, add-in-specific code can run.
            app.initialize();
            $('#set-info').click(setResourceInfo);
        });
    };

    // Get the GUID of the resource, and then get the resource fields.
    function setResourceInfo() {
        getResourceGuid().then(
            function (data) {
                setResourceFields(data);
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

    // Set the specified fields for the selected resource.
    function setResourceFields(resourceGuid) {
        var targetFields = [Office.ProjectResourceFields.StandardRate, Office.ProjectResourceFields.Notes];
        var fieldValues = [.28, 'Notes for the resource.'];

        // Set the field value. If the call is successful, set the next field.
        for (var i = 0; i < targetFields.length; i++) {
            Office.context.document.setResourceFieldAsync(
                resourceGuid,
                targetFields[i],
                fieldValues[i],
                function (result) {
                    if (result.status === Office.AsyncResultStatus.Succeeded) {
                        i++;
                    }
                    else {
                        onError(result.error);
                    }
                }
            );
        }
        $('#message').html('Field values set');
    }

    function onError(error) {
        app.showNotification(error.name + ' ' + error.code + ': ' + error.message);
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
|**Minimum permission level**|[WriteDocument](../../docs/develop/requesting-permissions-for-api-use-in-content-and-task-pane-add-ins.md)|
|**Add-in types**|Task pane|
|**Library**|Office.js|
|**Namespace**|Office|

## Support history

|**Version**|**Changes**|
|:-----|:-----|
|1.1|Introduced|

## See also



#### Other resources


[getSelectedResourceAsync](../../reference/shared/projectdocument.getselectedtaskasync.md)
[getResourceByIndexAsync](../../reference/shared/projectdocument.getresourcebyindexasync.md)
[AsyncResult object](../../reference/shared/asyncresult.md)
[ProjectResourceFields enumeration](../../reference/shared/projectresourcefields-enumeration.md)
[ProjectDocument object](../../reference/shared/projectdocument.projectdocument.md)

