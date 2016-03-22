

# ProjectDocument.getTaskByIndexAsync method
Asynchronously gets the GUID of the task that has the specified index in the task collection.

**Important:** This API works only in Project 2016 on Windows desktop.

|||
|:-----|:-----|
|**Hosts:**|Project|
|**Available in [Requirement set](../../docs/overview/specify-office-hosts-and-api-requirements.md)**|Selection|
|**Added in**|1.1|

```js
Office.context.document.getTaskByIndexAsync(taskIndex[, options][, callback]);
```


## Parameters

_taskIndex_<br/>
&nbsp;&nbsp;&nbsp;&nbsp;Type:  **number**

&nbsp;&nbsp;&nbsp;&nbsp;The index of the task in the collection of tasks for the project. Required.

    
_options_<br/>
&nbsp;&nbsp;&nbsp;&nbsp;The following [optional parameter](../../docs/develop/asynchronous-programming-in-office-add-ins.md#passing-optional-parameters-to-asynchronous-methods):


&nbsp;&nbsp;&nbsp;&nbsp;_asyncContext_<br/>
&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;Type: **array, boolean, null, number, object, string,** or **undefined**<br/></br>&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;A user-defined item of any type that is returned in the [AsyncResult](../../reference/shared/asyncresult.md) object without being altered. Optional.</br></br>&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;For example, you can pass the _asyncContext_ argument by using the format `{asyncContext: 'Some text'}` or `{asyncContext: <object>}`.

_callback_<br/>
&nbsp;&nbsp;&nbsp;&nbsp;Type:  **function**

&nbsp;&nbsp;&nbsp;&nbsp;A function that is invoked when the method call returns, where the only parameter is of type [AsyncResult](../../reference/shared/asyncresult.md). Optional.


## Callback Value

When the  _callback_ function executes, it receives an [AsyncResult](../../reference/shared/asyncresult.md) object that you can access from the parameter in the callback function.

For the  **getTaskByIndexAsync** method, the returned [AsyncResult](../../reference/shared/asyncresult.md) object contains following properties.


|**Name**|**Description**|
|:-----|:-----|
|[asyncContext](../../reference/shared/asyncresult.asynccontext.md)|The data passed in the optional  _asyncContext_ parameter, if the parameter was used.|
|[error](../../reference/shared/asyncresult.error.md)|Information about the error, if the  **status** property equals **failed**.|
|[status](../../reference/shared/asyncresult.status.md)|The  **succeeded** or **failed** status of the asynchronous call.|
|[value](../../reference/shared/asyncresult.value.md)|The GUID of the task as a  **string**.|

## Remarks

To get the maximum index of the collection of tasks for the project, use the [getMaxTaskIndexAsync](../../reference/shared/projectdocument.getmaxtaskindexasync.md) method. The 0 index task represents the project summary task.


## Example

The following code example calls [getMaxTaskIndexAsync](../../reference/shared/projectdocument.getmaxtaskindexasync.md) to get the maximum index in the project's task collection, and then calls **getTaskByIndexAsync** to get the GUID for each task.

The example assumes that your add-in has a reference to the jQuery library and that the following page controls are defined in the content div in the page body.




```HTML
<input id="get-info" type="button" value="Get info" /><br />
<span id="message"></span>
```




```js
(function () {
    "use strict";
    var taskGuids = [];

    // The initialize function must be run each time a new page is loaded.
    Office.initialize = function (reason) {
        $(document).ready(function () {

            // After the DOM is loaded, add-in-specific code can run.
            app.initialize();
            $('#get-info').click(getTaskInfo);
        });
    };

    // Get the maximum task index, and then get the task GUIDs.
    function getTaskInfo() {
        getMaxTaskIndex().then(
            function (data) {
                getTaskGuids(data);
            }
        );
    }

    // Get the maximum index of the tasks for the current project.
    function getMaxTaskIndex() {
        var defer = $.Deferred();
        Office.context.document.getMaxTaskIndexAsync(
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

    // Get each task GUID, and then display the GUIDs in the add-in.
    function getTaskGuids(maxTaskIndex) {
        var defer = $.Deferred();
        for (var i = 0; i <= maxTaskIndex; i++) {
            getTaskGuid(i);
        }
        return defer.promise();
        function getTaskGuid(index) {
            Office.context.document.getTaskByIndexAsync(index,
                function (result) {
                    if (result.status === Office.AsyncResultStatus.Succeeded) {
                        taskGuids.push(result.value);
                        if (index == maxTaskIndex) {
                            defer.resolve();
                            $('#message').html(taskGuids.toString());
                        }
                    }
                    else {
                        onError(result.error);
                    }
                }
            );
        }
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
|**Minimum permission level**|[ReadDocument](../../docs/develop/requesting-permissions-for-api-use-in-content-and-task-pane-add-ins.md)|
|**Add-in types**|Task pane|
|**Library**|Office.js|
|**Namespace**|Office|

## Support history

|**Version**|**Changes**|
|:-----|:-----|
|1.1|Introduced|

## See also



#### Other resources


[getMaxTaskIndexAsync](../../reference/shared/projectdocument.getmaxtaskindexasync.md)
[AsyncResult object](../../reference/shared/asyncresult.md)
[ProjectDocument object](../../reference/shared/projectdocument.projectdocument.md)
