

# ProjectDocument.getWSSUrlAsync method
Asynchronously gets the URL of the synchronized SharePoint task list.

|||
|:-----|:-----|
|**Hosts:**|Project|
|**Available in [Requirement set](../../docs/overview/specify-office-hosts-and-api-requirements.md)**|Selection|
|**Added in**|1.0|

```js
Office.context.document.getWSSUrlAsync([options,] [callback]);
```


## Parameters



|**Name**|**Type**|**Description**|**Support notes**|
|:-----|:-----|:-----|:-----|
| _options_|**object**|Specifies any of the following [optional parameters](../../docs/develop/asynchronous-programming-in-office-add-ins.md#passing-optional-parameters-to-asynchronous-methods).||
| _asyncContext_|**array**,  **boolean**,  **null**,  **number**,  **object**, **string**, or  **undefined**|A user-defined item of any type that is returned in the  **AsyncResult** object without being altered.||
| _callback_|**object**|A function that is invoked when the callback returns, whose only parameter is of type  **AsyncResult**.||

## Callback Value

When the  _callback_ function executes, it receives an [AsyncResult](../../reference/shared/asyncresult.md) object that you can access from the parameter in the callback function.

For the  **getWSSUrlAsync** method, the returned [AsyncResult](../../reference/shared/asyncresult.md) object contains following properties.


|**Name**|**Description**|
|:-----|:-----|
|[asyncContext](../../reference/shared/asyncresult.asynccontext.md)|The data passed in the optional  _asyncContext_ parameter, if the parameter was used.|
|[error](../../reference/shared/asyncresult.error.md)|Information about the error, if the  **status** property equals **failed**.|
|[status](../../reference/shared/asyncresult.status.md)|The  **succeeded** or **failed** status of the asynchronous call.|
|[value](../../reference/shared/asyncresult.value.md)|Contains the following properties:<br/><br/><ul><li>The <b>listName</b> property is the name of the synchronized SharePoint task list.</li><li>The <b>serverUrl</b> property is the URL of the synchronized SharePoint task list.</li></ul>|

## Remarks

If the active project is not synchronized with a SharePoint task list, the  **listName** and **serverUrl** values are empty.


## Example

The following code example calls  **getWSSUrlAsync** to get the name and URL of the synchronized SharePoint task list.

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
            getSharePointTaskListUrl();
        });
    };

    // Get the URL of the the synchronized SharePoint task list.
    function getSharePointTaskListUrl() {
        Office.context.document.getWSSUrlAsync(
            function (result) {
                if (result.status === Office.AsyncResultStatus.Succeeded) {
                    var output = String.format(
                        'List name: {0}<br />List URL: {1}',
                        result.value.listName, result.value.serverUrl);
                    $('#message').html(output);
                }
                else {
                    onError(result.error);
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


[AsyncResult object](../../reference/shared/asyncresult.md)
[ProjectDocument object](../../reference/shared/projectdocument.projectdocument.md)
