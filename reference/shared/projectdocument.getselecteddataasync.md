
# ProjectDocument.getSelectedDataAsync method
Asynchronously gets the text value of the data that is contained in the current selection of one or more cells in the Gantt Chart view.

|||
|:-----|:-----|
|**Hosts:**|Project|
|**Available in [Requirement set](../../docs/overview/specify-office-hosts-and-api-requirements.md)**|Selection|
|**Added in**|1.0|

```
Office.context.document.getSelectedDataAsync(coercionType[, options][, callback]);
```


## Parameters



|**Name**|**Type**|**Description**|**Support notes**|
|:-----|:-----|:-----|:-----|
| _coercionType_|[CoercionType](https://dev.office.com/reference/add-ins/shared/coerciontype-enumeration)|The type of data structure to return. Required.<br/>Project 2013 supports only  **Office.CoercionType.Text** or `"text"`.||
| _options_|**object**|Specifies any of the following [optional parameters](../../docs/develop/asynchronous-programming-in-office-add-ins.md#passing-optional-parameters-to-asynchronous-methods).||
| _valueFormat_|[ValueFormat](https://dev.office.com/reference/add-ins/shared/valueformat-enumeration)|The formatting to use for number or date values.<br/>Project 2013 ignores this parameter and internally sets it to  `unformatted`.||
| _filterType_|[FilterType](https://dev.office.com/reference/add-ins/shared/filtertype-enumeration)|Specifies whether to include only visible data or all data. <br/>Project 2013 ignores this parameter and internally sets it to  `all`.||
| _asyncContext_|**array**,  **boolean**,  **null**,  **number**,  **object**, **string**, or  **undefined**|A user-defined item of any type that is returned in the  **AsyncResult** object without being altered.||
| _callback_|**object**|A function that is invoked when the callback returns, whose only parameter is of type  **AsyncResult**.||

## Callback Value

When the  _callback_ function executes, it receives an [AsyncResult](https://dev.office.com/reference/add-ins/shared/asyncresult) object that you can access from the parameter in the callback function.

For the  **getSelectedDataAsync** method, the returned [AsyncResult](https://dev.office.com/reference/add-ins/shared/asyncresult) object contains the following properties.


****


|**Name**|**Description**|
|:-----|:-----|
|[asyncContext](https://dev.office.com/reference/add-ins/shared/asyncresult.asynccontext)|The data that was passed in the optional  _asyncContext_ parameter, if the parameter was used.|
|[error](https://dev.office.com/reference/add-ins/shared/asyncresult.error)|Information about the error, if the  **status** property equals **failed**.|
|[status](https://dev.office.com/reference/add-ins/shared/asyncresult.status)|The  **succeeded** or **failed** status of the asynchronous call.|
|[value](https://dev.office.com/reference/add-ins/shared/asyncresult.value)|The text value of the selected cells.|

## Remarks

The  **ProjectDocument.getSelectedDataAsync** method overrides the [Document.getSelectedDataAsync](https://dev.office.com/reference/add-ins/shared/document.getselecteddataasync) method and returns the text value of data that is selected in one or more cells in the Gantt Chart view. **ProjectDocument.getSelectedDataAsync** supports only a text format as the [CoercionType](https://dev.office.com/reference/add-ins/shared/coerciontype-enumeration) - it does not support  `matrix`,  `table`, or other formats.


## Example

The following code example gets the values of the selected cells. It uses the optional  _asyncContext_ parameter to pass some text to the callback function.

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
            $('#get-info').click(getSelectedText);
        });
    };

    // Get the text from the selected cells in the document, and display it in the add-in.
    function getSelectedText() {
        Office.context.document.getSelectedDataAsync(
            Office.CoercionType.Text,
            {asyncContext: 'Some related info'},
            function (result) {
                if (result.status === Office.AsyncResultStatus.Failed) {
                    onError(result.error);
                }
                else {
                    var output = String.format(
                        'Selected text: {0}<br/>Passed info: {1}',
                        result.value, result.asyncContext);
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


[AsyncResult object](https://dev.office.com/reference/add-ins/shared/asyncresult)

[Office.CoercionType](https://dev.office.com/reference/add-ins/shared/coerciontype-enumeration)

[ProjectDocument object](https://dev.office.com/reference/add-ins/shared/projectdocument.projectdocument)
