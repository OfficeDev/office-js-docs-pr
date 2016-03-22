
# CustomXmlParts.addAsync method
Asynchronously adds a new custom XML part to a file.

|||
|:-----|:-----|
|**Hosts:**|Word|
|**Available in [Requirement set](../../docs/overview/specify-office-hosts-and-api-requirements.md)**|CustomXmlParts|
|**Added in**|1.1|

```
Office.context.document.customXmlParts.addAsync(xml [, options], callback);
```


## Parameters



|**Name**|**Type**|**Description**|**Support notes**|
|:-----|:-----|:-----|:-----|
| _xml_|**string**|The XML to add to the newly created custom XML part. Required.||
| _options_|**object**|Specifies any of the following [optional parameters](../../docs/develop/asynchronous-programming-in-office-add-ins.md#passing-optional-parameters-to-asynchronous-methods)||
| _asyncContext_|**array**,  **boolean**,  **null**,  **number**,  **object**, **string**, or  **undefined**|A user-defined item of any type that is returned in the  **AsyncResult** object without being altered.||
| _callback_|**object**|A function that is invoked when the callback returns, whose only parameter is of type  **AsyncResult**.||

## Callback Value

When the function you passed to the  _callback_ parameter executes, it receives an [AsyncResult](../../reference/shared/asyncresult.md) object that you can access from the callback function's only parameter.

In the callback function passed to the  **addAsync** method, you can use the properties of the **AsyncResult** object to return the following information.



|**Property**|**Use to...**|
|:-----|:-----|
|[AsyncResult.value](../../reference/shared/asyncresult.value.md)|Returns the newly created [CustomXmlPart](../../reference/shared/customxmlpart.customxmlpart.md) object.|
|[AsyncResult.status](../../reference/shared/asyncresult.status.md)|Determine the success or failure of the operation.|
|[AsyncResult.error](../../reference/shared/asyncresult.error.md)|Access an [Error](../../reference/shared/error.md) object that provides error information if the operation failed.|
|[AsyncResult.asyncContext](../../reference/shared/asyncresult.asynccontext.md)|Access your user-defined  **object** or value, if you passed one as the _asyncContext_ parameter.|

## Example




```js
function addXMLPart() {
    Office.context.document.customXmlParts.addAsync('<root categoryId="1" xmlns="http://tempuri.org"><item name="Cheap Item" price="$193.95"/><item name="Expensive Item" price="$931.88"/></root>', function (result) {
        });
}
```








```
function addXMLPartandHandler() {
    Office.context.document.customXmlParts.addAsync("<testns:book xmlns:testns='http://testns.com'><testns:page number='1'>Hello</testns:page><testns:page number='2'>world!</testns:page></testns:book>", 
        function(r) { r.value.addHandlerAsync("nodeDeleted", 
            function(a) {write(a.type)
            }, 
                function(s) {write(s.status)
                }); 
        });
}

// Function that writes to a div with id='message' on the page.
function write(message){
    document.getElementById('message').innerText += message; 
}
```


## Support details


A capital Y in the following matrix indicates that this method is supported in the corresponding Office host application. An empty cell indicates that the Office host application doesn't support this method.

For more information about Office host application and server requirements, see [Requirements for running Office Add-ins](../../docs/overview/requirements-for-running-office-add-ins.md).


||**Office for Windows desktop**|**Office Online (in browser)**|**Office for iPad**|
|:-----|:-----|:-----|:-----|
|**Word**|Y||Y|

|||
|:-----|:-----|
|**Available in requirement sets**|CustomXmlParts|
|**Minimum permission level**|[ReadWriteDocument](../../docs/develop/requesting-permissions-for-api-use-in-content-and-task-pane-add-ins.md)|
|**Add-in types**|Task pane|
|**Library**|Office.js|
|**Namespace**|Office|

## Support history



****


|**Version**|**Changes**|
|:-----|:-----|
|1.1|Added support for Word in Office for iPad.|
|1.0|Introduced|
