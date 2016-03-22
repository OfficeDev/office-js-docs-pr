
# CustomXmlNode.getTextAsync method
Asynchronously gets the text of an XML node in a custom XML part.

|||
|:-----|:-----|
|**Hosts:**|Word|
|**Available in [Requirement set](../../docs/overview/specify-office-hosts-and-api-requirements.md)**|CustomXmlParts|
|**Added in**|1.2|

```js
customXmlNodeObj.getTextAsync([asyncContext,]callback(asyncResult);
```


## Parameters



|**Name**|**Type**|**Description**|
|:-----|:-----|:-----|
| _asyncContext_|**object**|Optional. A user-defined object that is available on the [AsyncResult](../../reference/shared/asyncresult.md) object's asyncContext property. Use this to provide an object or value to the **AsyncResult** when the callback is a named function.|
| _callback_|**object**|A function that is invoked when the callback returns, whose only parameter is of type  **AsyncResult**.|

## Callback Value

When the function you passed to the  _callback_ parameter executes, it receives an [AsyncResult](../../reference/shared/asyncresult.md) object that you can access from the callback function's only parameter.

In the callback function passed to the  **getTextAsync** method, you can use the properties of the **AsyncResult** object to return the following information.



|**Property**|**Use to...**|
|:-----|:-----|
|[AsyncResult.value](../../reference/shared/asyncresult.value.md)|Access a  **string** that contains the inner text of the referenced nodes.|
|[AsyncResult.status](../../reference/shared/asyncresult.status.md)|Indicates the success or failure of the operation.|
|[AsyncResult.error](../../reference/shared/asyncresult.error.md)|Access an [Error](../../reference/shared/error.md) object that provides error information if the operation failed.|
|[AsyncResult.asyncContext](../../reference/shared/asyncresult.asynccontext.md)|Access your user-defined  **object** or value, if you passed one as the _asyncContext_ parameter. This property returns undefined if the _asyncContext_ has not been set.|

## Example

Learn how to get the text value of a node in a custom XML part.


```js
// Get the built-in core properties XML part by using its ID. This results in a call to Word.
Office.context.document.customXmlParts.getByIdAsync("{6C3C8BC8-F283-45AE-878A-BAB7291924A1}", function (getByIdAsyncResult) {
    
    // Access the XML part.
    var xmlPart = getByIdAsyncResult.value;
    
    // Add namespaces to the namespace manager. These two calls result in two calls to Word.
    xmlPart.namespaceManager.addNamespaceAsync('cp', 'http://schemas.openxmlformats.org/package/2006/metadata/core-properties', function () {
        xmlPart.namespaceManager.addNamespaceAsync('dc', 'http://purl.org/dc/elements/1.1/', function () {

            // Get XML nodes by using an Xpath expression. This results in a call to Word.
            xmlPart.getNodesAsync("/cp:coreProperties/dc:title", function (getNodesAsyncResult) {
                
                // Get the first node returned by using the Xpath expression. 
                var node = getNodesAsyncResult.value[0];
                
                // Get the text value of the node and use the asyncContext. This results in a call to Word. 
                // The results are logged to the browser console.
                node.getTextAsync({asyncContext: "StateNormal"}, function (getTextAsyncResult) {
                   console.log("Text of the title element = " + getTextAsyncResult.value;
                   console.log("The asyncContext value = " + getTextAsyncResult.asyncContext;
                });
            });
        });
    });
});
```


## Support details


A capital Y in the following matrix indicates that this method is supported in the corresponding Office host application. An empty cell indicates that the Office host application doesn't support this method.

For more information about Office host application and server requirements, see [Requirements for running Office Add-ins](../../docs/overview/requirements-for-running-office-add-ins.md).


||**Office for Windows desktop**|**Office Online (in browser)**|**Office for iPad**|
|:-----|:-----|:-----|:-----|
|**Word**|Y|Y|Y|

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
|1.1|Added getTextAsync.|
