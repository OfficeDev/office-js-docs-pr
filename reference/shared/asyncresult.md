
# AsyncResult object
An object which encapsulates the result of an asynchronous request, including status and error information if the request failed.

|||
|:-----|:-----|
|**Hosts:**|Access, Excel, Outlook, PowerPoint, Project, Word|
|**Last changed in**|1.1|

```
AsyncResult
```


## Members


**Properties**


|**Name**|**Description**|
|:-----|:-----|
|**[asyncContext](../../reference/shared/asyncresult.asynccontext.md)**|Gets the user-defined item passed to the optional  _asyncContext_ parameter of the invoked method in the same state as it was passed in.|
|**[error](../../reference/shared/asyncresult.error.md)**|Gets an  **Error** object that provides a description of the error, if any error occurred.|
|**[status](../../reference/shared/asyncresult.status.md)**|Gets the status of the asynchronous operation.|
|**[value](../../reference/shared/asyncresult.value.md)**|Gets the payload or content of this asynchronous operation, if any.|

## Remarks

When the function you pass to the  _callback_ parameter of an "Async" method executes, it receives an [AsyncResult](../../reference/shared/asyncresult.md) object that you can access from the callback function's only parameter.

The following is an example applicable to content and task pane add-ins. The example shows a call to the [getSelectedDataAsync](../../reference/shared/document.getselecteddataasync.md) method of the **Document** object.




```js
Office.context.document.getSelectedDataAsync("text", {valueFormat:"unformatted", filterType:"all"}, 
   function (result) {
      if (result.status === "success")      
         var dataValue = result.value; // Get selected data.
         write('Selected data is ' + dataValue);
      else {            
         var err = result.error; 
         write(err.name + ": " + err.message);
      }
   });
// Function that writes to a div with id='message' on the page.
function write(message){
    document.getElementById('message').innerText += message; 
}

```

The anonymous function passed as the  _callback_ argument ( `function (result){...}`) has a single parameter named  _result_ that provides access to an **AsyncResult** object when the function executes. When the call to the **getSelectedDataAsync** method completes, the callback function executes, and the following line of code accesses the **value** property of the **AsyncResult** object to return the data selected in the document:

 `var dataValue = result.value;`

Note that other lines of code in the function use the  _result_ parameter of the callback function to access the **status** and **error** properties of the **AsyncResult** object.

The  **AsyncResult** object is available from the function passed as the argument to the _callback_ parameter of the following methods:



|**Parent Object**|**Method**|
|:-----|:-----|
|**Document**(Excel, PowerPoint, Project, and Word only)|[getSelectedDataAsync](../../reference/shared/document.getselecteddataasync.md)|
||[setSelectedDataAsync](../../reference/shared/document.setselecteddataasync.md)|
|**Bindings** (Excel and Word only)|[addFromPromptAsync](../../reference/shared/bindings.addfrompromptasync.md)|
||[addFromSelectionAsync](../../reference/shared/bindings.addfromselectionasync.md)|
||[getAllAsync](../../reference/shared/bindings.getallasync.md)|
||[getByIdAsync](../../reference/shared/bindings.getbyidasync.md)|
||[releaseByIdAsync](../../reference/shared/bindings.releasebyidasync.md)|
|**Binding** (Excel and Word only)|[getDataAsync](../../reference/shared/binding.getdataasync.md)|
||[setDataAsync](../../reference/shared/binding.setdataasync.md)|
||[removeHandlerAsync](../../reference/shared/binding.removehandlerasync.md)|
|**TableBinding** (Excel and Word only)||
||[addRowsAsync](../../reference/shared/binding.tablebinding.addrowsasync.md)|
||[deleteAllDataValuesAsync](../../reference/shared/binding.tablebinding.deletealldatavaluesasync.md)|
|**Settings** (Excel, PowerPoint, and Word only)|[refreshAsync](../../reference/shared/settings.refreshasync.md)|
||[saveAsync](../../reference/shared/settings.saveasync.md)|
|**CustomXmlNode** (Word only)|[getNodesAsync](../../reference/shared/customxmlnode.getnodesasync.md)|
||[getNodeValueAsync](../../reference/shared/customxmlnode.getnodevalueasync.md)|
||[getXmlAsync](../../reference/shared/customxmlnode.getxmlasync.md)|
||[setNodeValueAsync](../../reference/shared/customxmlnode.setnodevalueasync.md)|
||[setXmlAsync](../../reference/shared/customxmlnode.setxmlasync.md)|
|**CustomXmlPart** (Word only)|[deleteAsync](../../reference/shared/customxmlpart.deleteasync.md)|
||[getNodesAsync](../../reference/shared/customxmlpart.getnodesasync.md)|
||[getXmlAsync](../../reference/shared/customxmlpart.getxmlasync.md)|
|**CustomXmlParts** (Word only)|[addAsync](../../reference/shared/customxmlparts.addasync.md)|
||[getByIdAsync](../../reference/shared/customxmlparts.getbyidasync.md)|
||[getByNamespaceAsync](../../reference/shared/customxmlparts.getbynamespaceasync.md)|
|**CustomXmlPrefixMappings** (Word only)|[addNamespaceAsync](../../reference/shared/customxmlprefixmappings.addnamespaceasync.md)|
||[getNamespaceAsync](../../reference/shared/customxmlprefixmappings.getnamespaceasync.md)|
||[getPrefixAsync](../../reference/shared/customxmlprefixmappings.getprefixasync.md)|
|**Mailbox** (Outlook only)|[getUserIdentityTokenAsync](http://msdn.microsoft.com/library/c658518b-6867-41a0-99cf-810303e4c539%28Office.15%29.aspx)|
||[makeEwsRequestAsync](http://msdn.microsoft.com/library/2ec380e0-4a67-4146-92a6-6a39f65dc6f2%28Office.15%29.aspx)|
|**CustomProperties** (Outlook only)|[saveAsync](http://msdn.microsoft.com/library/690d5aa9-62b5-4e5c-9548-62dfdbb5fa56%28Office.15%29.aspx)|
|**Item** (Outlook only)|[loadCustomPropertiesAsync](http://msdn.microsoft.com/library/dfbec151-8ea7-4915-b723-09ea1396a261%28Office.15%29.aspx)|
|**RoamingSettings** (Outlook only)|[saveAsync](http://msdn.microsoft.com/library/a616f71c-a447-423f-a0d2-e9d6f1ac32f8%28Office.15%29.aspx)|

## Support details


A capital Y in the following matrix indicates that this method is supported in the corresponding Office host application. An empty cell indicates that the Office host application doesn't support this method.

For more information about Office host application and server requirements, see [Requirements for running Office Add-ins](../../docs/overview/requirements-for-running-office-add-ins.md).



| |**Office for Windows desktop**|**Office Online (in browser)**|**Office for iPad**|**OWA for Devices**|**Outlook for Mac**|
|:-----|:-----|:-----|:-----|:-----|:-----|
|**Access**||Y||||
|**Excel**|Y|Y|Y|||
|**Outlook**|Y|Y||Y|Y|
|**PowerPoint**|Y|Y|Y|||
|**Project**|Y|||||
|**Word**|Y|Y|Y|||

|||
|:-----|:-----|
|**Add-in types**|Content, task pane, Outlook|
|**Library**|Office.js|
|**Namespace**|Office|

## Support history



****


|**Version**|**Changes**|
|:-----|:-----|
|1.1|Added support for Excel, PowerPoint, and Word in Office for iPad.|
|1.1|Added support for add-ins for Access.|
|1.0|Introduced|
