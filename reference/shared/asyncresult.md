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
|**[asyncContext](asyncresult.asynccontext.md)**|Gets the user-defined item passed to the optional  _asyncContext_ parameter of the invoked method in the same state as it was passed in.|
|**[error](asyncresult.error.md)**|Gets an  **Error** object that provides a description of the error, if any error occurred.|
|**[status](asyncresult.status.md)**|Gets the status of the asynchronous operation.|
|**[value](asyncresult.value.md)**|Gets the payload or content of this asynchronous operation, if any.|

## Remarks

When the function you pass to the  _callback_ parameter of an "Async" method executes, it receives an **AsyncResult** object that you can access from the callback function's only parameter.

The following is an example applicable to content and task pane add-ins. The example shows a call to the [getSelectedDataAsync][] method of the **Document** object.




```js
Office.context.document.getSelectedDataAsync("text", {
        valueFormat: "unformatted",
        filterType: "all"
    },
    function (result) {
        if (result.status === "success") {
            var dataValue = result.value; // Get selected data.
            console.log('Selected data is ' + dataValue);
        } else {
            var err = result.error;
            console.log(err.name + ": " + err.message);
        }
    });
```

The anonymous function passed as the  _callback_ argument ( `function (result){...}`) has a single parameter named  _result_ that provides access to an **AsyncResult** object when the function executes. When the call to the **getSelectedDataAsync** method completes, the callback function executes, and the following line of code accesses the **value** property of the **AsyncResult** object to return the data selected in the document:

 `var dataValue = result.value;`

Note that other lines of code in the function use the  _result_ parameter of the callback function to access the **status** and **error** properties of the **AsyncResult** object.

The  **AsyncResult** object is available from the function passed as the argument to the _callback_ parameter of the following methods:



| **Parent Object** | **Method** |
|:------------------|:-----------|
|**Auth**|[getAccessTokenAsync][]
|**Binding** (Excel and Word only)|[getDataAsync][]
|          |[setDataAsync][]
|          |[removeHandlerAsync][]
|**Bindings** (Excel and Word only)|[addFromPromptAsync][]
|          |[addFromSelectionAsync][]
|          |[addFromNamedItemAsync][]
|          |[getAllAsync][]
|          |[getByIdAsync][]
|          |[releaseByIdAsync][]
|**CustomProperties** (Outlook only)|[saveAsync][]
|**CustomXmlNode** (Word only)|[getNodesAsync][]
|          |[getNodeValueAsync][]
|          |[getXmlAsync][]
|          |[setNodeValueAsync][]
|          |[setXmlAsync][]
|**CustomXmlPart** (Word only)|[deleteAsync][]
|          |[getNodesAsync][]
|          |[getXmlAsync][]
|**CustomXmlParts** (Word only)|[addAsync][]
|          |[getByIdAsync][]
|          |[getByNamespaceAsync][]
|**CustomXmlPrefixMappings** (Word only)|[addNamespaceAsync][]
|          |[getNamespaceAsync][]
|          |[getPrefixAsync][]
|**Document** (Excel, PowerPoint, Project, and Word only)|[getSelectedDataAsync][]
|          |[setSelectedDataAsync][]
|          |[getFileAsync][]
|          |[getFilePropertiesAsync][]
|          |[getActiveViewAsync][]
|**File**|[getSliceAsync][]
|**Mailbox** (Outlook only)|[getUserIdentityTokenAsync][]
|          |[makeEwsRequestAsync][]
|**Item** (Outlook only)|[loadCustomPropertiesAsync][]
|**TableBinding** (Excel and Word only)|[addRowsAsync][]
|          |[deleteAllDataValuesAsync][]
|**RoamingSettings** (Outlook only)|[saveAsync][]
|**Settings** (Excel, PowerPoint, and Word only)|[refreshAsync][]
|          |[saveAsync][]
|**UI**|[displayDialogAsync][]


## Support history


|**Version**|**Changes**|
|:-----|:-----|
|1.1|Added support for Excel, PowerPoint, and Word in Office for iPad.|
|1.1|Added support for add-ins for Access.|
|1.0|Introduced|


[addRowsAsync]: binding.tablebinding.addrowsasync.md
[addFromNamedItemAsync]: bindings.addfromnameditemasync.md
[addFromPromptAsync]: bindings.addfrompromptasync.md
[addFromSelectionAsync]: bindings.addfromselectionasync.md
[addAsync]: customxmlparts.addasync.md
[addNamespaceAsync]: customxmlprefixmappings.addnamespaceasync.md
[deleteAllDataValuesAsync]: binding.tablebinding.deletealldatavaluesasync.md
[displayDialogAsync]: office.ui.displaydialogasync.md
[deleteAsync]: customxmlpart.deleteasync.md
[getActiveViewAsync]: document.getactiveviewasync.md
[getAccessTokenAsync]: office.context.auth.getAccessTokenAsync.md
[getAllAsync]: bindings.getallasync.md
[getByNamespaceAsync]: customxmlparts.getbynamespaceasync.md
[getByIdAsync]: bindings.getbyidasync.md
[getDataAsync]: binding.getdataasync.md
[getFilePropertiesAsync]: document.getfilepropertiesasync.md
[getFileAsync]: document.getfileasync.md
[getNodesAsync]: customxmlnode.getnodesasync.md
[getNodeValueAsync]: customxmlnode.getnodevalueasync.md
[getNamespaceAsync]: customxmlprefixmappings.getnamespaceasync.md
[getPrefixAsync]: customxmlprefixmappings.getprefixasync.md
[getSliceAsync]: file.getsliceasync.md
[getSelectedDataAsync]: document.getselecteddataasync.md
[getUserIdentityTokenAsync]: ../../reference/outlook/1.5/Office.context.mailbox.md
[getXmlAsync]: customxmlpart.getxmlasync.md
[loadCustomPropertiesAsync]: ../../reference/outlook/1.5/CustomProperties.md
[makeEwsRequestAsync]: ../../reference/outlook/1.5/Office.context.mailbox.md
[releaseByIdAsync]: bindings.releasebyidasync.md
[removeHandlerAsync]: binding.removehandlerasync.md
[refreshAsync]: settings.refreshasync.md
[setNodeValueAsync]: customxmlnode.setnodevalueasync.md
[setXmlAsync]: customxmlnode.setxmlasync.md
[saveAsync]: settings.saveasync.md
[setSelectedDataAsync]: document.setselecteddataasync.md
[setDataAsync]: binding.setdataasync.md