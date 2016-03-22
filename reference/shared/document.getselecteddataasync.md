
# Document.getSelectedDataAsync method
Reads the data contained in the current selection in the document.

|||
|:-----|:-----|
|**Hosts:**|Access, Excel, PowerPoint, Project, Word|
|**Available in requirement sets**|Selection|
|**Last changed in Selection**|1.1|

[![Try out this call in the interactive API Tutorial for Excel](../../images/819b84bf-151c-4a12-80c3-d6f8d7c03251.png)](http://officeapitutorial.azurewebsites.net/Redirect.html?scenario=Write+and+Read+Text&amp;task=writeSelectedDataText)


```js
Office.context.document.getSelectedDataAsync(coercionType [, options], callback); 
```


## Parameters



|**Name**|**Type**|**Description**|**Support notes**|
|:-----|:-----|:-----|:-----|
| _coercionType_|[CoercionType](../../reference/shared/coerciontype-enumeration.md)<br/><table><tr><td></td><td><b>Host support</b></td></tr><tr><td><b>Office.CoercionType.Text</b> (string)</td><td>Excel, Excel Online, PowerPoint, PowerPoint Online, Word, and Word Online only</td></tr><tr><td><b>Office.CoercionType.Matrix</b> (array of arrays)</td><td>Excel, Word, and Word Online only</td></tr><tr><td><b>Office.CoercionType.Table</b> ([TableData](../../reference/shared/tabledata.md) object)</td><td>Access, Excel, Word, and Word Online only</td></tr><tr><td><b>Office.CoercionType.Html</b></td><td>Word only.</td></tr><tr><td><b>Office.CoercionType.Ooxml</b> (Office Open XML)</td><td>Word and Word Online only</td></tr><tr><td><b>Office.CoercionType.SlideRange</b></td><td>PowerPoint, and PowerPoint Online only</td></tr></table>|The type of data structure to return. Required.||
| _options_|**object**<br/><table><tr><td><i>valueFormat</i></td><td><b>[ValueFormat](../../reference/shared/valueformat-enumeration.md)</b></td><td>Specifies whether to return the result with its number or date values formatted or unformatted.</td><td></td></tr><tr><td><i>filterType</i></td><td>[FilterType](../../reference/shared/filtertype-enumeration.md)</td><td>Specifies whether to apply filtering when the data is retrieved. Optional.</td><td>This parameter is ignored in Word documents.</td></tr><tr><td><i>asyncContext</i></td><td><b>array</b>,  <b>boolean</b>,  <b>null</b>,  <b>number</b>,  <b>object</b>,  <b>string</b>, or <b>undefined</b></td><td>A user-defined item of any type that is returned in the  <b>AsyncResult</b> object without being altered.</td><td></td></tr></table>|Specifies any of the following [optional parameters](../../docs/develop/asynchronous-programming-in-office-add-ins.md#passing-optional-parameters-to-asynchronous-methods)||
| _callback_|**object**|A function that is invoked when the callback returns, whose only parameter is of type  **AsyncResult**.||

## Callback Value

When the function you passed to the  _callback_ parameter executes, it receives an [AsyncResult](../../reference/shared/asyncresult.md) object that you can access from the callback function's only parameter.

In the callback function passed to the  **getSelectedDataAsync** method, you can use the properties of the **AsyncResult** object to return the following information.



|**Property**|**Use to...**|
|:-----|:-----|
|[AsyncResult.value](../../reference/shared/asyncresult.value.md)|Access the values in the current selection, which are returned in the data structure or format you specified with the  _coercionType_ parameter. (See **Remarks** for more information about data coercion.)|
|[AsyncResult.status](../../reference/shared/asyncresult.status.md)|Determine the success or failure of the operation.|
|[AsyncResult.error](../../reference/shared/asyncresult.error.md)|Access an [Error](../../reference/shared/error.md) object that provides error information if the operation failed.|
|[AsyncResult.asyncContext](../../reference/shared/asyncresult.asynccontext.md)|Access your user-defined  **object** or value, if you passed one as the _asyncContext_ parameter.|

## Remarks

In your task pane or content add-in, use the  **getSelectedDataAsync** method to write script that reads the data from the user's selection in a document, spreadsheet, presentation, or project. For example, after a user selects content in a Word document, you can use the **getSelectedDataAsync** method to read that selection, and then submit it to a web service as a query or some other operation.

After reading the selection, you can also use the [setSelectedDataAsync](../../reference/shared/document.setselecteddataasync.md) and [addHandlerAsync](../../reference/shared/document.addhandlerasync.md) methods of the **Document** object to [write back to the selection or add an event handler](../../docs/develop/read-and-write-data-to-the-active-selection-in-a-document-or-spreadsheet.md) to detect if the user changes the selection.

The  **getSelectedDataAsync** method can read from the selection only as long as it's active. In add-ins for Word and Excel, if you need to make a persistent association to read and write to the user's selection, instead use the [Bindings.addFromSelectionAsync](../../reference/shared/bindings.addfromselectionasync.md) method to [bind to that selection](../../docs/develop/bind-to-regions-in-a-document-or-spreadsheet.md).

Use the  _coercionType_ parameter of the **getSelectedDataAsync** method to specify the data structure or format of the selected data being read.



|**Specified  _coercionType_**|**Data returned**|**Office host application support**|
|:-----|:-----|:-----|
|**Office.CoercionType.Text** or `"text"`|A string.|Word, Excel, PowerPoint, and Project.<br/><br/> **Note**: In Excel, even when a subset of a cell is selected, the entire cell contents are returned.|
|**Office.CoercionType.Matrix** or `"matrix"`|An array of arrays. For example,  ` [['a','b'], ['c','d']]` for a selection of two rows in two columns.|Word and Excel.|
|**Office.CoercionType.Table** or `"table"`|A [TableData](../../reference/shared/tabledata.md) object for reading a table with headers.|Word and Excel.|
|**Office.CoercionType.Html** or `"html"`|In HTML format.|Word only.|
|**Office.CoercionType.Ooxml** or `"ooxml"`|In Open Office XML (OpenXML) format.|Word only.<br/><br/> **Tip**: When developing your add-in's code, you can use the  `"ooxml"` _coercionType_ of the **getSelectedDataAsync** method to see how the content you select in a Word document is defined as OpenXML tags. Then, use those tags in the data parameter of the [Document.setSelectedDataAsync](../../reference/shared/document.setselecteddataasync.md) method to write content with that formatting or structure to a document. For example, you can [insert an image into a document](http://blogs.msdn.com/b/officeapps/archive/2012/10/26/inserting-images-with-apps-for-office.aspx) as OpenXML.|
|**Office.CoercionType.SlideRange** or "slideRange"|A JSON object that contains an array named "slides" that contains the ids, titles, and indexes of the selected slides.  **Note:** To select more than one slide, the user must be editing the presentation in **Normal**,  **Outline View**, or  **Slide Sorter** view. Also, this method isn't supported in **Master Views**.For example,  `{"slides":[{"id":257,"title":"Slide 2","index":2},{"id":256,"title":"Slide 1","index":1}]}` for a selection of two slides.|PowerPoint only.|
If the data structure of the selection doesn't match the specified  _coercionType_, the  **getSelectedDataAsync** method will attempt to coerce the data into that type or structure. If the selection can't be coerced into the **Office.CoercionType** you specified, the **AsyncResult.status** property returns `"failed"`.


## Example

To read the value of the current selection, you need to write a callback function that reads the selection. The following example shows how to:


-  **Pass an anonymous callback function** that reads the value of the current selection to the _callback_ parameter of the **getSelectedDataAsync** method.
    
-  **Read the selection** as text, unformatted, and not filtered.
    
-  **Display the value** on the add-in's page.
    

```js
function getText() {
    Office.context.document.getSelectedDataAsync(Office.CoercionType.Text, 
        { valueFormat: "unformatted", filterType: "all" },
        function (asyncResult) {
            var error = asyncResult.error;
            if (asyncResult.status === Office.AsyncResultStatus.Failed) {
                write(error.name + ": " + error.message);
            } 
            else {
                // Get selected data.
                var dataValue = asyncResult.value; 
                write('Selected data is ' + dataValue);
            }            
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


**Supported hosts, by platform**


||**Office for Windows desktop**|**Office Online (in browser)**|**Office for iPad**|
|:-----|:-----|:-----|:-----|
|**Access**||Y||
|**Excel**|Y|Y|Y|
|**PowerPoint**|Y|Y|Y|
|**Project**|Y|||
|**Word**|Y|Y|Y|

|||
|:-----|:-----|
|**Available in requirement sets**|Selection|
|**Minimum permission level**|[ReadDocument (ReadAllDocument required to get Office Open XML)](../../docs/develop/requesting-permissions-for-api-use-in-content-and-task-pane-add-ins.md)|
|**Add-in types**|Content, task pane|
|**Library**|Office.js|
|**Namespace**|Office|

## Support history



|**Version**|**Changes**|
|:-----|:-----|
|1.1|Added support for PowerPoint Online.|
|1.1| In Word Online, added support for **Office.CoercionType.Matrix** and **Office.CoercionType.Table** as the _coercionType_ parameter.|
|1.1|In Excel, PowerPoint, and Word in Office for iPad, added the same level of support as Excel, PowerPoint and Word on Windows desktop.|
|1.1| In Word Online, added support for **Office.CoercionType.Text** as the _coercionType_ parameter.|
|1.1|In content add-ins for PowerPoint, you can get the ids, titles, and indexes of the selected range of slides by passing  **Office.CoercionType.SlideRange** as the _coercionType_ parameter of the **getSelectedDataAsync** method. See the [Document.goToByIdAsync](../../reference/shared/document.gotobyidasync.md) method topic for an example of how to use this value to navigate to the currently selected slide.|
|1.0|Introduced|
