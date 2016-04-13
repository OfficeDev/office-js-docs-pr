
# Document.goToByIdAsync method
Goes to the specified object or location in the document.

|||
|:-----|:-----|
|**Hosts:**|Excel, PowerPoint, Word|
|**Available in requirement sets**|Not in a set|
|**Added in**|1.1|

[![Try out this call in the interactive API Tutorial for Excel](../../images/819b84bf-151c-4a12-80c3-d6f8d7c03251.png)](http://officeapitutorial.azurewebsites.net/Redirect.html?scenario=Navigate+to+Binding)


```js
Office.context.document.goToByIdAsync(id, goToType, [,options], callback);
```


## Parameters



|**Name**|**Type**|**Description**|**Support notes**|
|:-----|:-----|:-----|:-----|
| _id_|**string** or **number**|The identifier of the object or location to go to. Required.||
| _goToType_|[GoToType](../../reference/shared/gototype-enumeration.md)|The type of the location to go to. Required.||
| _options_|**object**|Specifies any of the following [optional parameters](../../docs/develop/asynchronous-programming-in-office-add-ins.md#passing-optional-parameters-to-asynchronous-methods)||
| _selectionMode_|[SelectionMode](../../reference/shared/selectionmode-enumeration.md)|Specifies whether the location specified by the  _id_ parameter is selected (highlighted).|**In Excel:**<br/> **Office.SelectionMode.Selected** selects all content in the binding, or named item. <br/>**Office.SelectionMode.None** for text bindings, moves the cursor to the beginning of the text; for matrix bindings, table bindings, and named items, selects the first data cell (not first cell in header row for tables).<br/><br/> **In PowerPoint:**<br/> **Office.SelectionMode.Selected** selects the slide title or first textbox on the slide.<br/> **Office.SelectionMode.None** Doesn't select anything.<br/><br/> **In Word:**<br/> **Office.SelectionMode.Selected** selects all content in the binding. <br/>**Office.SelectionMode.None** for text bindings, moves the cursor to the beginning of the text; for matrix bindings and table bindings, selects the first data cell (not first cell in header row for tables).|
| _asyncContext_|**array**,  **boolean**,  **null**,  **number**,  **object**, **string**, or  **undefined**|A user-defined item of any type that is returned in the  **AsyncResult** object without being altered.||
| _callback_|**object**|A function that is invoked when the callback returns, whose only parameter is of type  **AsyncResult**.||

## Callback Value

When the function you passed to the  _callback_ parameter executes, it receives an [AsyncResult](../../reference/shared/asyncresult.md) object that you can access from the callback function's only parameter.

In the callback function passed to the  **goToByIdAsync** method, you can use the properties of the **AsyncResult** object to return the following information.



|**Property**|**Use to...**|
|:-----|:-----|
|[AsyncResult.value](../../reference/shared/asyncresult.value.md)|Return the current view.|
|[AsyncResult.status](../../reference/shared/asyncresult.status.md)|Determine the success or failure of the operation.|
|[AsyncResult.error](../../reference/shared/asyncresult.error.md)|Access an [Error](../../reference/shared/error.md) object that provides error information if the operation failed.|
|[AsyncResult.asyncContext](../../reference/shared/asyncresult.asynccontext.md)|Access your user-defined  **object** or value, if you passed one as the _asyncContext_ parameter.|

## Remarks

PowerPoint doesn't support the  **goToByIdAsync** method in **Master Views**.


## Example

 **Go to a binding by id (Word and Excel)**

The following example shows how to:


-  **Create a table binding** using the [addFromSelectionAsync](../../reference/shared/bindings.addfromselectionasync.md) method as a sample binding to work with.
    
-  **Specify that binding** as the binding to go to.
    
-  **Pass an anonymous callback function** that returns the status of the operation to the _callback_ parameter of the **goToByIdAsync** method.
    
-  **Display the value** on the add-in's page.
    



```js
function gotoBinding() {
    //Create a new table binding for the selected table.
    Office.context.document.bindings.addFromSelectionAsync("table",{ id: "MyTableBinding" }, function (asyncResult) {
    if (asyncResult.status == "failed") {
              showMessage("Action failed with error: " + asyncResult.error.message);
           }
           else {
              showMessage("Added new binding with type: " + asyncResult.value.type +" and id: " + asyncResult.value.id);
           }
    });

    //Go to binding by id.
    Office.context.document.goToByIdAsync("MyTableBinding", Office.GoToType.Binding, function (asyncResult) {
        if (asyncResult.status == "failed") {
            showMessage("Action failed with error: " + asyncResult.error.message);
        }
        else {
            showMessage("Navigation successful");
        }
    });
}
```



 **Go to a table in a spreadsheet (Excel)**

The following example shows how to:


-  **Specify a table by name** as the table to go to.
    
-  **Pass an anonymous callback function** that returns the status of the operation to the _callback_ parameter of the **goToByIdAsync** method.
    
-  **Display the value** on the add-in's page.
    



```js
function goToTable() {
    Office.context.document.goToByIdAsync("Table1", Office.GoToType.NamedItem, function (asyncResult) {
        if (asyncResult.status == "failed") {
            showMessage("Action failed with error: " + asyncResult.error.message);
        }
        else {
            showMessage("Navigation successful");
        }
    });
}
```



 **Go to the currently selected slide by id (PowerPoint)**

The following example shows how to:


-  **Get the id** of the currently selected slides using the [getSelectedDataAsync](../../reference/shared/document.getselecteddataasync.md) method.
    
-  **Specify the returned id** as the slide to go to.
    
-  **Pass an anonymous callback function** that returns the status of the operation to the _callback_ parameter of the **goToByIdAsync** method.
    
-  **Display the value** of the stringified JSON object returned by `asyncResult.value`, which contains information about the selected slides, on the add-in's page.
    



```js
var firstSlideId = 0;
function gotoSelectedSlide() {
    //Get currently selected slide's id
    Office.context.document.getSelectedDataAsync(Office.CoercionType.SlideRange, function (asyncResult) {
        if (asyncResult.status == "failed") {
            app.showNotification("Action failed with error: " + asyncResult.error.message);
        }
        else {
            firstSlideId = asyncResult.value.slides[0].id;
            app.showNotification(JSON.stringify(asyncResult.value));
        }
    });
    //Go to slide by id.
    Office.context.document.goToByIdAsync(firstSlideId, Office.GoToType.Slide, function (asyncResult) {
        if (asyncResult.status == "failed") {
            app.showNotification("Action failed with error: " + asyncResult.error.message);
        }
        else {
            app.showNotification("Navigation successful");
        }
    });
}
```



 **Go to slide by index (PowerPoint)**

The following example shows how to:


-  **Specify the index** of the first, last, previous, or next slide to go to.
    
-  **Pass an anonymous callback function** that returns the status of the operation to the _callback_ parameter of the **goToByIdAsync** method.
    
-  **Display the value** on the add-in's page.
    



```js
function goToSlideByIndex() {
    var goToFirst = Office.Index.First;
    var goToLast = Office.Index.Last;
    var goToPrevious = Office.Index.Previous;
    var goToNext = Office.Index.Next;

    Office.context.document.goToByIdAsync(goToNext, Office.GoToType.Index, function (asyncResult) {
        if (asyncResult.status == "failed") {
            showMessage("Action failed with error: " + asyncResult.error.message);
        }
        else {
            showMessage("Navigation successful");
        }
    });
}
```




## Support details


A capital Y in the following matrix indicates that this method is supported in the corresponding Office host application. An empty cell indicates that the Office host application doesn't support this method.

For more information about Office host application and server requirements, see [Requirements for running Office Add-ins](../../docs/overview/requirements-for-running-office-add-ins.md).


**Supported hosts, by platform**


||**Office for Windows desktop**|**Office Online (in browser)**|**Office for iPad**|
|:-----|:-----|:-----|:-----|
|**Excel**|Y|Y|Y|
|**PowerPoint**|Y|Y|Y|
|**Word**|Y||Y|

|||
|:-----|:-----|
|**Available in requirement sets**|Not in a set|
|**Minimum permission level**|[ReadDocument](../../docs/develop/requesting-permissions-for-api-use-in-content-and-task-pane-add-ins.md)|
|**Add-in types**|Content, task pane|
|**Library**|Office.js|
|**Namespace**|Office|

## Support history



|**Version**|**Changes**|
|:-----|:-----|
|1.1|Added support for PowerPoint Online.|
|1.1|Added support for Excel, PowerPoint, and Word in Office for iPad.|
|1.1|Introduced|
