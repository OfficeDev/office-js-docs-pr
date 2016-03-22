
# Read and write data to the active selection in a document or spreadsheet

The [Document](../../reference/shared/document.md) object exposes methods that let you to read and write to the user's current selection in a document or spreadsheet. To do that, the **Document** object provides the **getSelectedDataAsync** and **setSelectedDataAsync** methods. This topic also describes how to read, write, and create event handlers to detect changes to the user's selection.

The  **getSelectedDataAsync** method only works against the user's current selection. If you need to persist the selection in the document, so that the same selection is available to read and write to across sessions of running your add-in, you must add abinding using the [Bindings.addFromSelectionAsync](http://msdn.microsoft.com/en-us/library/edc99214-e63e-43f2-9392-97ead42fc155.aspx) method (or create a binding with one of the other "addFrom" methods of the [Bindings](http://msdn.microsoft.com/en-us/library/09979e31-3bfb-45be-adda-0f7cc2db1fe1.aspx) object). For information about creating a binding to a region of a document, and then reading and writing to a binding, see [Bind to regions in a document or spreadsheet](../../docs/develop/bind-to-regions-in-a-document-or-spreadsheet.md).


### Read selected data


The following example shows how to get data from a selection in a document by using the [getSelectedDataAsync](../../reference/shared/document.getselecteddataasync.md) method.


```js
Office.context.document.getSelectedDataAsync(Office.CoercionType.Text, function (asyncResult) {
    if (asyncResult.status == Office.AsyncResultStatus.Failed) {
        write('Action failed. Error: ' + asyncResult.error.message);
    }
    else {
        write('Selected data: ' + asyncResult.value);
    }
});

// Function that writes to a div with id='message' on the page.
function write(message){
    document.getElementById('message').innerText += message; 
}
```

In this example, the first  _coercionType_ parameter is specified as **Office.CoercionType.Text** (you can also specify this parameter by using the literal string `"text"`). This means that the [value](../../reference/shared/asyncresult.status.md) property of the [AsyncResult](../../reference/shared/asyncresult.md) object that is available from the _asyncResult_ parameter in the callback function will return a **string** that contains the selected text in the document. Specifying different coercion types will result in different values. [Office.CoercionType](../../reference/shared/coerciontype-enumeration.md) is an enumeration of available coercion type values. **Office.CoercionType.Text** evaluates to the string "text".


 >**Tip**   **When should you use the matrix versus table coercionType for data access?** If you need your selected tabular data to grow dynamically when rows and columns are added, and you must work with table headers, you should use the table data type (by specifying the _coercionType_ parameter of the **getSelectedDataAsync** method as `"table"` or **Office.CoercionType.Table**). Adding rows and columns within the data structure is supported in both table and matrix data, but appending rows and columns is supported only for table data. If you are you aren't planning on adding rows and columns, and your data doesn't require header functionality, then you should use the matrix data type (by specifying the  _coercionType_ parameter of **getSelecteDataAsync** method as `"matrix"` or **Office.CoercionType.Matrix**), which provides a simpler model of interacting with the data.

The anonymous function that is passed into the function as the second  _callback_ parameter is executed when the **getSelectedDataAsync** operation is completed. The function is called with a single parameter, _asyncResult_, which contains the result and the status of the call. If the call fails, the [error](../../reference/shared/asyncresult.context.md) property of the **AsyncResult** object provides access to the [Error](../../reference/shared/error.md) object. You can check the value of the [Error.name](../../reference/shared/error.name.md) and [Error.message](../../reference/shared/error.message.md) properties to determine why the set operation failed. Otherwise, the selected text in the document is displayed.

The [AsyncResult.status](../../reference/shared/asyncresult.error.md) property is used in the **if** statement to test whether the call succeeded. [Office.AsyncResultStatus](../../reference/shared/asyncresultstatus-enumeration.md) is an enumeration of available **AsyncResult.status** property values. **Office.AsyncResultStatus.Failed** evaluates to the string "failed" (and, again, can also be specified as that literal string).


### Write data to the selection


The following example shows how to set the selection to show "Hello World!".


```js
Office.context.document.setSelectedDataAsync("Hello World!", function (asyncResult) {
    if (asyncResult.status == Office.AsyncResultStatus.Failed) {
        write(asyncResult.error.message);
    }
});

// Function that writes to a div with id='message' on the page.
function write(message){
    document.getElementById('message').innerText += message; 
}
```

Passing in different object types for the  _data_ parameter will have different results. The result depends on what is currently selected in the document, which application is hosting your add-in, and whether the data passed in can be coerced to the current selection.

The anonymous function passed into the [setSelectedDataAsync](../../reference/shared/document.setselecteddataasync.md) method as the _callback_ parameter is executed when the asynchronous call is completed. When you write data to the selection by using the **setSelectedDataAsync** method, the _asyncResult_ parameter of the callback provides access only to the status of the call, and to the [Error](../../reference/shared/error.md) object if the call fails.

 **Note:** Starting with the release of the Excel 2013 SP1 and the corresponding build of Excel Online, you can now [set formatting when writing a table to the current selection](../../docs/excel/format-tables-in-add-ins-for-excel.md).


### Detect changes in the selection


The following example shows how to detect changes in the selection by using the [Document.addHandlerAsync](../../reference/shared/document.addhandlerasync.md) method to add an event handler for the [SelectionChanged](../../reference/shared/document.selectionchanged.event.md) event on the document.


```
Office.context.document.addHandlerAsync("documentSelectionChanged", myHandler, function(result){} 
);

// Event handler function.
function myHandler(eventArgs){
write('Document Selection Changed');
}

// Function that writes to a div with id='message' on the page.
function write(message){
    document.getElementById('message').innerText += message; 
}
```

The first  _eventType_ parameter specifies the name of the event to subscribe to. Passing the string `"documentSelectionChanged"` for this parameter is equivalent to passing the **Office.EventType.DocumentSelectionChanged** event type of the [Office.EventType](../../reference/shared/eventtype-enumeration.md) enumeration.

The  `myHander()` function that is passed into the function as the second _handler_ parameter is an event handler that is executed when the selection is changed on the document. The function is called with a single parameter, _eventArgs_, which will contain a reference to a [DocumentSelectionChangedEventArgs](../../reference/shared/document.selectionchangedeventargs.md) object when the asynchronous operation completes. You can use the [DocumentSelectionChangedEventArgs.document](../../reference/shared/document.selectionchangedeventargs.document.md) property to access the document that raised the event.


 >**Note**  You can add multiple event handlers for a given event by calling the  **addHandlerAsync** method again and passing in an additional event handler function for the _handler_ parameter. This will work correctly as long as the name of each event handler function is unique.


### Stop detecting changes in the selection


The following example shows how to stop listening to the [Document.SelectionChanged](../../reference/shared/document.selectionchanged.event.md) event by calling the [document.removeHandlerAsync](../../reference/shared/document.removehandlerasync.md) method.


```
Office.context.document.removeHandlerAsync("documentSelectionChanged", {handler:myHandler}, function(result){});
```

The  `myHandler` function name that is passed as the second _handler_ parameter specifies the event handler that will be removed from the **SelectionChanged** event.


 >**Important**  If the optional  _handler_ parameter is omitted when the **removeHandlerAsync** method is called, all event handlers for the specified _eventType_ will be removed.


## Additional resources


- [Read data from a binding](../../docs/develop/bind-to-regions-in-a-document-or-spreadsheet.md#BindRegions_Read)
    
- [Write data to a binding](../../docs/develop/bind-to-regions-in-a-document-or-spreadsheet.md#BindRegions_Write)
    
