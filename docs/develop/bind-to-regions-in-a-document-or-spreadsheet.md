---
title: Bind to regions in a document or spreadsheet
description: Learn how to use binding to ensure consistent access a specific region or element of a document or spreadsheet through an identifier.
ms.date: 03/21/2023
ms.localizationpriority: medium
---


# Bind to regions in a document or spreadsheet

Binding-based data access enables content and task pane add-ins to consistently access a particular region of a document or spreadsheet through an identifier. The add-in first needs to establish the binding by calling one of the methods that associates a portion of the document with a unique identifier: [addFromPromptAsync], [addFromSelectionAsync], or [addFromNamedItemAsync]. After the binding is established, the add-in can use the provided identifier to access the data contained in the associated region of the document or spreadsheet. Creating bindings provides the following value to your add-in.

- Permits access to common data structures across supported Office applications, such as: tables, ranges, or text (a contiguous run of characters).
- Enables read/write operations without requiring the user to make a selection.
- Establishes a relationship between the add-in and the data in the document. Bindings are persisted in the document, and can be accessed at a later time.

Establishing a binding also allows you to subscribe to data and selection change events that are scoped to that particular region of the document or spreadsheet. This means that the add-in is only notified of changes that happen within the bound region as opposed to general changes across the whole document or spreadsheet.

The [Bindings] object exposes a [getAllAsync] method that gives access to the set of all bindings established on the document or spreadsheet. An individual binding can be accessed by its ID using either the Bindings.[getByIdAsync] method or [Office.select] function. You can establish new bindings as well as remove existing ones by using one of the following methods of the [Bindings] object: [addFromSelectionAsync], [addFromPromptAsync], [addFromNamedItemAsync], or [releaseByIdAsync].

## Binding types

There are [three different types of bindings][Office.BindingType] that you specify with the _bindingType_ parameter when you create a binding with the [addFromSelectionAsync], [addFromPromptAsync] or [addFromNamedItemAsync] methods.

1. **[Text Binding][TextBinding]** - Binds to a region of the document that can be represented as text.

    In Word, most contiguous selections are valid, while in Excel only single cell selections can be the target of a text binding. In Excel, only plain text is supported. In Word, three formats are supported: plain text, HTML, and Open XML for Office.

1. **[Matrix Binding][MatrixBinding]** - Binds to a fixed region of a document that contains tabular data without headers.Data in a matrix binding is written or read as a two dimensional **Array**, which in JavaScript is implemented as an array of arrays. For example, two rows of **string** values in two columns can be written or read as `[['a', 'b'], ['c', 'd']]`, and a single column of three rows can be written or read as `[['a'], ['b'], ['c']]`.

    In Excel, any contiguous selection of cells can be used to establish a matrix binding. In Word, only tables support matrix binding.

1. **[Table Binding][TableBinding]** - Binds to a region of a document that contains a table with headers.Data in a table binding is written or read as a [TableData](/javascript/api/office/office.tabledata) object. The `TableData` object exposes the data through the `headers` and `rows` properties.

    Any Excel or Word table can be the basis for a table binding. After you establish a table binding, each new row or column a user adds to the table is automatically included in the binding.

After a binding is created by using one of the three "addFrom" methods of the `Bindings` object, you can work with the binding's data and properties by using the methods of the corresponding object: [MatrixBinding], [TableBinding], or [TextBinding]. All three of these objects inherit the [getDataAsync] and [setDataAsync] methods of the `Binding` object that enable you to interact with the bound data.

> [!NOTE]
> **When should you use matrix versus table bindings?**
> When the tabular data you are working with contains a total row, you must use a matrix binding if your add-in's script needs to access values in the total row or detect that the user's selection is in the total row. If you establish a table binding for tabular data that contains a total row, the [TableBinding.rowCount] property and the `rowCount` and `startRow` properties of the [BindingSelectionChangedEventArgs] object in event handlers won't reflect the total row in their values. To work around this limitation, you must use establish a matrix binding to work with the total row.

## Add a binding to the user's current selection

The following example shows how to add a text binding called  `myBinding` to the current selection in a document by using the [addFromSelectionAsync] method.

```js
Office.context.document.bindings.addFromSelectionAsync(Office.BindingType.Text, { id: 'myBinding' }, function (asyncResult) {
    if (asyncResult.status == Office.AsyncResultStatus.Failed) {
        write('Action failed. Error: ' + asyncResult.error.message);
    } else {
        write('Added new binding with type: ' + asyncResult.value.type + ' and id: ' + asyncResult.value.id);
    }
});

// Function that writes to a div with id='message' on the page.
function write(message){
    document.getElementById('message').innerText += message; 
}
```

In this example, the specified binding type is text. This means that a [TextBinding] will be created for the selection. Different binding types expose different data and operations. [Office.BindingType] is an enumeration of available binding type values.

The second optional parameter is an object that specifies the ID of the new binding being created. If an ID is not specified, one is generated automatically.

The anonymous function that is passed into the method as the final  _callback_ parameter is executed when the creation of the binding is complete. The function is called with a single parameter, `asyncResult`, which provides access to an [AsyncResult] object that provides the status of the call. The `AsyncResult.value` property contains a reference to a [Binding] object of the type that is specified for the newly created binding. You can use this [Binding] object to get and set data.

## Add a binding from a prompt

The following example shows how to add a text binding called  `myBinding` by using the [addFromPromptAsync] method. This method lets the user specify the range for the binding by using the application's built-in range selection prompt.

```js
function bindFromPrompt() {
    Office.context.document.bindings.addFromPromptAsync(Office.BindingType.Text, { id: 'myBinding' }, function (asyncResult) {
        if (asyncResult.status == Office.AsyncResultStatus.Failed) {
            write('Action failed. Error: ' + asyncResult.error.message);
        } else {
            write('Added new binding with type: ' + asyncResult.value.type + ' and id: ' + asyncResult.value.id);
        }
    });
}

// Function that writes to a div with id='message' on the page.
function write(message){
    document.getElementById('message').innerText += message;
}
```

In this example, the specified binding type is text. This means that a [TextBinding] will be created for the selection that the user specifies in the prompt.

The second parameter is an object that contains the ID of the new binding being created. If an ID is not specified, one is generated automatically.

The anonymous function passed into the method as the third _callback_ parameter is executed when the creation of the binding is complete. When the callback function executes, the [AsyncResult] object contains the status of the call and the newly created binding.

Figure 1 shows the built-in range selection prompt in Excel.

*Figure 1. Excel Select Data UI*

![The Select Data dialog.](../images/agave-api-overview-excel-selection-ui.png)

## Add a binding to a named item

The following example shows how to add a binding to the existing `myRange` named item as a "matrix" binding by using the [addFromNamedItemAsync] method, and assigns the binding's `id` as "myMatrix".

```js
function bindNamedItem() {
    Office.context.document.bindings.addFromNamedItemAsync("myRange", "matrix", {id:'myMatrix'}, function (result) {
        if (result.status == 'succeeded'){
            write('Added new binding with type: ' + result.value.type + ' and id: ' + result.value.id);
            }
        else
            write('Error: ' + result.error.message);
    });
}

// Function that writes to a div with id='message' on the page.
function write(message){
    document.getElementById('message').innerText += message; 
}

```

**For Excel**, the `itemName` parameter of the [addFromNamedItemAsync] method can refer to an existing named range, a range specified with the `A1` reference style `("A1:A3")`, or a table. By default, adding a table in Excel assigns the name "Table1" for the first table you add, "Table2" for the second table you add, and so on. To assign a meaningful name for a table in the Excel UI, use the `Table Name` property on the **Table Tools | Design** tab of the ribbon.

> [!NOTE]
> In Excel, when specifying a table as a named item, you must fully qualify the name to include the worksheet name in the name of the table in this format: `"Sheet1!Table1"`

The following example creates a binding in Excel to the first three cells in column A ( `"A1:A3"`), assigns the id `"MyCities"`, and then writes three city names to that binding.

```js
 function bindingFromA1Range() {
    Office.context.document.bindings.addFromNamedItemAsync("A1:A3", "matrix", {id: "MyCities" },
        function (asyncResult) {
            if (asyncResult.status == "failed") {
                write('Error: ' + asyncResult.error.message);
            }
            else {
                // Write data to the new binding.
                Office.select("bindings#MyCities").setDataAsync([['Berlin'], ['Munich'], ['Duisburg']], { coercionType: "matrix" },
                    function (asyncResult) {
                        if (asyncResult.status == "failed") {
                            write('Error: ' + asyncResult.error.message);
                        }
                    });
            }
        });
}
// Function that writes to a div with id='message' on the page.
function write(message){
    document.getElementById('message').innerText += message; 
}
```

**For Word**, the `itemName` parameter of the [addFromNamedItemAsync] method refers to the `Title` property of a `Rich Text` content control. (You can't bind to content controls other than the `Rich Text` content control.)

By default, a content control has no `Title*`value assigned. To assign a meaningful name in the Word UI, after inserting a **Rich Text** content control from the **Controls** group on the **Developer** tab of the ribbon, use the **Properties** command in the **Controls** group to display the **Content Control Properties** dialog box. Then set the `Title` property of the content control to the name you want to reference from your code.

The following example creates a text binding in Word to a rich text content control named `"FirstName"`, assigns the **id** `"firstName"`, and then displays that information.

```js
function bindContentControl() {
    Office.context.document.bindings.addFromNamedItemAsync('FirstName', 
        Office.BindingType.Text, {id:'firstName'},
        function (result) {
            if (result.status === Office.AsyncResultStatus.Succeeded) {
                write('Control bound. Binding.id: '
                    + result.value.id + ' Binding.type: ' + result.value.type);
            } else {
                write('Error:', result.error.message);
            }
    });
}
// Function that writes to a div with id='message' on the page.
function write(message){
    document.getElementById('message').innerText += message; 
}
```

## Get all bindings

The following example shows how to get all bindings in a document by using the Bindings.[getAllAsync] method.

```js
Office.context.document.bindings.getAllAsync(function (asyncResult) {
    let bindingString = '';
    for (let i in asyncResult.value) {
        bindingString += asyncResult.value[i].id + '\n';
    }
    write('Existing bindings: ' + bindingString);
});

// Function that writes to a div with id='message' on the page.
function write(message){
    document.getElementById('message').innerText += message; 
}
```

The anonymous function that is passed into the method as the `callback` parameter is executed when the operation is complete. The function is called with a single parameter, `asyncResult`, which contains an array of the bindings in the document. The array is iterated to build a string that contains the IDs of the bindings. The string is then displayed in a message box.

## Get a binding by ID using the getByIdAsync method of the Bindings object

The following example shows how to use the [getByIdAsync] method to get a binding in a document by specifying its ID. This example assumes that a binding named `'myBinding'` was added to the document using one of the methods described earlier in this topic.

```js
Office.context.document.bindings.getByIdAsync('myBinding', function (asyncResult) {
    if (asyncResult.status == Office.AsyncResultStatus.Failed) {
        write('Action failed. Error: ' + asyncResult.error.message);
    }
    else {
        write('Retrieved binding with type: ' + asyncResult.value.type + ' and id: ' + asyncResult.value.id);
    }
});

// Function that writes to a div with id='message' on the page.
function write(message){
    document.getElementById('message').innerText += message; 
}
```

In the example, the first `id` parameter is the ID of the binding to retrieve.

The anonymous function that is passed into the method as the second _callback_ parameter is executed when the operation is completed. The function is called with a single parameter, _asyncResult_, which contains the status of the call and the binding with the ID "myBinding".

## Get a binding by ID using the select function of the Office object

The following example shows how to use the [Office.select] function to get a [Binding] object promise in a document by specifying its ID in a selector string. It then calls the Binding.[getDataAsync] method to get data from the specified binding. This example assumes that a binding named `'myBinding'` was added to the document using one of the methods described earlier in this topic.

```js
Office.select("bindings#myBinding", function onError(){}).getDataAsync(function (asyncResult) {
    if (asyncResult.status == Office.AsyncResultStatus.Failed) {
        write('Action failed. Error: ' + asyncResult.error.message);
    } else {
        write(asyncResult.value);
    }
});

// Function that writes to a div with id='message' on the page.
function write(message){
    document.getElementById('message').innerText += message; 
}
```

> [!NOTE]
> If the `select` function promise successfully returns a [Binding] object, that object exposes only the following four methods of the object: [getDataAsync], [setDataAsync], [addHandlerAsync], and [removeHandlerAsync]. If the promise cannot return a Binding object, the `onError` callback can be used to access an [asyncResult].error object to get more information. If you need to call a member of the Binding object other than the four methods exposed by the [Binding] object promise returned by the `select` function, instead use the [getByIdAsync] method by using the [Document.bindings] property and Bindings.[getByIdAsync] method to retrieve the [Binding] object.

## Release a binding by ID

The following example shows how use the [releaseByIdAsync] method to release a binding in a document by specifying its ID.

```js
Office.context.document.bindings.releaseByIdAsync('myBinding', function (asyncResult) {
    write('Released myBinding!');
});

// Function that writes to a div with id='message' on the page.
function write(message){
    document.getElementById('message').innerText += message; 
}
```

In the example, the first `id` parameter is the ID of the binding to release.

The anonymous function that is passed into the method as the second parameter is a callback that is executed when the operation is complete. The function is called with a single parameter,  [asyncResult], which contains the status of the call.

## Read data from a binding

The following example shows how to use the [getDataAsync] method to get data from an existing binding.

```js
myBinding.getDataAsync(function (asyncResult) {
    if (asyncResult.status == Office.AsyncResultStatus.Failed) {
        write('Action failed. Error: ' + asyncResult.error.message);
    } else {
        write(asyncResult.value);
    }
});

// Function that writes to a div with id='message' on the page.
function write(message){
    document.getElementById('message').innerText += message; 
}
```

`myBinding` is a variable that contains an existing text binding in the document. Alternatively, you could use the [Office.select] to access the binding by its ID, and start your call to the [getDataAsync] method, like this:

```js
Office.select("bindings#myBindingID").getDataAsync
```

The anonymous function that is passed into the method is a callback that is executed when the operation is complete. The [AsyncResult].value property contains the data within `myBinding`. The type of the value depends on the binding type. The binding in this example is a text binding. Therefore, the value will contain a string. For additional examples of working with matrix and table bindings, see the [getDataAsync] method topic.

## Write data to a binding

The following example shows how to use the [setDataAsync] method to set data in an existing binding.

```js
myBinding.setDataAsync('Hello World!', function (asyncResult) { });
```

`myBinding` is a variable that contains an existing text binding in the document.

In the example, the first parameter is the value to set on `myBinding`. Because this is a text binding, the value is a `string`. Different binding types accept different types of data.

The anonymous function that is passed into the method is a callback that is executed when the operation is complete. The function is called with a single parameter, `asyncResult`, which contains the status of the result.

## Detect changes to data or the selection in a binding

The following example shows how to attach an event handler to the [DataChanged](/javascript/api/office/office.binding) event of a binding with an id of "MyBinding".

```js
function addHandler() {
Office.select("bindings#MyBinding").addHandlerAsync(
    Office.EventType.BindingDataChanged, dataChanged);
}
function dataChanged(eventArgs) {
    write('Bound data changed in binding: ' + eventArgs.binding.id);
}
// Function that writes to a div with id='message' on the page.
function write(message){
    document.getElementById('message').innerText += message;
}
```

The `myBinding` is a variable that contains an existing text binding in the document.

The first _eventType_ parameter of the [addHandlerAsync] method specifies the name of the event to subscribe to. [Office.EventType] is an enumeration of available event type values. `Office.EventType.BindingDataChanged` evaluates to the string "bindingDataChanged".

The `dataChanged` function that is passed into the method as the second _handler_ parameter is an event handler that is executed when the data in the binding is changed. The function is called with a single parameter, _eventArgs_, which contains a reference to the binding. This binding can be used to retrieve the updated data.

Similarly, you can detect when a user changes selection in a binding by attaching an event handler to the [SelectionChanged] event of a binding. To do that, specify the `eventType` parameter of the [addHandlerAsync] method as `Office.EventType.BindingSelectionChanged` or `"bindingSelectionChanged"`.

You can add multiple event handlers for a given event by calling the [addHandlerAsync] method again and passing in an additional event handler function for the `handler` parameter. This will work correctly as long as the name of each event handler function is unique.

### Remove an event handler

To remove an event handler for an event, call the [removeHandlerAsync] method passing in the event type as the first _eventType_ parameter, and the name of the event handler function to remove as the second _handler_ parameter. For example, the following function will remove the `dataChanged` event handler function added in the previous section's example.

```js
function removeEventHandlerFromBinding() {
    Office.select("bindings#MyBinding").removeHandlerAsync(
        Office.EventType.BindingDataChanged, {handler:dataChanged});
}
```

> [!IMPORTANT]
> If the optional _handler_ parameter is omitted when the [removeHandlerAsync] method is called, all event handlers for the specified `eventType` will be removed.

## See also

- [Understanding the Office JavaScript API](understanding-the-javascript-api-for-office.md)
- [Asynchronous programming in Office Add-ins](asynchronous-programming-in-office-add-ins.md)
- [Read and write data to the active selection in a document or spreadsheet](read-and-write-data-to-the-active-selection-in-a-document-or-spreadsheet.md)

[Binding]: 				 /javascript/api/office/office.binding
[MatrixBinding]:		 /javascript/api/office/office.matrixbinding
[TableBinding]: 		 /javascript/api/office/office.tablebinding
[TextBinding]: 			 /javascript/api/office/office.textbinding
[getDataAsync]: 		 /javascript/api/office/office.binding#getDataAsync_options__callback_
[setDataAsync]: 		 /javascript/api/office/office.binding#setDataAsync_data__options__callback_
[SelectionChanged]: 	 /javascript/api/office/office.bindingselectionchangedeventargs
[addHandlerAsync]: 		 /javascript/api/office/office.binding#addHandlerAsync_eventType__handler__options__callback_
[removeHandlerAsync]: 	 /javascript/api/office/office.binding#removeHandlerAsync_eventType__options__callback_

[Bindings]:				 /javascript/api/office/office.bindings
[getByIdAsync]: 		 /javascript/api/office/office.bindings#getByIdAsync_id__options__callback_
[getAllAsync]: 			 /javascript/api/office/office.bindings#getAllAsync_options__callback_
[addFromNamedItemAsync]: /javascript/api/office/office.bindings#addFromNamedItemAsync_itemName__bindingType__options__callback_
[addFromSelectionAsync]: /javascript/api/office/office.bindings#addFromSelectionAsync_bindingType__options__callback_
[addFromPromptAsync]: 	 /javascript/api/office/office.bindings#addFromPromptAsync_bindingType__options__callback_
[releaseByIdAsync]: 	 /javascript/api/office/office.bindings#releaseByIdAsync_id__options__callback_

[AsyncResult]: 			/javascript/api/office/office.asyncresult
[Office.BindingType]: 	/javascript/api/office/office.bindingtype
[Office.select]:		/javascript/api/office 
[Office.EventType]: 	/javascript/api/office/office.eventtype 
[Document.bindings]: 	/javascript/api/office/office.document

[TableBinding.rowCount]: /javascript/api/office/office.tablebinding
[BindingSelectionChangedEventArgs]: /javascript/api/office/office.bindingselectionchangedeventargs