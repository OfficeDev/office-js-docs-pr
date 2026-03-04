---
title: Bind to regions in a document or spreadsheet
description: Learn how to use binding to ensure consistent access to a specific region or element of a document or spreadsheet through an identifier.
ms.date: 07/29/2025
ms.localizationpriority: medium
---

# Bind to regions in a document or spreadsheet

[!include[information about the common API](../includes/alert-common-api-info.md)]

Bindings let your add-in consistently access specific regions of a document or spreadsheet. Think of a binding as a bookmark that remembers a specific location, even if users change their selection or navigate elsewhere in the document. Specifically, here are what bindings offer your add-in.

- **Access common data structures** across supported Office applications, such as tables, ranges, or text.
- **Read and write data** without requiring users to make a selection first.
- **Create persistent relationships** between your add-in and document data. Bindings are saved with the document and work across sessions.

To create a binding, call one of these [Bindings] object methods to associate a document region with a unique identifier: [addFromPromptAsync], [addFromSelectionAsync], or [addFromNamedItemAsync]. Once you've established the binding, use its identifier to read from or write to that region anytime.

You can also subscribe to data and selection change events for specific bound regions. This means your add-in only gets notified about changes within the bound area, not the entire document.

## Choose the right binding type

Office supports [three different types of bindings][Office.BindingType]. You specify the type with the _bindingType_ parameter when creating a binding using [addFromSelectionAsync], [addFromPromptAsync], or [addFromNamedItemAsync].

### Text Binding

**[Text Binding][TextBinding]** - Binds to a document region that can be represented as text.

In Word, most contiguous selections work. In Excel, only single cell selections can use text binding. Excel supports only plain text, while Word supports three formats: plain text, HTML, and Open XML for Office.

### Matrix Binding  

**[Matrix Binding][MatrixBinding]** - Binds to a fixed region containing tabular data without headers.

Data in a matrix binding is read or written as a two-dimensional **Array** (an array of arrays in JavaScript). For example, two rows of **string** values in two columns would look like `[['a', 'b'], ['c', 'd']]`, and a single column of three rows would be `[['a'], ['b'], ['c']]`.

In Excel, any contiguous selection of cells works for matrix binding. In Word, only tables support matrix binding.

### Table Binding

**[Table Binding][TableBinding]** - Binds to a document region containing a table with headers.

Data in a table binding is read or written as a [TableData](/javascript/api/office/office.tabledata) object. The `TableData` object exposes data through the `headers` and `rows` properties.

Any Excel or Word table can be the basis for a table binding. After you establish a table binding, new rows or columns that users add to the table are automatically included in the binding.

After creating a binding with one of the three "addFrom" methods, you can work with the binding's data and properties using the corresponding object: [MatrixBinding], [TableBinding], or [TextBinding]. All three objects inherit the [getDataAsync] and [setDataAsync] methods from the `Binding` object for interacting with bound data.

> [!NOTE]
> **Should you use matrix or table bindings?**
> When working with tabular data that includes a total row, use matrix binding if your add-in needs to access values in the total row or detect when a user selects the total row. Table bindings don't include total rows in their [TableBinding.rowCount] property or in the `rowCount` and `startRow` properties of [BindingSelectionChangedEventArgs] in event handlers. To work with total rows, you must use matrix binding.

## Create a binding from the current selection

The following example adds a text binding called `myBinding` to the current selection using the [addFromSelectionAsync] method.

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

In this example, the binding type is text, so a [TextBinding] is created for the selection. Different binding types expose different data and operations. [Office.BindingType] is an enumeration of available binding types.

The second optional parameter specifies the ID of the new binding. If you don't specify an ID, one is generated automatically.

The anonymous function passed as the final _callback_ parameter runs when the binding creation is complete. The function receives a single parameter, `asyncResult`, which provides access to an [AsyncResult] object with the call's status. The `AsyncResult.value` property contains a reference to a [Binding] object of the specified type for the newly created binding. You can use this [Binding] object to get and set data.

## Create a binding from a prompt

The following function adds a text binding called `myBinding` using the [addFromPromptAsync] method. This method lets users specify the range for the binding using the application's built-in range selection prompt.

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

In this example, the binding type is text, so a [TextBinding] is created for the user's selection in the prompt.

The second parameter contains the ID of the new binding. If you don't specify an ID, one is generated automatically.

The anonymous function passed as the third _callback_ parameter runs when the binding creation is complete. When the callback function runs, the [AsyncResult] object contains the call's status and the newly created binding.

The following screenshot shows the built-in range selection prompt in Excel.

:::image type="content" source="../images/agave-api-overview-excel-selection-ui.png" alt-text="The Select Data dialog.":::

## Add a binding to a named item

The following function adds a binding to the existing `myRange` named item as a "matrix" binding using the [addFromNamedItemAsync] method and assigns the binding's `id` as "myMatrix".

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

**For Excel**, the `itemName` parameter of [addFromNamedItemAsync] refers to an existing named range, a range specified with A1 reference style (`"A1:A3"`), or a table. By default, Excel assigns the names "Table1" for the first table, "Table2" for the second table, and so on. To assign a meaningful name to a table in the Excel UI, use the **Table Name** property on the **Table Tools | Design** tab.

> [!NOTE]
> In Excel, when specifying a table as a named item, you must fully qualify the name to include the worksheet name in this format (e.g., `"Sheet1!Table1"`).

The following function creates a binding in Excel to the first three cells in column A (`"A1:A3"`), assigns the ID `"MyCities"`, and then writes three city names to that binding.

```js
 function bindingFromA1Range() {
    Office.context.document.bindings.addFromNamedItemAsync("A1:A3", "matrix", { id: "MyCities" },
        function (asyncResult) {
            if (asyncResult.status == "failed") {
                write('Error: ' + asyncResult.error.message);
            } else {
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

**For Word**, the `itemName` parameter of [addFromNamedItemAsync] refers to the `Title` property of a `Rich Text` content control. (You can't bind to content controls other than the `Rich Text` content control.)

By default, a content control has no `Title` value assigned. To assign a meaningful name in the Word UI, after inserting a **Rich Text** content control from the **Controls** group on the **Developer** tab, use the **Properties** command in the **Controls** group to display the **Content Control Properties** dialog. Then set the `Title` property of the content control to the name you want to reference from your code.

The following function creates a text binding in Word to a rich text content control named `"FirstName"`, assigns the **id** `"firstName"`, and then displays that information.

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

The following example gets all bindings in a document using the [getAllAsync] method.

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

The anonymous function passed as the `callback` parameter runs when the operation is complete. The function is called with a single parameter, `asyncResult`, which contains an array of the bindings in the document. The array is iterated to build a string that contains the IDs of the bindings. The string is then displayed in a message box.

## Get a binding by ID using getByIdAsync

The following example uses the [getByIdAsync] method to get a binding in a document by specifying its ID. This example assumes that a binding named `'myBinding'` was added to the document using one of the methods described earlier in this article.

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

In this example, the first `id` parameter is the ID of the binding to retrieve.

The anonymous function passed as the second _callback_ parameter runs when the operation is completed. The function is called with a single parameter, _asyncResult_, which contains the call's status and the binding with the ID "myBinding".

## Get a binding by ID using `Office.select`

The following example uses the [Office.select] function to get a [Binding] object promise in a document by specifying its ID in a selector string. It then calls the [getDataAsync] method to get data from the specified binding. This example assumes that a binding named `'myBinding'` was added to the document using one of the methods described earlier in this article.

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

If the `select` function promise successfully returns a [Binding] object, that object exposes only the following four methods: [getDataAsync], [setDataAsync], [addHandlerAsync], and [removeHandlerAsync]. If the promise can't return a Binding object, the `onError` callback can be used to access an [asyncResult].error object to get more information. If you need to call a member of the Binding object other than the four methods exposed by the [Binding] object promise returned by the `select` function, instead use the [getByIdAsync] method by using the [Document.bindings] property and [getByIdAsync] method to retrieve the [Binding] object.

## Release a binding by ID

The following example uses the [releaseByIdAsync] method to release a binding in a document by specifying its ID.

```js
Office.context.document.bindings.releaseByIdAsync('myBinding', function (asyncResult) {
    write('Released myBinding!');
});

// Function that writes to a div with id='message' on the page.
function write(message){
    document.getElementById('message').innerText += message; 
}
```

In this example, the first `id` parameter is the ID of the binding to release.

The anonymous function passed as the second parameter is a callback that runs when the operation is complete. The function is called with a single parameter, [asyncResult], which contains the call's status.

## Read data from a binding

The following example uses the [getDataAsync] method to get data from an existing binding.

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

`myBinding` is a variable that contains an existing text binding in the document. Alternatively, you could use [Office.select] to access the binding by its ID, and start your call to the [getDataAsync] method, like this:

```js
Office.select("bindings#myBindingID").getDataAsync
```

The anonymous function passed into the method is a callback that runs when the operation is complete. The [AsyncResult].value property contains the data within `myBinding`. The type of the value depends on the binding type. The binding in this example is a text binding, so the value will contain a string. For additional examples of working with matrix and table bindings, see the [getDataAsync] method topic.

## Write data to a binding

The following example uses the [setDataAsync] method to set data in an existing binding.

```js
myBinding.setDataAsync('Hello World!', function (asyncResult) { });
```

`myBinding` is a variable that contains an existing text binding in the document.

In this example, the first parameter is the value to set on `myBinding`. Because this is a text binding, the value is a `string`. Different binding types accept different types of data.

The anonymous function passed into the method is a callback that runs when the operation is complete. The function is called with a single parameter, `asyncResult`, which contains the result's status.

## Detect changes to data or selection in a binding

The following function attaches an event handler to the [DataChanged](/javascript/api/office/office.binding) event of a binding with an ID of "MyBinding".

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

The first _eventType_ parameter of [addHandlerAsync] specifies the name of the event to subscribe to. [Office.EventType] is an enumeration of available event type values. `Office.EventType.BindingDataChanged` evaluates to the string "bindingDataChanged".

The `dataChanged` function passed as the second _handler_ parameter is an event handler that runs when the data in the binding is changed. The function is called with a single parameter, _eventArgs_, which contains a reference to the binding. This binding can be used to retrieve the updated data.

Similarly, you can detect when a user changes selection in a binding by attaching an event handler to the [SelectionChanged] event of a binding. To do that, specify the `eventType` parameter of [addHandlerAsync] as `Office.EventType.BindingSelectionChanged` or `"bindingSelectionChanged"`.

You can add multiple event handlers for a given event by calling [addHandlerAsync] again and passing in an additional event handler function for the `handler` parameter. The name of each event handler function must be unique.

### Remove an event handler

To remove an event handler for an event, call [removeHandlerAsync] passing in the event type as the first _eventType_ parameter, and the name of the event handler function to remove as the second _handler_ parameter. For example, the following function removes the `dataChanged` event handler function added in the previous section's example.

```js
function removeEventHandlerFromBinding() {
    Office.select("bindings#MyBinding").removeHandlerAsync(
        Office.EventType.BindingDataChanged, {handler:dataChanged});
}
```

> [!IMPORTANT]
> If the optional _handler_ parameter is omitted when [removeHandlerAsync] is called, all event handlers for the specified `eventType` will be removed.

## See also

- [Understanding the Office JavaScript API](understanding-the-javascript-api-for-office.md)
- [Asynchronous programming in Office Add-ins](asynchronous-programming-in-office-add-ins.md)
- [Read and write data to the active selection in a document or spreadsheet](read-and-write-data-to-the-active-selection-in-a-document-or-spreadsheet.md)

[Binding]: /javascript/api/office/office.binding
[MatrixBinding]: /javascript/api/office/office.matrixbinding
[TableBinding]: /javascript/api/office/office.tablebinding
[TextBinding]: /javascript/api/office/office.textbinding
[getDataAsync]: /javascript/api/office/office.binding#getDataAsync_options__callback_
[setDataAsync]: /javascript/api/office/office.binding#setDataAsync_data__options__callback_
[SelectionChanged]: /javascript/api/office/office.bindingselectionchangedeventargs
[addHandlerAsync]: /javascript/api/office/office.binding#addHandlerAsync_eventType__handler__options__callback_
[removeHandlerAsync]: /javascript/api/office/office.binding#removeHandlerAsync_eventType__options__callback_

[Bindings]: /javascript/api/office/office.bindings
[getByIdAsync]: /javascript/api/office/office.bindings#getByIdAsync_id__options__callback_
[getAllAsync]: /javascript/api/office/office.bindings#getAllAsync_options__callback_
[addFromNamedItemAsync]: /javascript/api/office/office.bindings#addFromNamedItemAsync_itemName__bindingType__options__callback_
[addFromSelectionAsync]: /javascript/api/office/office.bindings#addFromSelectionAsync_bindingType__options__callback_
[addFromPromptAsync]: /javascript/api/office/office.bindings#addFromPromptAsync_bindingType__options__callback_
[releaseByIdAsync]: /javascript/api/office/office.bindings#releaseByIdAsync_id__options__callback_

[AsyncResult]: /javascript/api/office/office.asyncresult
[Office.BindingType]: /javascript/api/office/office.bindingtype
[Office.select]: /javascript/api/office
[Office.EventType]: /javascript/api/office/office.eventtype
[Document.bindings]: /javascript/api/office/office.document

[TableBinding.rowCount]: /javascript/api/office/office.tablebinding
[BindingSelectionChangedEventArgs]: /javascript/api/office/office.bindingselectionchangedeventargs
