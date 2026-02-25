---
title: Read and write data to the active selection in a document or spreadsheet
description: Learn how to read and write data to the active selection in a Word document or Excel spreadsheet.
ms.date: 02/25/2026
ms.localizationpriority: medium
---


# Read and write data to the active selection in a document or spreadsheet

[!include[information about the common API](../includes/alert-common-api-info.md)]

Learn how to read and write the user's current selection with `Document.getSelectedDataAsync` and `Document.setSelectedDataAsync`. Examples include a callback-based snippet and a modern Promise/async wrapper so you can use async/await.

- **When to use this**: Quick, temporary edits to the user's selection (clipboard-like), lightweight interactions, or immediate UI updates. Use a binding when you need persistence across sessions.
- **Applies to**: Word and Excel (behavior differs by host — see host notes later). For web versus desktop differences, test in supported platforms.
- **Key APIs**: `Office.context.document.getSelectedDataAsync`, `Office.context.document.setSelectedDataAsync`, `Office.EventType.DocumentSelectionChanged`.

The [Document](/javascript/api/office/office.document) object exposes methods that let you read and write to the user's current selection in a document or spreadsheet. To do that, the `Document` object provides the `getSelectedDataAsync` and `setSelectedDataAsync` methods. This topic also describes how to read, write, and create event handlers to detect changes to the user's selection.

The `getSelectedDataAsync` method only works against the user's current selection. If you need to persist the selection in the document, so that the same selection is available to read and write across sessions of running your add-in, you must add a binding using the [Bindings.addFromSelectionAsync](/javascript/api/office/office.bindings#office-office-bindings-addfromselectionasync-member(1)) method (or create a binding with one of the other "addFrom" methods of the [Bindings](/javascript/api/office/office.bindings) object). For information about creating a binding to a region of a document, and then reading and writing to a binding, see [Bind to regions in a document or spreadsheet](bind-to-regions-in-a-document-or-spreadsheet.md).

## Read selected data

The following example shows how to get data from a selection in a document by using the [getSelectedDataAsync](/javascript/api/office/office.document#office-office-document-getselecteddataasync-member(1)) method.

```js
Office.context.document.getSelectedDataAsync(Office.CoercionType.Text, function (asyncResult) {
    if (asyncResult.status === Office.AsyncResultStatus.Failed) {
        // Failure handling: Show the error message in the UI.
        write('Action failed. Error: ' + asyncResult.error.message);
    }
    else {
        // Success: `asyncResult.value` contains the selected text.
        write('Selected data: ' + asyncResult.value);
    }
});

// Function that writes to a div with id='message' on the page.
function write(message){
    document.getElementById('message').innerText += message; 
}
```

In this example, the first parameter, _coercionType_, is specified as `Office.CoercionType.Text` (you can also specify this parameter by using the literal string `"text"`). This means that the [value](/javascript/api/office/office.asyncresult#office-office-asyncresult-status-member) property of the [AsyncResult](/javascript/api/office/office.asyncresult) object that is available from the _asyncResult_ parameter in the callback function will return a **string** that contains the selected text in the document. Specifying different coercion types will result in different values. [Office.CoercionType](/javascript/api/office/office.coerciontype) is an enumeration of available coercion type values. `Office.CoercionType.Text` evaluates to the string "text".

**Expected output**: Writes the selected text to the page element with id `message`.

> [!TIP]
> **When should you use the matrix versus table coercionType for data access?** If you need your selected tabular data to grow dynamically when rows and columns are added, and you must work with table headers, you should use the table data type (by specifying the _coercionType_ parameter of the `getSelectedDataAsync` method as `"table"` or `Office.CoercionType.Table`). Adding rows and columns within the data structure is supported in both table and matrix data, but appending rows and columns is supported only for table data. If you aren't planning on adding rows and columns, and your data doesn't require header functionality, then you should use the matrix data type (by specifying the  _coercionType_ parameter of `getSelectedDataAsync` method as `"matrix"` or `Office.CoercionType.Matrix`), which provides a simpler model of interacting with the data.

The anonymous function that is passed into the method as the second parameter, _callback_, is executed when the `getSelectedDataAsync` operation is completed. The function is called with a single parameter, _asyncResult_, which contains the result and the status of the call. If the call fails, the [error](/javascript/api/office/office.asyncresult#office-office-asyncresult-error-member) property of the `AsyncResult` object provides access to the [Error](/javascript/api/office/office.error) object. You can check the value of the [Error.name](/javascript/api/office/office.error#office-office-error-name-member) and [Error.message](/javascript/api/office/office.error#office-office-error-message-member) properties to determine why the set operation failed. Otherwise, the selected text in the document is displayed.

The [AsyncResult.status](/javascript/api/office/office.asyncresult#office-office-asyncresult-error-member) property is used in the **if** statement to test whether the call succeeded. [Office.AsyncResultStatus](/javascript/api/office/office.asyncresult#office-office-asyncresult-status-member) is an enumeration of available `AsyncResult.status` property values. `Office.AsyncResultStatus.Failed` evaluates to the string "failed" (and, again, can also be specified as that literal string).

## Write data to the selection

The following example shows how to set the selection to show "Hello World!".

```js
Office.context.document.setSelectedDataAsync("Hello World!", function (asyncResult) {
    if (asyncResult.status === Office.AsyncResultStatus.Failed) {
        // Show the error message if the call fails.
        write(asyncResult.error.message);
    }
});

// Function that writes to a div with id='message' on the page.
function write(message){
    document.getElementById('message').innerText += message;
}
```

Passing in different object types for the  _data_ parameter will have different results. The result depends on what is currently selected in the document, which Office client application is hosting your add-in, and whether the data passed in can be coerced to the current selection.

The anonymous function passed into the [setSelectedDataAsync](/javascript/api/office/office.document#office-office-document-setselecteddataasync-member(1)) method as the _callback_ parameter is executed when the asynchronous call is completed. When you write data to the selection by using the `setSelectedDataAsync` method, the _asyncResult_ parameter of the callback provides access only to the status of the call, and to the [Error](/javascript/api/office/office.error) object if the call fails.

**Expected output**: The current selection in the document is replaced with the text "Hello World!" (if the selection and host support text insertion).

### Modern Promise / async wrapper

If you prefer async/await, use a small Promise wrapper around the callback APIs. Example wrappers:

```js
function getSelectedDataAsyncWithPromise(coercionType) {
    return new Promise((resolve, reject) => {
        Office.context.document.getSelectedDataAsync(coercionType, (result) => {
            if (result.status === Office.AsyncResultStatus.Failed) reject(result.error);
            else resolve(result.value);
        });
    });
}

function setSelectedDataAsyncWithPromise(data) {
    return new Promise((resolve, reject) => {
        Office.context.document.setSelectedDataAsync(data, (result) => {
            if (result.status === Office.AsyncResultStatus.Failed) reject(result.error);
            else resolve();
        });
    });
}

// Usage with async/await.
async function example() {
    try {
        const text = await getSelectedDataAsyncWithPromise(Office.CoercionType.Text);
        console.log('Selected:', text);
        await setSelectedDataAsyncWithPromise(text + ' (processed)');
    } catch (err) {
        console.error(err.message || err);
    }
}
```

For more information, see [Wrap Common APIs in promise-returning functions](asynchronous-programming-in-office-add-ins.md#wrap-common-apis-in-promise-returning-functions).

## Detect changes in the selection

The following example shows how to detect changes in the selection by using the [Document.addHandlerAsync](/javascript/api/office/office.document#office-office-document-addhandlerasync-member(1)) method to add an event handler for the [SelectionChanged](/javascript/api/office/office.documentselectionchangedeventargs) event on the document.

```js
Office.context.document.addHandlerAsync("documentSelectionChanged", myHandler, function(result){}
);

// Event handler function.
function myHandler(eventArgs){
    // `eventArgs` contains a `document` reference when available.
    write('Document Selection Changed');
}

// Function that writes to a div with id='message' on the page.
function write(message){
    document.getElementById('message').innerText += message;
}
```

The first parameter, _eventType_, specifies the name of the event to subscribe to. Passing the string `"documentSelectionChanged"` for this parameter is equivalent to passing the `Office.EventType.DocumentSelectionChanged` event type of the [Office.EventType](/javascript/api/office/office.eventtype) enumeration.

The  `myHandler()` function that is passed into the method as the second parameter, _handler_, is an event handler that is executed when the selection is changed on the document. The function is called with a single parameter, _eventArgs_, which will contain a reference to a [DocumentSelectionChangedEventArgs](/javascript/api/office/office.documentselectionchangedeventargs) object when the asynchronous operation completes. You can use the [DocumentSelectionChangedEventArgs.document](/javascript/api/office/office.documentselectionchangedeventargs#office-office-documentselectionchangedeventargs-document-member) property to access the document that raised the event.

**Host application notes**: In Excel, selection change events are often about workbook ranges (use matrix or table coercion types). In Word, selection events are text or content focused. Test event handlers in the hosts your add-in supports.

> [!NOTE]
> You can add multiple event handlers for a given event by calling the `addHandlerAsync` method again and passing in an additional event handler function for the _handler_ parameter. This will work correctly as long as the name of each event handler function is unique.

## Stop detecting changes in the selection

The following example shows how to stop listening to the [Document.SelectionChanged](/javascript/api/office/office.documentselectionchangedeventargs) event by calling the [document.removeHandlerAsync](/javascript/api/office/office.document#office-office-document-removehandlerasync-member(1)) method.

```js
Office.context.document.removeHandlerAsync("documentSelectionChanged", {handler:myHandler}, function(result){});
```

The  `myHandler` function name that is passed as the second parameter, _handler_, specifies the event handler that will be removed from the `SelectionChanged` event.

> [!IMPORTANT]
> If the optional  _handler_ parameter is omitted when the `removeHandlerAsync` method is called, all event handlers for the specified _eventType_ will be removed.

## Troubleshooting

- **Empty selection**: Nothing is returned if the user has no selection. Check and prompt the user to select content.
- **Coercion mismatch**: Use the right **coercionType** (`text`, `matrix`, `table`, `html`) for the data you expect.
- **Host limitations**: Some hosts may not allow certain coercions. See [Office.CoercionType](/javascript/api/office/office.coerciontype) for details.
- **Permissions**: Ensure your add-in's manifest has the appropriate permissions for the APIs you use.

## Next steps

For persistent data across sessions, see [Bind to regions in a document or spreadsheet](bind-to-regions-in-a-document-or-spreadsheet.md). Other useful topics: authentication, shared runtime, and application-specific examples.
