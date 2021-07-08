---
title: Office JavaScript API support for content and task pane add-ins in Office 2013
description: 'Use the Office JavaScript API to create a task pane in Office 2013.'
ms.date: 02/27/2020
localization_priority: Normal
---

# Office JavaScript API support for content and task pane add-ins in Office 2013

[!include[information about the common API](../includes/alert-common-api-info.md)]

You can use the [Office JavaScript API](../reference/javascript-api-for-office.md) to create task pane or content add-ins for Office 2013 client applications. The objects and methods that content and task pane add-ins support are categorized as follows:

1. **Common objects shared with other Office Add-ins.** These objects include [Office](/javascript/api/office), [Context](/javascript/api/office/office.context), and [AsyncResult](/javascript/api/office/office.asyncresult). The `Office` object is the root object of the Office JavaScript API. The `Context` object represents the add-in's runtime environment. Both `Office` and `Context` are the fundamental objects for any Office Add-in. The `AsyncResult` object represents the results of an asynchronous operation, such as the data returned to the `getSelectedDataAsync` method, which reads what a user has selected in a document.

2. **The Document object.** The majority of the API available to content and task pane add-ins is exposed through the methods, properties, and events of the [Document](/javascript/api/office/office.document) object. A content or task pane add-in can use the [Office.context.document](/javascript/api/office/office.context#document) property to access the **Document** object, and through it, can access the key members of the API for working with data in documents, such as the [Bindings](/javascript/api/office/office.bindings) and [CustomXmlParts](/javascript/api/office/office.customxmlparts) objects, and the [getSelectedDataAsync](/javascript/api/office/office.document#getselecteddataasync-coerciontype--options--callback-), [setSelectedDataAsync](/javascript/api/office/office.document#setselecteddataasync-data--options--callback-), and [getFileAsync](/javascript/api/office/office.document#getfileasync-filetype--options--callback-) methods. The `Document` object also provides the [mode](/javascript/api/office/office.document#mode) property for determining whether a document is read-only or in edit mode, the [url](/javascript/api/office/office.document#url) property to get the URL of the current document, and access to the [Settings](/javascript/api/office/office.settings) object. The `Document` object also supports adding event handlers for the [SelectionChanged](/javascript/api/office/office.documentselectionchangedeventargs) event, so you can detect when a user changes their selection in the document.

   A content or task pane add-in can access the `Document` object only after the DOM and runtime environment has been loaded, typically in the event handler for the [Office.initialize](/javascript/api/office) event. For information about the flow of events when an add-in is initialized, and how to check that the DOM and runtime and loaded successfully, see [Loading the DOM and runtime environment](loading-the-dom-and-runtime-environment.md).

3. **Objects for working with specific features.** To work with specific features of the API, use the following objects and methods.

    - The methods of the [Bindings](/javascript/api/office/office.bindings) object to create or get bindings, and the methods and properties of the [Binding](/javascript/api/office/office.binding) object to work with data.

    - The [CustomXmlParts](/javascript/api/office/office.customxmlparts), [CustomXmlPart](/javascript/api/office/office.customxmlpart) and associated objects to create and manipulate custom XML parts in Word documents.

    - The [File](/javascript/api/office/office.file) and [Slice](/javascript/api/office/office.slice) objects to create a copy of the entire document, break it into chunks or "slices", and then read or transmit the data in those slices.

    - The [Settings](/javascript/api/office/office.settings) object to save custom data, such as user preferences, and add-in state.


> [!IMPORTANT]
> Some of the API members aren't supported across all Office applications that can host content and task pane add-ins. To determine which members are supported, see any of the following:

For a summary of Office JavaScript API support across Office client applications, see [Understanding the Office JavaScript API](understanding-the-javascript-api-for-office.md).


## Reading and writing to an active selection

You can read or write to the user's current selection in a document, spreadsheet, or presentation. Depending on the Office application for your add-in, you can specify the type of data structure to read or write as a parameter in the [getSelectedDataAsync](/javascript/api/office/office.document#getselecteddataasync-coerciontype--options--callback-) and [setSelectedDataAsync](/javascript/api/office/office.document#setselecteddataasync-data--options--callback-) methods of the [Document](/javascript/api/office/office.document) object. For example, you can specify any type of data (text, HTML, tabular data, or Office Open XML) for Word, text and tabular data for Excel, and text for PowerPoint and Project. You can also create event handlers to detect changes to the user's selection. The following example gets data from the selection as text using the `getSelectedDataAsync` method.


```js
Office.context.document.getSelectedDataAsync(
    Office.CoercionType.Text, function (asyncResult) {
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

For more details and examples, see [Read and write data to the active selection in a document or spreadsheet](read-and-write-data-to-the-active-selection-in-a-document-or-spreadsheet.md).


## Binding to a region in a document or spreadsheet

You can use the `getSelectedDataAsync` and `setSelectedDataAsync` methods to read or write to the user's *current* selection in a document, spreadsheet, or presentation. However, if you would like to access the same region in a document across sessions of running your add-in without requiring the user to make a selection, you should first bind to that region. You can also subscribe to data and selection change events for that bound region.

You can add a binding by using [addFromNamedItemAsync](/javascript/api/office/office.bindings#addfromnameditemasync-itemname--bindingtype--options--callback-), [addFromPromptAsync](/javascript/api/office/office.bindings#addfrompromptasync-bindingtype--options--callback-), or [addFromSelectionAsync](/javascript/api/office/office.bindings#addfromselectionasync-bindingtype--options--callback-) methods of the [Bindings](/javascript/api/office/office.bindings) object. These methods return an identifier that you can use to access data in the binding, or to subscribe to its data change or selection change events.

The following is an example that adds a binding to the currently selected text in a document, by using the `Bindings.addFromSelectionAsync` method.



```js
Office.context.document.bindings.addFromSelectionAsync(
    Office.BindingType.Text, { id: 'myBinding' }, function (asyncResult) {
    if (asyncResult.status == Office.AsyncResultStatus.Failed) {
        write('Action failed. Error: ' + asyncResult.error.message);
    } else {
        write('Added new binding with type: ' +
            asyncResult.value.type + ' and id: ' + asyncResult.value.id);
    }
});

// Function that writes to a div with id='message' on the page.
function write(message){
    document.getElementById('message').innerText += message; 
}
```

For more details and examples, see [Bind to regions in a document or spreadsheet](bind-to-regions-in-a-document-or-spreadsheet.md).


## Getting entire documents

If your task pane add-in runs in PowerPoint or Word, you can use the [Document.getFileAsync](/javascript/api/office/office.document#getfileasync-filetype--options--callback-), [File.getSliceAsync](/javascript/api/office/office.file#getsliceasync-sliceindex--callback-), and [File.closeAsync](/javascript/api/office/office.file#closeasync-callback-) methods to get an entire presentation or document.

When you call `Document.getFileAsync` you get a copy of the document in a [File](/javascript/api/office/office.file) object. The `File` object provides access to the document in "chunks" represented as [Slice](/javascript/api/office/office.slice) objects. When you call `getFileAsync`, you can specify the file type (text or compressed Open Office XML format), and size of the slices (up to 4MB). To access the contents of the `File` object, you then call `File.getSliceAsync` which returns the raw data in the [Slice.data](/javascript/api/office/office.slice#data) property. If you specified compressed format, you will get the file data as a byte array. If you are transmitting the file to a web service, you can transform the compressed raw data to a base64-encoded string before submission. Finally, when you are finished getting slices of the file, use the `File.closeAsync` method to close the document.

For more details, see how to [get the whole document from an add-in for PowerPoint or Word](../word/get-the-whole-document-from-an-add-in-for-word.md).


## Reading and writing custom XML parts of a Word document

Using the Open Office XML file format and content controls, you can add custom XML parts to a Word document and bind elements in the XML parts to content controls in that document. When you open the document, Word reads and automatically populates bound content controls with data from the custom XML parts. Users can also write data into the content controls, and when the user saves the document, the data in the controls will be saved to the bound XML parts. Task pane add-ins for Word, can use the [Document.customXmlParts](/javascript/api/office/office.document#customxmlparts) property,[CustomXmlParts](/javascript/api/office/office.customxmlparts), [CustomXmlPart](/javascript/api/office/office.customxmlpart), and [CustomXmlNode](/javascript/api/office/office.customxmlnode) objects to read and write data dynamically to the document.

Custom XML parts may be associated with namespaces. To get data from custom XML parts in a namespace, use the [CustomXmlParts.getByNamespaceAsync](/javascript/api/office/office.customxmlparts#getbynamespaceasync-ns--options--callback-) method.

You can also use the [CustomXmlParts.getByIdAsync](/javascript/api/office/office.customxmlparts#getbyidasync-id--options--callback-) method to access custom XML parts by their GUIDs. After getting a custom XML part, use the [CustomXmlPart.getXmlAsync](/javascript/api/office/office.customxmlpart#getxmlasync-options--callback-) method to get the XML data.

To add a new custom XML part to a document, use the `Document.customXmlParts` property to get the custom XML parts that are in the document, and call the [CustomXmlParts.addAsync](/javascript/api/office/office.customxmlparts#addasync-xml--options--callback-) method.

For detailed information about how to work with custom XML parts with a task pane add-in, see [Creating Better Add-ins for Word with Office Open XML](../word/create-better-add-ins-for-word-with-office-open-xml.md).


## Persisting add-in settings


Often you need to save custom data for your add-in, such as a user's preferences or the add-in's state, and access that data the next time the add-in is opened. You can use common web programming techniques to save that data, such as browser cookies or HTML 5 web storage. Alternatively, if your add-in runs in Excel, PowerPoint, or Word, you can use the methods of the [Settings](/javascript/api/office/office.settings) object. Data created with the `Settings` object is stored in the spreadsheet, presentation, or document that the add-in was inserted into and saved with. This data is available to only the add-in that created it.

To avoid roundtrips to the server where the document is stored, data created with the `Settings` object is managed in memory at run time. Previously saved settings data is loaded into memory when the add-in is initialized, and changes to that data are only saved back to the document when you call the [Settings.saveAsync](/javascript/api/office/office.settings#saveasync-options--callback-) method. Internally, the data is stored in a serialized JSON object as name/value pairs. You use the [get](/javascript/api/office/office.settings#get-name-), [set](/javascript/api/office/office.settings#set-name--value-), and [remove](/javascript/api/office/office.settings#remove-name-) methods of the **Settings** object, to read, write, and delete items from the in-memory copy of the data. The following line of code shows how to create a setting named `themeColor` and set its value to 'green'.




```js
Office.context.document.settings.set('themeColor', 'green');
```

Because settings data created or deleted with the `set` and `remove` methods is acting on an in-memory copy of the data, you must call `saveAsync` to persist changes to settings data into the document your add-in is working with.

For more details about working with custom data using the methods of the `Settings` object, see [Persisting add-in state and settings](persisting-add-in-state-and-settings.md).


## Reading properties of a project document

If your task pane add-in runs in Project, your add-in can read data from some of the project fields, resource, and task fields in the active project. To do that, you use the methods and events of the [ProjectDocument](/javascript/api/office/office.document) object, which extends the `Document` object to provide additional Project-specific functionality.

For examples of reading Project data, see [Create your first task pane add-in for Project 2013 by using a text editor](../project/create-your-first-task-pane-add-in-for-project-by-using-a-text-editor.md).


## Permissions model and governance

Your add-in uses the `Permissions` element in its manifest to request permission to access the level of functionality it requires from the Office JavaScript API. For example, if your add-in requires read/write access to the document, its manifest must specify `ReadWriteDocument` as the text value in its `Permissions` element. Because permissions exist to protect a user's privacy and security, as a best practice you should request the minimum level of permissions it needs for its features. The following example shows how to request the **ReadDocument** permission in a task pane's manifest.


```XML
<?xml version="1.0" encoding="utf-8"?>
<OfficeApp xmlns="http://schemas.microsoft.com/office/appforoffice/1.0"
 xmlns:xsi="http://www.w3.org/2001/XMLSchema-instance" 
 xsi:type="TaskPaneApp">
???<!-- Other manifest elements omitted. -->
  <Permissions>ReadDocument</Permissions>
???
</OfficeApp>

```

For more information, see [Requesting permissions for API use in add-ins](requesting-permissions-for-api-use-in-content-and-task-pane-add-ins.md).


## See also

- [Office JavaScript API](../reference/javascript-api-for-office.md)
- [Schema reference for Office Add-ins manifests](../develop/add-in-manifests.md)
- [Troubleshoot user errors with Office Add-ins](../testing/testing-and-troubleshooting.md)
