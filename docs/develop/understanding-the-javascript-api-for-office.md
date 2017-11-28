---
title: Understanding the JavaScript API for Office
description: 
ms.date: 11/20/2017 
---


# Understanding the JavaScript API for Office

This article provides information about the JavaScript API for Office and how to use it. For reference information, see [JavaScript API for Office](https://dev.office.com/reference/add-ins/javascript-api-for-office). For information about updating Visual Studio project files to the most current version of the JavaScript API for Office, see [Update the version of your JavaScript API for Office and manifest schema files](update-your-javascript-api-for-office-and-manifest-schema-version.md).

> [!NOTE]
> When you build your add-in, if you plan to [publish](../publish/publish.md) your add-in to the Office Store, make sure that you conform to the [Office Store validation policies](https://dev.office.com/officestore/docs/validation-policies). For example, to pass validation, your add-in must work across all platforms that support the methods that you define (for more information, see [section 4.12](https://dev.office.com/officestore/docs/validation-policies#4-apps-and-add-ins-behave-predictably) and the [Office Add-in host and availability page](https://dev.office.com/add-in-availability)).

## Referencing the JavaScript API for Office library in your add-in

The [JavaScript API for Office](https://dev.office.com/reference/add-ins/javascript-api-for-office) library consists of the Office.js file and associated host application-specific .js files, such as Excel-15.js and Outlook-15.js. The simplest method of referencing the API is using our CDN by adding the following `<script>` to your page's `<head>` tag:  

```html
<script src="https://appsforoffice.microsoft.com/lib/1/hosted/Office.js" type="text/javascript"></script>
```

This will download and cache the JavaScript API for Office files the first time your add-in loads to make sure that it is using the most up-to-date implementation of Office.js and its associated files for the specified version.

For more details around the Office.js CDN, including how versioning and backward compatability is handled, see [Referencing the JavaScript API for Office library from its content delivery network (CDN)](referencing-the-javascript-api-for-office-library-from-its-cdn.md).

## Initializing your add-in

**Applies to:** All add-in types

Office.js provides an initialization event which gets fired when the API is fully loaded and ready to begin interacting with the user. You can use the **initialize** event handler to implement common add-in initialization scenarios, such as prompting the user to select some cells in Excel, and then inserting a chart initialized with those selected values. You can also use the initialize event handler to initialize other custom logic for your add-in, such as establishing bindings, prompting for default add-in settings values, and so on.

At a minimum, the initialize event would look like the follow example:     

```js
Office.initialize = function () { };
```
If you are using additional JavaScript frameworks that include their own initialization handler or tests, these should be placed within the Office.initialize event. For example, [JQuery's](https://jquery.com) `$(document).ready()` function would be referenced as follows:

```js
Office.initialize = function () {
    // Office is ready
    $(document).ready(function () {        
        // The document is ready
    });
  };
```

All pages within an Office Add-ins are required to assign an event handler to the initialize event, **Office.initialize**.
If you fail to assign an event handler, your add-in may raise an error when it starts. Also, if a user attempts to use your add-in with an Office Online web client, such as Excel Online, PowerPoint Online, or Outlook Web App, it will fail to run. If you don't need any initialization code, then the body of the function you assign to **Office.initialize** can be empty, as it is in the first example above.

For more detail about the sequence of events when an add-in is initialized, see [Loading the DOM and runtime environment](loading-the-dom-and-runtime-environment.md).

#### Initialization reason
For task pane and content add-ins, Office.initialize provides an additional _reason_ parameter. This parameter can be used to determine how an add-in was added to the current document. You can use this to provide different logic for when an add-in is first inserted versus when it already existed within the document. 

```js
Office.initialize = function (reason) {
    $(document).ready(function () {
      switch (reason) {
        case 'inserted': console.log('The add-in was just inserted.');
        case 'documentOpened': console.log('The add-in is already part of the document.');
    }
}
```
For more information, see [Office.initialize Event](https://dev.office.com/reference/add-ins/shared/office.initialize) and [InitializationReason Enumeration](https://dev.office.com/reference/add-ins/shared/initializationreason-enumeration). 

## Context object

**Applies to:** All add-in types

When an add-in is initialized, it has many different objects that it can interact with in the runtime environment. The add-in's runtime context is reflected in the API by the [Context](https://dev.office.com/reference/add-ins/shared/office.context) object. The **Context** is the main object that provides access to the most important objects of the API, such as the [Document](https://dev.office.com/reference/add-ins/shared/document) and [Mailbox](https://dev.office.com/reference/add-ins/outlook/Office.context.mailbox) objects, which in turn provide access to document and mailbox content.

For example, in task pane or content add-ins, you can use the [document](https://dev.office.com/reference/add-ins/shared/office.context.document) property of the **Context** object to access the properties and methods of the **Document** object to interact with the content of Word documents, Excel worksheets, or Project schedules. Similarly, in Outlook add-ins, you can use the [mailbox](https://dev.office.com/reference/add-ins/outlook/Office.context.mailbox) property of the **Context** object to access the properties and methods of the **Mailbox** object to interact with the message, meeting request, or appointment content.

The **Context** object also provides access to the [contentLanguage](https://dev.office.com/reference/add-ins/shared/office.context.contentlanguage) and [displayLanguage](https://dev.office.com/reference/add-ins/shared/office.context.displaylanguage) properties that let you determine the locale (language) used in the document or item, or by the host application. And, the [roamingSettings](https://dev.office.com/reference/add-ins/outlook/Office.context) property that lets you access the members of the [RoamingSettings](https://dev.office.com/reference/add-ins/outlook/RoamingSettings) object. Finally, the **Context** object provides a [ui](https://dev.office.com/reference/add-ins/shared/officeui) property that enables your add-in to launch pop-up dialogs.


## Document object

**Applies to:** Content and task pane add-in types

To interact with document data in Excel, PowerPoint, and Word, the API provides the [Document](https://dev.office.com/reference/add-ins/shared/document) object. You can use **Document** object members to access data from the following ways:

- Read and write to active selections in the form of text, contiguous cells (matrices), or tables.
    
- Tabular data (matrices or tables).
    
- Bindings (created with the "add" methods of the  **Bindings** object).
    
- Custom XML parts (only for Word).
    
- Settings or add-in state persisted per add-in on the document.
    
You can also use the  **Document** object to interact with data in Project documents. The Project-specific functionality of the API is documented in the members [ProjectDocument](https://dev.office.com/reference/add-ins/shared/projectdocument.projectdocument) abstract class. For more information about creating task pane add-ins for Project, see [Task pane add-ins for Project](../project/project-add-ins.md).

All these forms of data access start from an instance of the abstract  **Document** object.

You can access an instance of the  **Document** object when the task pane or content add-in is initialized by using the [document](https://dev.office.com/reference/add-ins/shared/office.context.document) property of the **Context** object. The **Document** object defines common data access functions shared across Word and Excel documents, and also provides access to the **CustomXmlParts** object for Word documents.

The  **Document** object supports four ways for developers to access document contents:


- Selection-based access
    
- Binding-based access
    
- Custom XML part-based access (Word only)
    
- Entire document-based access (PowerPoint and Word only)
    
To help you understand how selection- and binding-based data access methods work, we will first explain how the data-access APIs provide consistent data access across different Office applications.


### Consistent data access across Office applications

 **Applies to:** Content and task pane add-in types

To create extensions that seamlessly work across different Office documents, the JavaScript API for Office abstracts away the particularities of each Office application through common data types and the ability to coerce different document contents into three common data types.


#### Common data types

In both selection-based and binding-based data access, document contents are exposed through data types that are common across all the supported Office applications. In Office 2013, three main data types are supported:



|**Data type**|**Description**|**Host application support**|
|:-----|:-----|:-----|
|Text|Provides a string representation of the data in the selection or binding.|In Excel 2013, Project 2013, and PowerPoint 2013, only plain text is supported. In Word 2013, three text formats are supported: plain text, HTML, and Office Open XML (OOXML). When text is selected in a cell in Excel, selection-based methods read and write to the entire contents of the cell, even if only a portion of the text is selected in the cell. When text is selected in Word and PowerPoint, selection-based methods read and write only to the run of characters that are selected. Project 2013 and PowerPoint 2013 support only selection-based data access.|
|Matrix|Provides the data in the selection or binding as a two dimensional **Array**, which in JavaScript is implemented as an array of arrays. For example, two rows of **string** values in two columns would be ` [['a', 'b'], ['c', 'd']]`, and a single column of three rows would be `[['a'], ['b'], ['c']]`.|Matrix data access is supported only in Excel 2013 and Word 2013.|
|Table|Provides the data in the selection or binding as a [TableData](https://dev.office.com/reference/add-ins/shared/tabledata) object. The **TableData** object exposes the data through the **headers** and **rows** properties.|Table data access is supported only in Excel 2013 and Word 2013.|

#### Data type coercion

The data access methods on the **Document** and [Binding](https://dev.office.com/reference/add-ins/shared/binding) objects support specifying the desired data type using the _coercionType_ parameter of these methods, and corresponding [CoercionType](https://dev.office.com/reference/add-ins/shared/coerciontype-enumeration) enumeration values. Regardless of the actual shape of the binding, the different Office applications support the common data types by trying to coerce the data into the requested data type. For example, if a Word table or paragraph is selected, the developer can specify to read it as plain text, HTML, Office Open XML, or a table, and the API implementation handles the necessary transformations and data conversions.


> [!TIP]
> **When should you use the matrix versus table coercionType for data access?** If you need your tabular data to grow dynamically when rows and columns are added, and you must work with table headers, you should use the table data type (by specifying the _coercionType_ parameter of a **Document** or **Binding** object data access method as `"table"` or **Office.CoercionType.Table**). Adding rows and columns within the data structure is supported in both table and matrix data, but appending rows and columns is supported only for table data. If you aren't planning on adding rows and columns, and your data doesn't require header functionality, then you should use the matrix data type (by specifying the  _coercionType_ parameter of the data access method as `"matrix"` or **Office.CoercionType.Matrix**), which provides a simpler model of interacting with the data.

If the data can't be coerced to the specified type, the [AsyncResult.status](https://dev.office.com/reference/add-ins/shared/asyncresult.error) property in the callback returns `"failed"`, and you can use the [AsyncResult.error](https://dev.office.com/reference/add-ins/shared/asyncresult.context) property to access an [Error](https://dev.office.com/reference/add-ins/shared/error) object with information about why the method call failed.


## Working with selections using the Document object


The  **Document** object exposes methods that let you to read and write to the user's current selection in a "get and set" fashion. To do that, the **Document** object provides the **getSelectedDataAsync** and **setSelectedDataAsync** methods.

For code examples that demonstrate how to perform tasks with selections, see [Read and write data to the active selection in a document or spreadsheet](read-and-write-data-to-the-active-selection-in-a-document-or-spreadsheet.md).


## Working with bindings using the Bindings and Binding objects


Binding-based data access enables content and task pane add-ins to consistently access a particular region of a document or spreadsheet through an identifier associated with a binding. The add-in first needs to establish the binding by calling one of the methods that associates a portion of the document with a unique identifier: [addFromPromptAsync](https://dev.office.com/reference/add-ins/shared/bindings.addfrompromptasync), [addFromSelectionAsync](https://dev.office.com/reference/add-ins/shared/bindings.addfromselectionasync), or [addFromNamedItemAsync](https://dev.office.com/reference/add-ins/shared/bindings.addfromnameditemasync). After the binding is established, the add-in can use the provided identifier to access the data contained in the associated region of the document or spreadsheet. Creating bindings provides the following value to your add-in:


- Permits access to common data structures across supported Office applications, such as: tables, ranges, or text (a contiguous run of characters).
    
- Enables read/write operations without requiring the user to make a selection.
    
- Establishes a relationship between the add-in and the data in the document. Bindings are persisted in the document, and can be accessed at a later time.
    
Establishing a binding also allows you to subscribe to data and selection change events that are scoped to that particular region of the document or spreadsheet. This means that the add-in is only notified of changes that happen within the bound region as opposed to general changes across the whole document or spreadsheet.

The [Bindings](https://dev.office.com/reference/add-ins/shared/bindings.bindings) object exposes a [getAllAsync](https://dev.office.com/reference/add-ins/shared/bindings.getallasync) method that gives access to the set of all bindings established on the document or spreadsheet. An individual binding can be accessed by its ID using either the [Bindings.getBindingByIdAsync](https://dev.office.com/reference/add-ins/shared/bindings.getbyidasync) or [Office.select](https://dev.office.com/reference/add-ins/shared/office.select) methods. You can establish new bindings as well as remove existing ones by using one of the following methods of the **Bindings** object: [addFromSelectionAsync](https://dev.office.com/reference/add-ins/shared/bindings.addfromselectionasync), [addFromPromptAsync](https://dev.office.com/reference/add-ins/shared/bindings.addfrompromptasync), [addFromNamedItemAsync](https://dev.office.com/reference/add-ins/shared/bindings.addfromnameditemasync), or [releaseByIdAsync](https://dev.office.com/reference/add-ins/shared/bindings.releasebyidasync).

There are three different types of bindings that you specify with the  _bindingType_ parameter when you create a binding with the **addFromSelectionAsync**, **addFromPromptAsync** or **addFromNamedItemAsync** methods:



|**Binding type**|**Description**|**Host application support**|
|:-----|:-----|:-----|
|Text binding|Binds to a region of the document that can be represented as text.|In Word, most contiguous selections are valid, while in Excel only single cell selections can be the target of a text binding. In Excel, only plain text is supported. In Word, three formats are supported: plain text, HTML, and Open XML for Office.|
|Matrix binding|Binds to a fixed region of a document that contains tabular data without headers. Data in a matrix binding is written or read as a two dimensional **Array**, which in JavaScript is implemented as an array of arrays. For example, two rows of **string** values in two columns can be written or read as ` [['a', 'b'], ['c', 'd']]`, and a single column of three rows can be written or read as `[['a'], ['b'], ['c']]`.|In Excel, any contiguous selection of cells can be used to establish a matrix binding. In Word, only tables support matrix binding.|
|Table binding|Binds to a region of a document that contains a table with headers. Data in a table binding is written or read as a [TableData](https://dev.office.com/reference/add-ins/shared/tabledata) object. The **TableData** object exposes the data through the **headers** and **rows** properties.|Any Excel or Word table can be the basis for a table binding. After you establish a table binding, each new row or column a user adds to the table is automatically included in the binding. |

<br/>

After a binding is created by using one of the three "add" methods of the  **Bindings** object, you can work with the binding's data and properties by using the methods of the corresponding object: [MatrixBinding](https://dev.office.com/reference/add-ins/shared/binding.matrixbinding), [TableBinding](https://dev.office.com/reference/add-ins/shared/binding.tablebinding), or [TextBinding](https://dev.office.com/reference/add-ins/shared/binding.textbinding). All three of these objects inherit the [getDataAsync](https://dev.office.com/reference/add-ins/shared/binding.getdataasync) and [setDataAsync](https://dev.office.com/reference/add-ins/shared/binding.setdataasync) methods of the **Binding** object that enable to you interact with the bound data.

For code examples that demonstrate how to perform tasks with bindings, see [Bind to regions in a document or spreadsheet](bind-to-regions-in-a-document-or-spreadsheet.md).


## Working with custom XML parts using the CustomXmlParts and CustomXmlPart objects


 **Applies to:** Task pane add-ins for Word

The [CustomXmlParts](https://dev.office.com/reference/add-ins/shared/customxmlparts.customxmlparts) and [CustomXmlPart](https://dev.office.com/reference/add-ins/shared/customxmlpart.customxmlpart) objects of the API provide access to custom XML parts in Word documents, which enable XML-driven manipulation of the contents of the document. For demonstrations of working with the **CustomXmlParts** and **CustomXmlPart** objects, see the [Word-Add-in-Work-with-custom-XML-parts](https://github.com/OfficeDev/Word-Add-in-Work-with-custom-XML-parts) code sample.


## Working with the entire document using the getFileAsync method


 **Applies to:** Task pane add-ins for Word and PowerPoint

The [Document.getFileAsync](https://dev.office.com/reference/add-ins/shared/document.getfileasync) method and members of the [File](https://dev.office.com/reference/add-ins/shared/file) and [Slice](https://dev.office.com/reference/add-ins/shared/slice) objects to provide functionality for getting entire Word and PowerPoint document files in slices (chunks) of up to 4 MB at a time. For more information, see [Get the whole document from an add-in for PowerPoint or Word](../word/get-the-whole-document-from-an-add-in-for-word.md).


## Mailbox object


 **Applies to:** Outlook add-ins

Outlook add-ins primarily use a subset of the API exposed through the [Mailbox](https://dev.office.com/reference/add-ins/outlook/Office.context.mailbox) object. To access the objects and members specifically for use in Outlook add-ins, such as the [Item](https://dev.office.com/reference/add-ins/outlook/Office.context.mailbox.item) object, you use the [mailbox](https://dev.office.com/reference/add-ins/outlook/Office.context.mailbox) property of the **Context** object to access the **Mailbox** object, as shown in the following line of code.




```js
// Access the Item object.
var item = Office.context.mailbox.item;

```

Additionally, Outlook add-ins can use the following objects:


-  **Office** object: for initialization.
    
-  **Context** object: for access to content and display language properties.
    
-  **RoamingSettings** object: for saving Outlook add-in-specific custom settings to the user's mailbox where the add-in is installed.
    
For information about using JavaScript in Outlook add-ins, see [Outlook add-ins](https://docs.microsoft.com/en-us/outlook/add-ins/).


## API support matrix


This table summarizes the API and features supported across add-in types (content, task pane, and Outlook) and the Office applications that can host them when you specify the Office host applications your add-in supports by using the [1.1 add-in manifest schema and features supported by v1.1 JavaScript API for Office](update-your-javascript-api-for-office-and-manifest-schema-version.md).


|||||||||
|:-----|:-----|:-----:|:-----:|:-----:|:-----:|:-----:|:-----:|
||**Host name**|Database|Workbook|Mailbox|Presentation|Document|Project|
||**Supported** **Host applications**|Access web apps|Excel,<br/>Excel Online|Outlook,<br/>Outlook Web App,<br/>OWA for Devices|PowerPoint,<br/>PowerPoint Online|Word|Project|
|**Supported add-in types**|Content|Y|Y||Y|||
||Task pane||Y||Y|Y|Y|
||Outlook|||Y||||
|**Supported API features**|Read/Write Text||Y||Y|Y|Y<br/>(Read only)|
||Read/Write Matrix||Y|||Y||
||Read/Write Table||Y|||Y||
||Read/Write HTML|||||Y||
||Read/Write<br/>Office Open XML|||||Y||
||Read task, resource, view, and field properties||||||Y|
||Selection changed events||Y|||Y||
||Get whole document||||Y|Y||
||Bindings and binding events|Y<br/>(Only full and partial table bindings)|Y|||Y||
||Read/Write Custom XML Parts|||||Y||
||Persist add-in state data (settings)|Y<br/>(Per host add-in)|Y<br/>(Per document)|Y<br/>(Per mailbox)|Y<br/>(Per document)|Y<br/>(Per document)||
||Settings changed events|Y|Y||Y|Y||
||Get active view mode<br/>and view changed events||||Y|||
||Navigate to locations<br/>in the document||Y||Y|Y||
||Activate contextually<br/>using rules and RegEx|||Y||||
||Read Item properties|||Y||||
||Read User profile|||Y||||
||Get attachments|||Y||||
||Get User identity token|||Y||||
||Call Exchange Web Services|||Y||||
