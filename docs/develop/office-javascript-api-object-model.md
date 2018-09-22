---
title: Office JavaScript API object model
description: ''
ms.date: 07/27/2018
---


# Office JavaScript API object model
Office JavaScript add-ins give access to the host’s underlying functionality. Most of this access goes through a few important objects. The [Context](#context-object) object gives access to the runtime environment after initialization. The [Document](#document-object) object gives the user control over an Excel, PowerPoint, or Word document. The [Mailbox](#mailbox-object) object gives an Outlook add-in access to messages and user profiles. Understanding the relationships between these high-level objects is the foundation of a JavaScript add-in.

## Context object

**Applies to:** All add-in types

When an add-in is [initialized](https://docs.microsoft.com/office/dev/add-ins/develop/understanding-the-javascript-api-for-office#initializing-your-add-in), it has many different objects that it can interact with in the runtime environment. The add-in's runtime context is reflected in the API by the [Context](https://docs.microsoft.com/javascript/api/office/office.context?view=office-js) object. The **Context** is the main object that provides access to the most important objects of the API, such as the [Document](https://docs.microsoft.com/javascript/api/office/office.document?view=office-js) and [Mailbox](https://docs.microsoft.com/javascript/api/outlook/Office.mailbox?view=office-js) objects, which in turn provide access to document and mailbox content.

For example, in task pane or content add-ins, you can use the [document](https://docs.microsoft.com/javascript/api/office/office.context?view=office-js#document) property of the **Context** object to access the properties and methods of the **Document** object to interact with the content of Word documents, Excel worksheets, or Project schedules. Similarly, in Outlook add-ins, you can use the [mailbox](https://docs.microsoft.com/javascript/api/outlook/Office.mailbox?view=office-js) property of the **Context** object to access the properties and methods of the **Mailbox** object to interact with the message, meeting request, or appointment content.

The **Context** object also provides access to the [contentLanguage](https://docs.microsoft.com/javascript/api/office/office.context?view=office-js#contentlanguage) and [displayLanguage](https://docs.microsoft.com/javascript/api/office/office.context?view=office-js#displaylanguage) properties that let you determine the locale (language) used in the document or item, or by the host application. The [roamingSettings](https://docs.microsoft.com/javascript/api/office/office.context?view=office-js#roamingsettings) property lets you access the members of the [RoamingSettings](https://docs.microsoft.com/javascript/api/office/office.context?view=office-js#roamingsettings) object, which stores settings specific to your add-in for individual users' mailboxes. Finally, the **Context** object provides a [ui](https://docs.microsoft.com/javascript/api/office/office.ui?view=office-js) property that enables your add-in to launch pop-up dialogs.


## Document object

**Applies to:** Content and task pane add-in types

To interact with document data in Excel, PowerPoint, and Word, the API provides the [Document](https://docs.microsoft.com/javascript/api/office/office.document?view=office-js) object. You can use **Document** object members to access data from the following ways:

- Read and write to active selections in the form of text, contiguous cells (matrices), or tables.
    
- Tabular data (matrices or tables).
    
- Bindings (created with the "add" methods of the  **Bindings** object).
    
- Custom XML parts (only for Word).
    
- Settings or add-in state persisted per add-in on the document.
    
You can also use the  **Document** object to interact with data in Project documents. The Project-specific functionality of the API is documented in the members [ProjectDocument](https://docs.microsoft.com/javascript/api/office/office.document?view=office-js) abstract class. For more information about creating task pane add-ins for Project, see [Task pane add-ins for Project](../project/project-add-ins.md).

All these forms of data access start from an instance of the abstract  **Document** object.

You can access an instance of the  **Document** object when the task pane or content add-in is initialized by using the [document](https://docs.microsoft.com/javascript/api/office/office.context?view=office-js#document) property of the **Context** object. The **Document** object defines common data access functions shared across Word and Excel documents, and also provides access to the **CustomXmlParts** object for Word documents.

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
|Table|Provides the data in the selection or binding as a [TableData](https://docs.microsoft.com/javascript/api/office/office.tabledata?view=office-js) object. The **TableData** object exposes the data through the **headers** and **rows** properties.|Table data access is supported only in Excel 2013 and Word 2013.|

#### Data type coercion

The data access methods on the **Document** and [Binding](https://docs.microsoft.com/javascript/api/office/office.binding?view=office-js) objects support specifying the desired data type using the _coercionType_ parameter of these methods, and corresponding [CoercionType](https://docs.microsoft.com/javascript/api/office/office.coerciontype?view=office-js) enumeration values. Regardless of the actual shape of the binding, the different Office applications support the common data types by trying to coerce the data into the requested data type. For example, if a Word table or paragraph is selected, the developer can specify to read it as plain text, HTML, Office Open XML, or a table, and the API implementation handles the necessary transformations and data conversions.


> [!TIP]
> **When should you use the matrix versus table coercionType for data access?** If you need your tabular data to grow dynamically when rows and columns are added, and you must work with table headers, you should use the table data type (by specifying the _coercionType_ parameter of a **Document** or **Binding** object data access method as `"table"` or **Office.CoercionType.Table**). Adding rows and columns within the data structure is supported in both table and matrix data, but appending rows and columns is supported only for table data. If you aren't planning on adding rows and columns, and your data doesn't require header functionality, then you should use the matrix data type (by specifying the  _coercionType_ parameter of the data access method as `"matrix"` or **Office.CoercionType.Matrix**), which provides a simpler model of interacting with the data.

If the data can't be coerced to the specified type, the [AsyncResult.status](https://docs.microsoft.com/javascript/api/office/office.asyncresult?view=office-js.error) property in the callback returns `"failed"`, and you can use the [AsyncResult.error](https://docs.microsoft.com/javascript/api/office/office.asyncresult?view=office-js.context) property to access an [Error](https://docs.microsoft.com/javascript/api/office/office.error?view=office-js) object with information about why the method call failed.


## Working with selections using the Document object


The  **Document** object exposes methods that let you to read and write to the user's current selection in a "get and set" fashion. To do that, the **Document** object provides the **getSelectedDataAsync** and **setSelectedDataAsync** methods.

For code examples that demonstrate how to perform tasks with selections, see [Read and write data to the active selection in a document or spreadsheet](read-and-write-data-to-the-active-selection-in-a-document-or-spreadsheet.md).


## Working with bindings using the Bindings and Binding objects


Binding-based data access enables content and task pane add-ins to consistently access a particular region of a document or spreadsheet through an identifier associated with a binding. The add-in first needs to establish the binding by calling one of the methods that associates a portion of the document with a unique identifier: [addFromPromptAsync](https://docs.microsoft.com/javascript/api/office/office.bindings?view=office-js#addfrompromptasync-bindingtype--options--callback-), [addFromSelectionAsync](https://docs.microsoft.com/javascript/api/office/office.bindings?view=office-js#addfromselectionasync-bindingtype--options--callback-), or [addFromNamedItemAsync](https://docs.microsoft.com/javascript/api/office/office.bindings?view=office-js#addfromnameditemasync-itemname--bindingtype--options--callback-). After the binding is established, the add-in can use the provided identifier to access the data contained in the associated region of the document or spreadsheet. Creating bindings provides the following value to your add-in:


- Permits access to common data structures across supported Office applications, such as: tables, ranges, or text (a contiguous run of characters).
    
- Enables read/write operations without requiring the user to make a selection.
    
- Establishes a relationship between the add-in and the data in the document. Bindings are persisted in the document, and can be accessed at a later time.
    
Establishing a binding also allows you to subscribe to data and selection change events that are scoped to that particular region of the document or spreadsheet. This means that the add-in is only notified of changes that happen within the bound region as opposed to general changes across the whole document or spreadsheet.

The [Bindings](https://docs.microsoft.com/javascript/api/office/office.bindings?view=office-js) object exposes a [getAllAsync](https://docs.microsoft.com/javascript/api/office/office.bindings?view=office-js#getallasync-options--callback-) method that gives access to the set of all bindings established on the document or spreadsheet. An individual binding can be accessed by its ID using either the [Bindings.getBindingByIdAsync](https://docs.microsoft.com/javascript/api/office/office.bindings?view=office-js#getbyidasync-id--options--callback-) or [Office.select](https://docs.microsoft.com/javascript/api/office?view=office-js) methods. You can establish new bindings as well as remove existing ones by using one of the following methods of the **Bindings** object: [addFromSelectionAsync](https://docs.microsoft.com/javascript/api/office/office.bindings?view=office-js#addfromselectionasync-bindingtype--options--callback-), [addFromPromptAsync](https://docs.microsoft.com/javascript/api/office/office.bindings?view=office-js#addfrompromptasync-bindingtype--options--callback-), [addFromNamedItemAsync](https://docs.microsoft.com/javascript/api/office/office.bindings?view=office-js#addfromnameditemasync-itemname--bindingtype--options--callback-), or [releaseByIdAsync](https://docs.microsoft.com/javascript/api/office/office.bindings?view=office-js#releasebyidasync-id--options--callback-).

There are three different types of bindings that you specify with the  _bindingType_ parameter when you create a binding with the **addFromSelectionAsync**, **addFromPromptAsync** or **addFromNamedItemAsync** methods:



|**Binding type**|**Description**|**Host application support**|
|:-----|:-----|:-----|
|Text binding|Binds to a region of the document that can be represented as text.|In Word, most contiguous selections are valid, while in Excel only single cell selections can be the target of a text binding. In Excel, only plain text is supported. In Word, three formats are supported: plain text, HTML, and Open XML for Office.|
|Matrix binding|Binds to a fixed region of a document that contains tabular data without headers. Data in a matrix binding is written or read as a two dimensional **Array**, which in JavaScript is implemented as an array of arrays. For example, two rows of **string** values in two columns can be written or read as ` [['a', 'b'], ['c', 'd']]`, and a single column of three rows can be written or read as `[['a'], ['b'], ['c']]`.|In Excel, any contiguous selection of cells can be used to establish a matrix binding. In Word, only tables support matrix binding.|
|Table binding|Binds to a region of a document that contains a table with headers. Data in a table binding is written or read as a [TableData](https://docs.microsoft.com/javascript/api/office/office.tabledata?view=office-js) object. The **TableData** object exposes the data through the **headers** and **rows** properties.|Any Excel or Word table can be the basis for a table binding. After you establish a table binding, each new row or column a user adds to the table is automatically included in the binding. |

<br/>

After a binding is created by using one of the three "add" methods of the  **Bindings** object, you can work with the binding's data and properties by using the methods of the corresponding object: [MatrixBinding](https://docs.microsoft.com/javascript/api/office/office.matrixbinding?view=office-js), [TableBinding](https://docs.microsoft.com/javascript/api/office/office.tablebinding?view=office-js), or [TextBinding](https://docs.microsoft.com/javascript/api/office/office.textbinding?view=office-js). All three of these objects inherit the [getDataAsync](https://docs.microsoft.com/javascript/api/office/office.binding?view=office-js#getdataasync-options--callback-) and [setDataAsync](https://docs.microsoft.com/javascript/api/office/office.binding?view=office-js#setdataasync-data--options--callback-) methods of the **Binding** object that enable to you interact with the bound data.

For code examples that demonstrate how to perform tasks with bindings, see [Bind to regions in a document or spreadsheet](bind-to-regions-in-a-document-or-spreadsheet.md).


## Working with custom XML parts using the CustomXmlParts and CustomXmlPart objects


 **Applies to:** Task pane add-ins for Word

The [CustomXmlParts](https://docs.microsoft.com/javascript/api/office/office.customxmlparts?view=office-js) and [CustomXmlPart](https://docs.microsoft.com/javascript/api/office/office.customxmlpart?view=office-js) objects of the API provide access to custom XML parts in Word documents, which enable XML-driven manipulation of the contents of the document. For demonstrations of working with the **CustomXmlParts** and **CustomXmlPart** objects, see the [Word-add-in-Work-with-custom-XML-parts](https://github.com/OfficeDev/Word-Add-in-Work-with-custom-XML-parts) code sample.


## Working with the entire document using the getFileAsync method


 **Applies to:** Task pane add-ins for Word and PowerPoint

The [Document.getFileAsync](https://docs.microsoft.com/javascript/api/office/office.document?view=office-js#getfileasync-filetype--options--callback-) method and members of the [File](https://docs.microsoft.com/javascript/api/office/office.file?view=office-js) and [Slice](https://docs.microsoft.com/javascript/api/office/office.slice?view=office-js) objects to provide functionality for getting entire Word and PowerPoint document files in slices (chunks) of up to 4 MB at a time. For more information, see [Get the whole document from an add-in for PowerPoint or Word](../word/get-the-whole-document-from-an-add-in-for-word.md).


## Mailbox object

**Applies to:** Outlook add-ins

Outlook add-ins primarily use a subset of the API exposed through the [Mailbox](https://docs.microsoft.com/javascript/api/outlook/Office.mailbox?view=office-js) object. To access the objects and members specifically for use in Outlook add-ins, such as the [Item](https://docs.microsoft.com/javascript/api/outlook/Office.mailbox?view=office-js) object, you use the [mailbox](https://docs.microsoft.com/javascript/api/outlook/Office.mailbox?view=office-js) property of the **Context** object to access the **Mailbox** object, as shown in the following line of code.




```js
// Access the Item object.
var item = Office.context.mailbox.item;

```

Additionally, Outlook add-ins can use the following objects:


-  **Office** object: for initialization.
    
-  **Context** object: for access to content and display language properties.
    
-  **RoamingSettings** object: for saving Outlook add-in-specific custom settings to the user's mailbox where the add-in is installed.
    
For information about using JavaScript in Outlook add-ins, see [Outlook add-ins](https://docs.microsoft.com/outlook/add-ins/).