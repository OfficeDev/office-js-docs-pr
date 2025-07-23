---
title: Common JavaScript API object model
description: Learn about the Office JavaScript common API object model.
ms.topic: overview
ms.date: 03/21/2023
ms.localizationpriority: medium
---

# Common JavaScript API object model

[!include[information about the common API](../includes/alert-common-api-info.md)]

Office JavaScript APIs give access to the Office client application's underlying functionality. Most of this access goes through a few important objects. The [Context](#context-object) object gives access to the runtime environment after initialization. The [Document](#document-object) object gives the user control over an Excel, PowerPoint, or Word document. The [Mailbox](#mailbox-object) object gives an Outlook add-in access to messages, appointments, and user profiles. Understanding the relationships between these high-level objects is the foundation of an Office Add-in.

## Context object

**Applies to:** All add-in types

When an add-in is [initialized](initialize-add-in.md), it has many different objects that it can interact with in the runtime environment. The add-in's runtime context is reflected in the API by the [Context](/javascript/api/office/office.context) object. The **Context** is the main object that provides access to the most important objects of the API, such as the [Document](/javascript/api/office/office.document) and [Mailbox](/javascript/api/outlook/office.mailbox) objects, which in turn provide access to document and mailbox content.

For example, in task pane or content add-ins, you can use the [document](/javascript/api/office/office.context#office-office-context-document-member) property of the **Context** object to access the properties and methods of the **Document** object to interact with the content of Word documents, Excel worksheets, or Project schedules. Similarly, in Outlook add-ins, you can use the [mailbox](/javascript/api/office/office.context#office-office-context-mailbox-member) property of the **Context** object to access the properties and methods of the **Mailbox** object to interact with the message, meeting request, or appointment content.

The **Context** object also provides access to the [contentLanguage](/javascript/api/office/office.context#office-office-context-contentlanguage-member) and [displayLanguage](/javascript/api/office/office.context#office-office-context-displaylanguage-member) properties that let you determine the locale (language) used in the document or item, or by the Office application. The [roamingSettings](/javascript/api/office/office.context#office-office-context-roamingsettings-member) property lets you access the members of the [RoamingSettings](/javascript/api/office/office.context#office-office-context-roamingsettings-member) object, which stores settings specific to your add-in for individual users' mailboxes. Finally, the **Context** object provides a [ui](/javascript/api/office/office.context#office-office-context-ui-member) property that enables your add-in to launch pop-up dialogs.

## Document object

**Applies to:** Content and task pane add-in types

To interact with document data in Excel, PowerPoint, and Word, the API provides the [Document](/javascript/api/office/office.document) object. You can use `Document` object members to access data from the following ways.

- Read and write to active selections in the form of text, contiguous cells (matrices), or tables.

- Tabular data (matrices or tables).

- Bindings (created with the "add" methods of the `Bindings` object).

- Custom XML parts (only for Word).

- Settings or add-in state persisted per add-in on the document.

You can also use the `Document` object to interact with data in Project documents. The Project-specific functionality of the API is documented in the members [ProjectDocument](/javascript/api/office/office.document) abstract class. For more information about creating task pane add-ins for Project, see [Task pane add-ins for Project](../project/project-add-ins.md).

All these forms of data access start from an instance of the abstract `Document` object.

You can access an instance of the `Document` object when the task pane or content add-in is initialized by using the [document](/javascript/api/office/office.context#office-office-context-document-member) property of the `Context` object. The `Document` object defines common data access methods shared across Word and Excel documents, and also provides access to the `CustomXmlParts` object for Word documents.

The `Document` object supports four ways for developers to access document contents.

- Selection-based access

- Binding-based access

- Custom XML part-based access (Word only)

- Entire document-based access (PowerPoint and Word only)

To help you understand how selection- and binding-based data access methods work, we will first explain how the data-access APIs provide consistent data access across different Office applications.

### Consistent data access across Office applications

 **Applies to:** Content and task pane add-in types

To create extensions that seamlessly work across different Office documents, the Office JavaScript API abstracts away the particularities of each Office application through common data types and the ability to coerce different document contents into three common data types.

#### Common data types

In both selection-based and binding-based data access, document contents are exposed through data types that are common across all the supported Office applications. Three main data types are supported.

|Data type|Description|Host application support|
|:-----|:-----|:-----|
|Text|Provides a string representation of the data in the selection or binding.|In Excel, Project, and PowerPoint, only plain text is supported. In Word, three text formats are supported: plain text, HTML, and Office Open XML (OOXML). When text is selected in a cell in Excel, selection-based methods read and write to the entire contents of the cell, even if only a portion of the text is selected in the cell. When text is selected in Word and PowerPoint, selection-based methods read and write only to the run of characters that are selected. Project and PowerPoint support only selection-based data access.|
|Matrix|Provides the data in the selection or binding as a two dimensional **Array**, which in JavaScript is implemented as an array of arrays. For example, two rows of **string** values in two columns would be ` [['a', 'b'], ['c', 'd']]`, and a single column of three rows would be `[['a'], ['b'], ['c']]`.|Matrix data access is supported only in Excel and Word.|
|Table|Provides the data in the selection or binding as a [TableData](/javascript/api/office/office.tabledata) object. The `TableData` object exposes the data through the `headers` and `rows` properties.|Table data access is supported only in Excel and Word.|

#### Data type coercion

The data access methods on the `Document` and [Binding](/javascript/api/office/office.binding) objects support specifying the desired data type using the _coercionType_ parameter of these methods, and corresponding [CoercionType](/javascript/api/office/office.coerciontype) enumeration values. Regardless of the actual shape of the binding, the different Office applications support the common data types by trying to coerce the data into the requested data type. For example, if a Word table or paragraph is selected, the developer can specify to read it as plain text, HTML, Office Open XML, or a table, and the API implementation handles the necessary transformations and data conversions.

> [!TIP]
> **When should you use the matrix versus table coercionType for data access?** If you need your tabular data to grow dynamically when rows and columns are added, and you must work with table headers, you should use the table data type (by specifying the _coercionType_ parameter of a `Document` or `Binding` object data access method as `"table"` or `Office.CoercionType.Table`). Adding rows and columns within the data structure is supported in both table and matrix data, but appending rows and columns is supported only for table data. If you aren't planning on adding rows and columns, and your data doesn't require header functionality, then you should use the matrix data type (by specifying the  _coercionType_ parameter of the data access method as `"matrix"` or `Office.CoercionType.Matrix`), which provides a simpler model of interacting with the data.

If the data can't be coerced to the specified type, the [AsyncResult.status](/javascript/api/office/office.asyncresult#office-office-asyncresult-status-member) property in the callback returns `"failed"`, and you can use the [AsyncResult.error](/javascript/api/office/office.asyncresult#office-office-asyncresult-error-member) property to access an [Error](/javascript/api/office/office.error) object with information about why the method call failed.

## Work with selections using the Document object

The `Document` object exposes methods that let you to read and write to the user's current selection in a "get and set" fashion. To do that, the `Document` object provides the `getSelectedDataAsync` and `setSelectedDataAsync` methods.

For code examples that demonstrate how to perform tasks with selections, see [Read and write data to the active selection in a document or spreadsheet](read-and-write-data-to-the-active-selection-in-a-document-or-spreadsheet.md).

## Work with bindings using the Bindings and Binding objects

Binding-based data access enables content and task pane add-ins to consistently access a particular region of a document or spreadsheet through an identifier associated with a binding. The add-in first needs to establish the binding by calling one of the methods that associates a portion of the document with a unique identifier: [addFromPromptAsync](/javascript/api/office/office.bindings#office-office-bindings-addfrompromptasync-member(1)), [addFromSelectionAsync](/javascript/api/office/office.bindings#office-office-bindings-addfromselectionasync-member(1)), or [addFromNamedItemAsync](/javascript/api/office/office.bindings#office-office-bindings-addfromnameditemasync-member(1)). After the binding is established, the add-in can use the provided identifier to access the data contained in the associated region of the document or spreadsheet. Creating bindings provides the following value to your add-in.

- Permits access to common data structures across supported Office applications, such as: tables, ranges, or text (a contiguous run of characters).

- Enables read/write operations without requiring the user to make a selection.

- Establishes a relationship between the add-in and the data in the document. Bindings are persisted in the document, and can be accessed at a later time.

Establishing a binding also allows you to subscribe to data and selection change events that are scoped to that particular region of the document or spreadsheet. This means that the add-in is only notified of changes that happen within the bound region as opposed to general changes across the whole document or spreadsheet.

The [Bindings](/javascript/api/office/office.bindings) object exposes a [getAllAsync](/javascript/api/office/office.bindings#office-office-bindings-getallasync-member(1)) method that gives access to the set of all bindings established on the document or spreadsheet. An individual binding can be accessed by its ID using either the [Bindings.getBindingByIdAsync](/javascript/api/office/office.bindings#office-office-bindings-getbyidasync-member(1)) method or [Office.select](/javascript/api/office) function. You can establish new bindings as well as remove existing ones by using one of the following methods of the `Bindings` object: [addFromSelectionAsync](/javascript/api/office/office.bindings#office-office-bindings-addfromselectionasync-member(1)), [addFromPromptAsync](/javascript/api/office/office.bindings#office-office-bindings-addfrompromptasync-member(1)), [addFromNamedItemAsync](/javascript/api/office/office.bindings#office-office-bindings-addfromnameditemasync-member(1)), or [releaseByIdAsync](/javascript/api/office/office.bindings#office-office-bindings-releasebyidasync-member(1)).

There are three different types of bindings that you specify with the  _bindingType_ parameter when you create a binding with the `addFromSelectionAsync`, `addFromPromptAsync` or `addFromNamedItemAsync` methods.

|Binding type|Description|Host application support|
|:-----|:-----|:-----|
|Text binding|Binds to a region of the document that can be represented as text.|In Word, most contiguous selections are valid, while in Excel only single cell selections can be the target of a text binding. In Excel, only plain text is supported. In Word, three formats are supported: plain text, HTML, and Open XML for Office.|
|Matrix binding|Binds to a fixed region of a document that contains tabular data without headers. Data in a matrix binding is written or read as a two dimensional **Array**, which in JavaScript is implemented as an array of arrays. For example, two rows of **string** values in two columns can be written or read as ` [['a', 'b'], ['c', 'd']]`, and a single column of three rows can be written or read as `[['a'], ['b'], ['c']]`.|In Excel, any contiguous selection of cells can be used to establish a matrix binding. In Word, only tables support matrix binding.|
|Table binding|Binds to a region of a document that contains a table with headers. Data in a table binding is written or read as a [TableData](/javascript/api/office/office.tabledata) object. The `TableData` object exposes the data through the **headers** and **rows** properties.|Any Excel or Word table can be the basis for a table binding. After you establish a table binding, each new row or column a user adds to the table is automatically included in the binding. |

<br/>

After a binding is created by using one of the three "add" methods of the `Bindings` object, you can work with the binding's data and properties by using the methods of the corresponding object: [MatrixBinding](/javascript/api/office/office.matrixbinding), [TableBinding](/javascript/api/office/office.tablebinding), or [TextBinding](/javascript/api/office/office.textbinding). All three of these objects inherit the [getDataAsync](/javascript/api/office/office.binding#office-office-binding-getdataasync-member(1)) and [setDataAsync](/javascript/api/office/office.binding#office-office-binding-setdataasync-member(1)) methods of the `Binding` object that enable to you interact with the bound data.

For code examples that demonstrate how to perform tasks with bindings, see [Bind to regions in a document or spreadsheet](bind-to-regions-in-a-document-or-spreadsheet.md).

## Work with custom XML parts using the CustomXmlParts and CustomXmlPart objects

 **Applies to:** Task pane add-ins for Word

The [CustomXmlParts](/javascript/api/office/office.customxmlparts) and [CustomXmlPart](/javascript/api/office/office.customxmlpart) objects of the API provide access to custom XML parts in Word documents, which enable XML-driven manipulation of the contents of the document. For demonstrations of working with the `CustomXmlParts` and `CustomXmlPart` objects, see the [Word-add-in-Work-with-custom-XML-parts](https://github.com/OfficeDev/Word-Add-in-Work-with-custom-XML-parts) code sample.

## Work with the entire document using the getFileAsync method

 **Applies to:** Task pane add-ins for Word and PowerPoint

The [Document.getFileAsync](/javascript/api/office/office.document#office-office-document-getfileasync-member(1)) method and members of the [File](/javascript/api/office/office.file) and [Slice](/javascript/api/office/office.slice) objects to provide functionality for getting entire Word and PowerPoint document files in slices (chunks) of up to 4 MB at a time. For more information, see [Get the whole document from an add-in for PowerPoint or Word](../develop/get-the-whole-document-from-an-add-in-for-powerpoint-or-word.md).

## Mailbox object

**Applies to:** Outlook add-ins

[!INCLUDE [Mailbox object information](../includes/mailbox-object-desc.md)]
