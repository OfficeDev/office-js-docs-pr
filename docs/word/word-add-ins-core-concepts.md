---
title: Word JavaScript object model in Office Add-ins
description: Learn how documents, bodies, ranges, paragraphs, and content controls fit together in the Word JavaScript API.
ms.date: 07/23/2026
ms.topic: concept-article
ms.localizationpriority: high
---

# Word JavaScript object model in Office Add-ins

To effectively read and edit Word documents with your add-in, you need to understand the Word JavaScript object model. This article explains how the main Word objects fit together so you can choose the right patterns.

## Office.js APIs for Word

A Word add-in interacts with objects in Word by using the Office JavaScript API. There are two Office JavaScript object models:

- **Word JavaScript API**: The [Word JavaScript API](/javascript/api/word) provides strongly typed objects that work with the document, ranges, tables, lists, formatting, and more. To learn how `Word.run`, proxy objects, and `context.sync()` work with the document, see [Using the application-specific API model](../develop/application-specific-api-model.md).

- **Common APIs**: The [Common API](/javascript/api/office) gives access to features such as UI, dialogs, and client settings that are common across multiple Office applications. To learn more about using the Common API, see [Common JavaScript API object model](../develop/office-javascript-api-object-model.md).

You'll likely use the Word JavaScript API to develop most of the functionality in add-ins that target Word. However, you might also use objects in the Common API for runtime information, requirement-set checks, dialogs, settings, selection APIs, bindings, and file access. For example:

- [Office.Context](/javascript/api/office/office.context): The `Context` object represents the add-in runtime environment. Use it to inspect details such as `contentLanguage`, `officeTheme`, `host`, and `platform`. You can also call `requirements.isSetSupported()` to check whether the current Word client supports a requirement set.
- [Office.Document](/javascript/api/office/office.document): The `Office.Document` object provides the `getFileAsync()` method, which you can use to download the Word file where the add-in is running. This object is separate from the [Word.Document](/javascript/api/word/word.document) object.

:::image type="content" source="../images/word-js-api-common-api.png" alt-text="Differences between the Word JS API and Common APIs.":::

## Start with the Word objects you use most

In the Word object model, start with the document and move inward to the content you want to work with.

| Object | Description |
| -------- | ------------- |
| [Document](/javascript/api/word/word.document) | The top-level object. Provides access to the body, sections, content controls, settings, and other document-wide properties. |
| [Body](/javascript/api/word/word.body) | The main document body. Use it to read, insert, search, and format content. |
| [Paragraph](/javascript/api/word/word.paragraph) | A paragraph in the document. Use it to work with paragraph text, formatting, and structure. |
| [Range](/javascript/api/word/word.range) | A contiguous area of content. Use it to read or update selected text, tables, images, or whitespace. |
| [ContentControl](/javascript/api/word/word.contentcontrol) | A structured region of content that you can identify, protect, and update. |
| [Table](/javascript/api/word/word.table) | A table in the document. Use it to read or update tabular content. |
| [List](/javascript/api/word/word.list) | A numbered or bulleted list. Use it when your add-in works with list structure or formatting. |
| [Window](/javascript/api/word/word.window) | The window that displays the document. Use it to work with the document's visible container. |
| [Pane](/javascript/api/word/word.pane) | A pane within a window. Use it when your add-in needs to work with the visible area that surrounds the document content. |

To explore these patterns in more depth, see the following related articles and samples:

- [Sample: Import a Word document template with a Word add-in](import-template.md)
- [Sample: Manage citations in a Word document using your Word add-in](citation-management.md)
- [Work with events using the Word JavaScript API](word-add-ins-events.md)
- [Use search options in your Word add-in to find text](search-option-guidance.md)

## Word-specific object model

To understand the Word APIs, you need to understand how key components of a document relate to one another.

- The document contains sections, pages, and document-level entities such as settings and custom XML parts.
- A section contains a body.
- A body contains paragraphs, content controls, ranges, tables, and inline pictures.
- A range is a contiguous area of content, including text, whitespace, tables, and images. The [Word.Range](/javascript/api/word/word.range) object contains many of the text manipulation methods that add-ins use most often.
- A list contains numbered or bulleted paragraphs.
- A window displays the document.
- A window has panes. A pane surrounds the visible area of the document.

For the full set of objects supported by the Word JavaScript API, see [Word JavaScript API](/javascript/api/word).

## See also

- [Word JavaScript API reference](/javascript/api/word)
- [Build your first Word add-in](../quickstarts/word-quickstart-yo.md)
- [Word add-in tutorial](../tutorials/word-tutorial.md)
- [Learn about the Microsoft 365 Developer Program](https://developer.microsoft.com/microsoft-365/dev-program)
