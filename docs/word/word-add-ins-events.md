---
title: Work with events using the Word JavaScript API
description: A list of events for Word JavaScript objects. This includes information on using event handlers and the associated patterns.
ms.date: 07/09/2024
ms.localizationpriority: medium
---

# Work with events using the Word JavaScript API

This article describes important concepts related to working with events in Word and provides code samples that show how to register event handlers, handle events, and remove event handlers using the Word JavaScript API.

## Events in Word

Each time certain types of changes occur in a Word document, an event notification fires. By using the Word JavaScript API, you can register event handlers that allow your add-in to automatically run a designated function when a specific event occurs. The following events are currently supported.

| Event | Description | Supported objects |
|:---------------|:-------------|:-----------|
| `onAnnotationClicked` | Occurs when the user clicks an annotation (or selects it using **Alt+Down**). | [**Document**](/javascript/api/word/word.document?view=word-js-preview#word-word-document-onannotationclicked-member) |
| `onAnnotationHovered` | Occurs when the user hovers the cursor over an annotation. | [**Document**](/javascript/api/word/word.document?view=word-js-preview#word-word-document-onannotationhovered-member) |
| `onAnnotationInserted` | Occurs when the user adds one or more annotations. | [**Document**](/javascript/api/word/word.document?view=word-js-preview#word-word-document-onannotationinserted-member) |
| `onAnnotationPopupAction` | Occurs when the user performs an action in an annotation pop-up menu. | [**Document**](/javascript/api/word/word.document?view=word-js-preview#word-word-document-onannotationpopupaction-member) |
| `onAnnotationRemoved` | Occurs when the user deletes one or more annotations. | [**Document**](/javascript/api/word/word.document?view=word-js-preview#word-word-document-onannotationremoved-member) |
| `onContentControlAdded` | Occurs when a content control is added. Run `context.sync()` in the handler to get the new content control's properties. | [**Document**](/javascript/api/word/word.document?view=word-js-preview#word-word-document-oncontentcontroladded-member) |
| `onDataChanged` | Occurs when data within the content control are changed. To get the new text, load this content control in the handler. To get the old text, do not load it. | [**ContentControl**](/javascript/api/word/word.contentcontrol?view=word-js-preview#word-word-contentcontrol-ondatachanged-member) |
| `onDeleted` | Occurs when the content control is deleted. Do not load this content control in the handler, otherwise you won't be able to get its original properties. | [**ContentControl**](/javascript/api/word/word.contentcontrol?view=word-js-preview#word-word-contentcontrol-ondeleted-member) |
| `onEntered` | Occurs when the content control is entered. | [**ContentControl**](/javascript/api/word/word.contentcontrol?view=word-js-preview#word-word-contentcontrol-onentered-member) |
| `onExited` | Occurs when the content control is exited, for example, when the cursor leaves the content control. | [**ContentControl**](/javascript/api/word/word.contentcontrol?view=word-js-preview#word-word-contentcontrol-onexited-member) |
| `onParagraphAdded` | Occurs when the user adds new paragraphs. | [**Document**](/javascript/api/word/word.document?view=word-js-preview#word-word-document-onparagraphadded-member) |
| `onParagraphChanged` | Occurs when the user changes paragraphs. | [**Document**](/javascript/api/word/word.document?view=word-js-preview#word-word-document-onparagraphchanged-member) |
| `onParagraphDeleted` | Occurs when the user deletes paragraphs. | [**Document**](/javascript/api/word/word.document?view=word-js-preview#word-word-document-onparagraphdeleted-member) |
| `onSelectionChanged` | Occurs when selection within the content control is changed. | [**ContentControl**](/javascript/api/word/word.contentcontrol?view=word-js-preview#word-word-contentcontrol-onselectionchanged-member) |

### Events in preview

> [!NOTE]
> The following events are currently available only in public preview. [!INCLUDE [Information about using preview APIs](../includes/using-preview-apis.md)]

| Event | Description | Supported objects |
|:---------------|:-------------|:-----------|
| `onCommentAdded` | Occurs when new comments are added. | <ul><li>[**Body**](/javascript/api/word/word.body?view=word-js-preview#word-word-body-oncommentadded-member)</li><li>[**ContentControl**](/javascript/api/word/word.contentcontrol?view=word-js-preview#word-word-contentcontrol-oncommentadded-member)</li><li>[**Paragraph**](/javascript/api/word/word.paragraph#word-word-paragraph-oncommentadded-member)</li><li>[**Range**](/javascript/api/word/word.range#word-word-range-oncommentadded-member)</li></ul> |
| `onCommentChanged` | Occurs when a comment or its reply is changed. | <ul><li>[**Body**](/javascript/api/word/word.body?view=word-js-preview#word-word-body-oncommentchanged-member)</li><li>[**ContentControl**](/javascript/api/word/word.contentcontrol?view=word-js-preview#word-word-contentcontrol-oncommentchanged-member)</li><li>[**Paragraph**](/javascript/api/word/word.paragraph#word-word-paragraph-oncommentchanged-member)</li><li>[**Range**](/javascript/api/word/word.range#word-word-range-oncommentchanged-member)</li></ul> |
| `onCommentDeleted` | Occurs when comments are deleted. | <ul><li>[**Body**](/javascript/api/word/word.body?view=word-js-preview#word-word-body-oncommentdeleted-member)</li><li>[**Paragraph**](/javascript/api/word/word.paragraph#word-word-paragraph-oncommentdeleted-member)</li></ul> |
| `onCommentDeselected` | Occurs when a comment is deselected. | <ul><li>[**Body**](/javascript/api/word/word.body?view=word-js-preview#word-word-body-oncommentdeselected-member)</li><li>[**ContentControl**](/javascript/api/word/word.contentcontrol?view=word-js-preview#word-word-contentcontrol-oncommentdeselected-member)</li><li>[**Paragraph**](/javascript/api/word/word.paragraph#word-word-paragraph-oncommentdeselected-member)</li><li>[**Range**](/javascript/api/word/word.range#word-word-range-oncommentdeselected-member)</li></ul> |
| `onCommentSelected` | Occurs when a comment is selected. | <ul><li>[**Body**](/javascript/api/word/word.body?view=word-js-preview#word-word-body-oncommentselected-member)</li><li>[**ContentControl**](/javascript/api/word/word.contentcontrol?view=word-js-preview#word-word-contentcontrol-oncommentselected-member)</li><li>[**Paragraph**](/javascript/api/word/word.paragraph#word-word-paragraph-oncommentselected-member)</li><li>[**Range**](/javascript/api/word/word.range#word-word-range-oncommentselected-member)</li></ul> |

### Event triggers

Events within a Word document can be triggered by:

- User interaction via the Word user interface (UI) that changes the document.
- Office Add-in (JavaScript) code that changes the document.
- VBA add-in (macro) code that changes the document.

Any change that complies with default behavior of Word will trigger the corresponding events in a document.

### Lifecycle of an event handler

An event handler is created when an add-in registers the event handler. It's destroyed when the add-in deregisters the event handler or when the add-in is refreshed, reloaded, or closed. Event handlers don't persist as part of the Word file, or across sessions with Word on the web.

> [!CAUTION]
> When an object to which events are registered is deleted (e.g., a table with an `onChanged` event registered), the event handler no longer triggers but remains in memory until the add-in or Word session refreshes or closes.

### Events and coauthoring

With [coauthoring](co-authoring-in-word-add-ins.md), multiple people can work together and edit the same Word document simultaneously. For events that can be triggered by a coauthor, such as `onParagraphChanged`, the corresponding **Event** object will contain a **source** property that indicates whether the event was triggered locally by the current user (`event.source = Local`) or was triggered by the remote coauthor (`event.source = Remote`).

## Register an event handler

The following code sample registers an event handler for the `onParagraphChanged` event in the document. The code specifies that when content changes in the document, the `handleChange` function should run.

```js
await Word.run(async (context) => {
    const worksheet = context.document.worksheets.getItem("Sample");
    worksheet.onChanged.add(handleChange);

    await context.sync();
    console.log("Event handler successfully registered for onChanged event in the worksheet.");
}).catch(errorHandlerFunction);
```

## Handle an event

As shown in the previous example, when you register an event handler, you indicate the function that should run when the specified event occurs. You can design that function to perform whatever actions your scenario requires. The following code sample shows an event handler function that simply writes information about the event to the console.

```js
async function handleChange(event) {
    await Word.run(async (context) => {
        await context.sync();        
        console.log("Change type of event: " + event.changeType);
        console.log("Address of event: " + event.address);
        console.log("Source of event: " + event.source);       
    }).catch(errorHandlerFunction);
}
```

## Remove an event handler

The following code sample registers an event handler for the `onSelectionChanged` event in the document named **Sample** and defines the `handleSelectionChange` function that will run when the event occurs. It also defines the `remove()` function that can subsequently be called to remove that event handler. Note that the `RequestContext` used to create the event handler is needed to remove it.

```js
let eventResult;

async function run() {
  await Word.run(async (context) => {
    const worksheet = context.workbook.worksheets.getItem("Sample");
    eventResult = worksheet.onSelectionChanged.add(handleSelectionChange);

    await context.sync();
    console.log("Event handler successfully registered for onSelectionChanged event in the worksheet.");
  });
}

async function handleSelectionChange(event) {
  await Word.run(async (context) => {
    await context.sync();
    console.log("Address of current selection: " + event.address);
  });
}

async function remove() {
  await Word.run(eventResult.context, async (context) => {
    eventResult.remove();
    await context.sync();
    
    eventResult = null;
    console.log("Event handler successfully removed.");
  });
}
```

## See also

- [Word JavaScript object model in Office Add-ins](word-add-ins-core-concepts.md)
