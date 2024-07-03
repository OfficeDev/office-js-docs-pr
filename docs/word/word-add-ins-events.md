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
| `onAnnotationClicked` | Occurs when the user clicks an annotation (or selects it using **Alt+Down**).<br><br>Event data object:<br>[AnnotationClickedEventArgs](/javascript/api/word/word.annotationclickedeventargs) | [**Document**](/javascript/api/word/word.document#word-word-document-onannotationclicked-member) |
| `onAnnotationHovered` | Occurs when the user hovers the cursor over an annotation.<br><br>Event data object:<br>[AnnotationHoveredEventArgs](/javascript/api/word/word.annotationhoveredeventargs) | [**Document**](/javascript/api/word/word.document#word-word-document-onannotationhovered-member) |
| `onAnnotationInserted` | Occurs when the user adds one or more annotations.<br><br>Event data object:<br>[AnnotationInsertedEventArgs](/javascript/api/word/word.annotationinsertedeventargs) | [**Document**](/javascript/api/word/word.document#word-word-document-onannotationinserted-member) |
| `onAnnotationPopupAction` | Occurs when the user performs an action in an annotation pop-up menu.<br><br>Event data object:<br>[AnnotationPopupActionEventArgs](/javascript/api/word/word.annotationpopupactioneventargs) | [**Document**](/javascript/api/word/word.document#word-word-document-onannotationpopupaction-member) |
| `onAnnotationRemoved` | Occurs when the user deletes one or more annotations.<br><br>Event data object:<br>[AnnotationRemovedEventArgs](/javascript/api/word/word.annotationremovedeventargs) | [**Document**](/javascript/api/word/word.document#word-word-document-onannotationremoved-member) |
| `onContentControlAdded` | Occurs when a content control is added. Run `context.sync()` in the handler to get the new content control's properties.<br><br>Event data object:<br>[ContentControlAddedEventArgs](/javascript/api/word/word.contentcontroladdedeventargs) | [**Document**](/javascript/api/word/word.document#word-word-document-oncontentcontroladded-member) |
| `onDataChanged` | Occurs when data within the content control are changed. To get the new text, load this content control in the handler. To get the old text, do not load it.<br><br>Event data object:<br>[ContentControlDataChangedEventArgs](/javascript/api/word/word.contentcontroldatachangedeventargs) | [**ContentControl**](/javascript/api/word/word.contentcontrol#word-word-contentcontrol-ondatachanged-member) |
| `onDeleted` | Occurs when the content control is deleted. Do not load this content control in the handler, otherwise you won't be able to get its original properties.<br><br>Event data object:<br>[ContentControlDeletedEventArgs](/javascript/api/word/word.contentcontroldeletedeventargs) | [**ContentControl**](/javascript/api/word/word.contentcontrol#word-word-contentcontrol-ondeleted-member) |
| `onEntered` | Occurs when the content control is entered.<br><br>Event data object:<br>[ContentControlEnteredEventArgs](/javascript/api/word/word.contentcontrolenteredeventargs) | [**ContentControl**](/javascript/api/word/word.contentcontrol#word-word-contentcontrol-onentered-member) |
| `onExited` | Occurs when the content control is exited, for example, when the cursor leaves the content control.<br><br>Event data object:<br>[ContentControlExitedEventArgs](/javascript/api/word/word.contentcontrolexitedeventargs) | [**ContentControl**](/javascript/api/word/word.contentcontrol#word-word-contentcontrol-onexited-member) |
| `onParagraphAdded` | Occurs when the user adds new paragraphs.<br><br>Event data object:<br>[ParagraphAddedEventArgs](/javascript/api/word/word.paragraphaddedeventargs) | [**Document**](/javascript/api/word/word.document#word-word-document-onparagraphadded-member) |
| `onParagraphChanged` | Occurs when the user changes paragraphs.<br><br>Event data object:<br>[ParagraphChangedEventArgs](/javascript/api/word/word.paragraphchangedeventargs) | [**Document**](/javascript/api/word/word.document#word-word-document-onparagraphchanged-member) |
| `onParagraphDeleted` | Occurs when the user deletes paragraphs.<br><br>Event data object:<br>[ParagraphDeletedEventArgs](/javascript/api/word/word.paragraphdeletedeventargs) | [**Document**](/javascript/api/word/word.document#word-word-document-onparagraphdeleted-member) |
| `onSelectionChanged` | Occurs when selection within the content control is changed.<br><br>Event data object:<br>[ContentControlSelectionChangedEventArgs](/javascript/api/word/word.contentcontrolselectionchangedeventargs) | [**ContentControl**](/javascript/api/word/word.contentcontrol#word-word-contentcontrol-onselectionchanged-member) |

### Events in preview

> [!NOTE]
> The following events are currently available only in public preview. [!INCLUDE [Information about using preview APIs](../includes/using-preview-apis.md)]

| Event | Description | Supported objects |
|:---------------|:-------------|:-----------|
| `onCommentAdded` | Occurs when new comments are added.<br><br>Event data object:<br>[CommentEventArgs](/javascript/api/word/word.commenteventargs) | <ul><li>[**Body**](/javascript/api/word/word.body?view=word-js-preview&preserve-view=true&preserve-view=true#word-word-body-oncommentadded-member)</li><li>[**ContentControl**](/javascript/api/word/word.contentcontrol?view=word-js-preview&preserve-view=true#word-word-contentcontrol-oncommentadded-member)</li><li>[**Paragraph**](/javascript/api/word/word.paragraph#word-word-paragraph-oncommentadded-member)</li><li>[**Range**](/javascript/api/word/word.range#word-word-range-oncommentadded-member)</li></ul> |
| `onCommentChanged` | Occurs when a comment or its reply is changed.<br><br>Event data object:<br>[CommentEventArgs](/javascript/api/word/word.commenteventargs) | <ul><li>[**Body**](/javascript/api/word/word.body?view=word-js-preview&preserve-view=true#word-word-body-oncommentchanged-member)</li><li>[**ContentControl**](/javascript/api/word/word.contentcontrol?view=word-js-preview&preserve-view=true#word-word-contentcontrol-oncommentchanged-member)</li><li>[**Paragraph**](/javascript/api/word/word.paragraph#word-word-paragraph-oncommentchanged-member)</li><li>[**Range**](/javascript/api/word/word.range#word-word-range-oncommentchanged-member)</li></ul> |
| `onCommentDeleted` | Occurs when comments are deleted.<br><br>Event data object:<br>[CommentEventArgs](/javascript/api/word/word.commenteventargs) | <ul><li>[**Body**](/javascript/api/word/word.body?view=word-js-preview&preserve-view=true#word-word-body-oncommentdeleted-member)</li><li>[**Paragraph**](/javascript/api/word/word.paragraph#word-word-paragraph-oncommentdeleted-member)</li></ul> |
| `onCommentDeselected` | Occurs when a comment is deselected.<br><br>Event data object:<br>[CommentEventArgs](/javascript/api/word/word.commenteventargs) | <ul><li>[**Body**](/javascript/api/word/word.body?view=word-js-preview&preserve-view=true#word-word-body-oncommentdeselected-member)</li><li>[**ContentControl**](/javascript/api/word/word.contentcontrol?view=word-js-preview&preserve-view=true#word-word-contentcontrol-oncommentdeselected-member)</li><li>[**Paragraph**](/javascript/api/word/word.paragraph#word-word-paragraph-oncommentdeselected-member)</li><li>[**Range**](/javascript/api/word/word.range#word-word-range-oncommentdeselected-member)</li></ul> |
| `onCommentSelected` | Occurs when a comment is selected.<br><br>Event data object:<br>[CommentEventArgs](/javascript/api/word/word.commenteventargs) | <ul><li>[**Body**](/javascript/api/word/word.body?view=word-js-preview&preserve-view=true#word-word-body-oncommentselected-member)</li><li>[**ContentControl**](/javascript/api/word/word.contentcontrol?view=word-js-preview&preserve-view=true#word-word-contentcontrol-oncommentselected-member)</li><li>[**Paragraph**](/javascript/api/word/word.paragraph#word-word-paragraph-oncommentselected-member)</li><li>[**Range**](/javascript/api/word/word.range#word-word-range-oncommentselected-member)</li></ul> |

### Event triggers

Events within a Word document can be triggered by:

- User interaction via the Word user interface (UI) that changes the document.
- Office Add-in (JavaScript) code that changes the document.
- VBA add-in (macro) code that changes the document.

Any change that complies with default behavior of Word will trigger the corresponding events in a document.

### Lifecycle of an event handler

An event handler is created when an add-in registers the event handler. It's destroyed when the add-in deregisters the event handler or when the add-in is refreshed, reloaded, or closed. Event handlers don't persist as part of the Word file, or across sessions with Word on the web.

> [!CAUTION]
> When an object to which events are registered is deleted (e.g., a content control with an `onDataChanged` event registered), the event handler no longer triggers but remains in memory until the add-in or Word session refreshes or closes.

### Events and coauthoring

With coauthoring, multiple people can work together and edit the same Word document simultaneously. For events that can be triggered by a coauthor, such as `onParagraphChanged`, the corresponding **Event** object will contain a **source** property that indicates whether the event was triggered locally by the current user (`event.source = Local`) or was triggered by the remote coauthor (`event.source = Remote`).

## Register an event handler

The following code sample registers an event handler for the `onParagraphChanged` event in the document. The code specifies that when content changes in the document, the `handleChange` function should run.

```js
await Word.run(async (context) => {
    eventContext = context.document.onParagraphChanged.add(handleChange);

    await context.sync();
    console.log("Event handler successfully registered for onParagraphChanged event in the document.");
}).catch(errorHandlerFunction);
```

## Handle an event

As shown in the previous example, when you register an event handler, you indicate the function that should run when the specified event occurs. You can design that function to perform whatever actions your scenario requires. The following code sample shows an event handler function that simply writes information about the event to the console.

```js
async function handleChange(event) {
    await Word.run(async (context) => {
        await context.sync();        
        console.log("Type of event: " + event.type);
        console.log("Source of event: " + event.source);       
    }).catch(errorHandlerFunction);
}
```

## Remove an event handler

The following code sample registers an event handler for the `onParagraphChanged` event in the document and defines the `handleChange` function that will run when the event occurs. It also defines the `deregisterEventHandler()` function that can subsequently be called to remove that event handler. Note that the `RequestContext` used to create the event handler is needed to remove it.

```js
let eventContext;

async function registerEventHandler() {
  await Word.run(async (context) => {
    eventContext = context.document.onParagraphChanged.add(handleChange);

    await context.sync();
    console.log("Event handler successfully registered for onParagraphChanged event in the document.");
  });
}

async function handleChange(event: Word.ParagraphChangedEventArgs) {
  await Word.run(async (context) => {
    await context.sync();
    console.log(`${event.type} event was detected.`);
  });
}

async function deregisterEventHandler() {
  await Word.run(eventContext.context, async (context) => {
    eventContext.remove();
    await context.sync();
    
    eventContext = null;
    console.log("Removed event handler that was tracking content changes in paragraphs.");
  });
}
```

## Use .track()

Certain event types also require you to call `track()` on the object you're adding the event to.

- Content control events
  - onDataChanged
  - onDeleted
  - onEntered
  - onExited
  - onSelectionChanged
- Comment events (preview)
  - onCommentAdded
  - onCommentChanged
  - onCommentDeleted
  - onCommentDeselected
  - onCommentSelected

The following code sample shows how to register an event handler on each content control. Because you're adding the event to the content controls, `track()` is called on each content control in the collection.

```typescript
await Word.run(async (context) => {
  const contentControls: Word.ContentControlCollection = context.document.contentControls;
  contentControls.load("items");
  await context.sync();

  // Register the onDeleted event handler on each content control.
  if (contentControls.items.length === 0) {
    console.log("There aren't any content controls in this document so can't register event handlers.");
  } else {
    for (let i = 0; i < contentControls.items.length; i++) {
      eventContexts[i] = contentControls.items[i].onDeleted.add(contentControlDeleted);

      // Call track() on each content control.
      contentControls.items[i].track();
    }

    await context.sync();

    console.log("Added event handlers for when content controls are deleted.");
  }
});
```

The following code sample shows how to register comment event handlers on the document's body object and include a `body.track();` statement.

```typescript
// Registers event handlers.
await Word.run(async (context) => {
  const body: Word.Body = context.document.body;

  // Track the body object since you're adding comment events to it.
  body.track();
  await context.sync();

  eventContexts[0] = body.onCommentAdded.add(onEventHandler);
  eventContexts[1] = body.onCommentChanged.add(onChangedHandler);
  eventContexts[2] = body.onCommentDeleted.add(onEventHandler);
  eventContexts[3] = body.onCommentDeselected.add(onEventHandler);
  eventContexts[4] = body.onCommentSelected.add(onEventHandler);
  await context.sync();

  console.log("Event handlers registered.");
});
```

## See also

- [Word JavaScript object model in Office Add-ins](word-add-ins-core-concepts.md)
- These and other examples are available in our [Script Lab](../overview/explore-with-script-lab.md) tool:
  - [On changing content in paragraphs](https://github.com/OfficeDev/office-js-snippets/blob/prod/samples/word/25-paragraph/onchanged-event.yaml)
  - [On deleting content controls](https://github.com/OfficeDev/office-js-snippets/blob/prod/samples/word/10-content-controls/content-control-ondeleted-event.yaml)
  - [Manage comments](https://github.com/OfficeDev/office-js-snippets/blob/prod/samples/word/99-preview-apis/manage-comments.yaml) (preview)
