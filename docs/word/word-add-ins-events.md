---
title: Work with events using the Word JavaScript API
description: A list of events for Word JavaScript objects. This includes information on using event handlers and the associated patterns.
ms.date: 10/21/2026
ms.localizationpriority: medium
---

# Work with events using the Word JavaScript API

This article introduces key concepts for working with events in Word using the JavaScript API. You find practical code samples for registering, handling, and removing event handlers, along with explanations of event life cycles and coauthoring scenarios. Explore the event tables to discover which triggers and objects are supported.

## Events in Word

When certain changes occur in a Word document, event notifications fire. The Word JavaScript APIs let you register event handlers that allow your add-in to automatically run designated functions when those changes occur. The following events are currently supported.

| Event | Description | Supported objects | Triggered during coauthoring? |
|:---------------|:-------------|:-----------|:-----------|
| `onAnnotationClicked` | Occurs when the user selects an annotation.<br><br>Event data object:<br>[AnnotationClickedEventArgs](/javascript/api/word/word.annotationclickedeventargs) | [**Document**](/javascript/api/word/word.document#word-word-document-onannotationclicked-member) | No |
| `onAnnotationHovered` | Occurs when the user hovers the cursor over an annotation.<br><br>Event data object:<br>[AnnotationHoveredEventArgs](/javascript/api/word/word.annotationhoveredeventargs) | [**Document**](/javascript/api/word/word.document#word-word-document-onannotationhovered-member) | No |
| `onAnnotationInserted` | Occurs when the user adds one or more annotations.<br><br>Event data object:<br>[AnnotationInsertedEventArgs](/javascript/api/word/word.annotationinsertedeventargs) | [**Document**](/javascript/api/word/word.document#word-word-document-onannotationinserted-member) | No |
| `onAnnotationPopupAction` | Occurs when the user performs an action in an annotation pop-up menu.<br><br>Event data object:<br>[AnnotationPopupActionEventArgs](/javascript/api/word/word.annotationpopupactioneventargs) | [**Document**](/javascript/api/word/word.document#word-word-document-onannotationpopupaction-member) | No |
| `onAnnotationRemoved` | Occurs when the user deletes one or more annotations.<br><br>Event data object:<br>[AnnotationRemovedEventArgs](/javascript/api/word/word.annotationremovedeventargs) | [**Document**](/javascript/api/word/word.document#word-word-document-onannotationremoved-member) | No |
| `onContentControlAdded` | Occurs when a content control is added. Run `context.sync()` in the handler to get the new content control's properties.<br><br>Event data object:<br>[ContentControlAddedEventArgs](/javascript/api/word/word.contentcontroladdedeventargs) | [**Document**](/javascript/api/word/word.document#word-word-document-oncontentcontroladded-member) | Yes |
| `onDataChanged` | Occurs when data within the content control are changed. To get the new text, load this content control in the handler. To get the old text, don't load it.<br><br>Event data object:<br>[ContentControlDataChangedEventArgs](/javascript/api/word/word.contentcontroldatachangedeventargs) | [**ContentControl**](/javascript/api/word/word.contentcontrol#word-word-contentcontrol-ondatachanged-member) | Yes |
| `onDeleted` | Occurs when the content control is deleted. Don't load this content control in the handler, otherwise you won't be able to get its original properties.<br><br>Event data object:<br>[ContentControlDeletedEventArgs](/javascript/api/word/word.contentcontroldeletedeventargs) | [**ContentControl**](/javascript/api/word/word.contentcontrol#word-word-contentcontrol-ondeleted-member) | Yes |
| `onEntered` | Occurs when the content control is entered.<br><br>Event data object:<br>[ContentControlEnteredEventArgs](/javascript/api/word/word.contentcontrolenteredeventargs) | [**ContentControl**](/javascript/api/word/word.contentcontrol#word-word-contentcontrol-onentered-member) | Yes |
| `onExited` | Occurs when the content control is exited, for example, when the cursor leaves the content control.<br><br>Event data object:<br>[ContentControlExitedEventArgs](/javascript/api/word/word.contentcontrolexitedeventargs) | [**ContentControl**](/javascript/api/word/word.contentcontrol#word-word-contentcontrol-onexited-member) | Yes |
| `onParagraphAdded` | Occurs when the user adds new paragraphs.<br><br>Event data object:<br>[ParagraphAddedEventArgs](/javascript/api/word/word.paragraphaddedeventargs) | [**Document**](/javascript/api/word/word.document#word-word-document-onparagraphadded-member) | Yes |
| `onParagraphChanged` | Occurs when the user changes paragraphs.<br><br>Event data object:<br>[ParagraphChangedEventArgs](/javascript/api/word/word.paragraphchangedeventargs) | [**Document**](/javascript/api/word/word.document#word-word-document-onparagraphchanged-member) | Yes |
| `onParagraphDeleted` | Occurs when the user deletes paragraphs.<br><br>Event data object:<br>[ParagraphDeletedEventArgs](/javascript/api/word/word.paragraphdeletedeventargs) | [**Document**](/javascript/api/word/word.document#word-word-document-onparagraphdeleted-member) | Yes |
| `onSelectionChanged` | Occurs when selection within the content control is changed.<br><br>Event data object:<br>[ContentControlSelectionChangedEventArgs](/javascript/api/word/word.contentcontrolselectionchangedeventargs) | [**ContentControl**](/javascript/api/word/word.contentcontrol#word-word-contentcontrol-onselectionchanged-member) | Yes |

### Events in preview

> [!NOTE]
> The following events are currently available only in public preview. [!INCLUDE [Information about using preview APIs](../includes/using-preview-apis.md)]

| Event | Description | Supported objects | Triggered during coauthoring? |
|:---------------|:-------------|:-----------|:-----------|
| `onCommentAdded` | Occurs when new comments are added.<br><br>Event data object:<br>[CommentEventArgs](/javascript/api/word/word.commenteventargs) | <ul><li>[**Body**](/javascript/api/word/word.body?view=word-js-preview&preserve-view=true&preserve-view=true#word-word-body-oncommentadded-member)</li><li>[**ContentControl**](/javascript/api/word/word.contentcontrol?view=word-js-preview&preserve-view=true#word-word-contentcontrol-oncommentadded-member)</li><li>[**Paragraph**](/javascript/api/word/word.paragraph#word-word-paragraph-oncommentadded-member)</li><li>[**Range**](/javascript/api/word/word.range#word-word-range-oncommentadded-member)</li></ul> | Yes |
| `onCommentChanged` | Occurs when a comment or its reply is changed.<br><br>Event data object:<br>[CommentEventArgs](/javascript/api/word/word.commenteventargs) | <ul><li>[**Body**](/javascript/api/word/word.body?view=word-js-preview&preserve-view=true#word-word-body-oncommentchanged-member)</li><li>[**ContentControl**](/javascript/api/word/word.contentcontrol?view=word-js-preview&preserve-view=true#word-word-contentcontrol-oncommentchanged-member)</li><li>[**Paragraph**](/javascript/api/word/word.paragraph#word-word-paragraph-oncommentchanged-member)</li><li>[**Range**](/javascript/api/word/word.range#word-word-range-oncommentchanged-member)</li></ul> | Yes |
| `onCommentDeleted` | Occurs when comments are deleted.<br><br>Event data object:<br>[CommentEventArgs](/javascript/api/word/word.commenteventargs) | <ul><li>[**Body**](/javascript/api/word/word.body?view=word-js-preview&preserve-view=true#word-word-body-oncommentdeleted-member)</li><li>[**Paragraph**](/javascript/api/word/word.paragraph#word-word-paragraph-oncommentdeleted-member)</li></ul> | Yes |
| `onCommentDeselected` | Occurs when a comment is deselected.<br><br>Event data object:<br>[CommentEventArgs](/javascript/api/word/word.commenteventargs) | <ul><li>[**Body**](/javascript/api/word/word.body?view=word-js-preview&preserve-view=true#word-word-body-oncommentdeselected-member)</li><li>[**ContentControl**](/javascript/api/word/word.contentcontrol?view=word-js-preview&preserve-view=true#word-word-contentcontrol-oncommentdeselected-member)</li><li>[**Paragraph**](/javascript/api/word/word.paragraph#word-word-paragraph-oncommentdeselected-member)</li><li>[**Range**](/javascript/api/word/word.range#word-word-range-oncommentdeselected-member)</li></ul> | Yes |
| `onCommentSelected` | Occurs when a comment is selected.<br><br>Event data object:<br>[CommentEventArgs](/javascript/api/word/word.commenteventargs) | <ul><li>[**Body**](/javascript/api/word/word.body?view=word-js-preview&preserve-view=true#word-word-body-oncommentselected-member)</li><li>[**ContentControl**](/javascript/api/word/word.contentcontrol?view=word-js-preview&preserve-view=true#word-word-contentcontrol-oncommentselected-member)</li><li>[**Paragraph**](/javascript/api/word/word.paragraph#word-word-paragraph-oncommentselected-member)</li><li>[**Range**](/javascript/api/word/word.range#word-word-range-oncommentselected-member)</li></ul> | Yes |

## Event triggers

Events within a Word document can be triggered by:

- Interacting with the Word user interface (UI) in a way that changes the document.
- Running Office Add-in (JavaScript) code that changes the document.
- Running VBA add-in (macro) code that changes the document.
- Coauthoring remotely and changing the document by using the Word UI or add-in code. For more information, see [Events and coauthoring](#events-and-coauthoring).

Any change that follows the default behavior of Word triggers the corresponding events in a document.

## Life cycle of an event handler

An add-in creates an event handler when it registers the event handler. The event handler is destroyed when the add-in deregisters the event handler or when the add-in is refreshed, reloaded, or closed. Event handlers don't persist as part of the Word file, or across sessions with Word on the web.

## Events and coauthoring

With coauthoring, multiple people can work together and edit the same Word document simultaneously. For events that a coauthor can trigger, such as `onParagraphChanged`, the corresponding **Event** object contains a **source** property that indicates whether the current user triggered the event locally (`event.source == Local`) or the remote coauthor triggered the event (`event.source == Remote`).

Events that use the following data objects are triggered during coauthoring.

- `CommentEventArgs` (preview)
- `ContentControlAddedEventArgs`
- `ContentControlDataChangedEventArgs`
- `ContentControlDeletedEventArgs`
- `ContentControlEnteredEventArgs`
- `ContentControlExitedEventArgs`
- `ContentControlSelectionChangedEventArgs`
- `ParagraphAddedEventArgs`
- `ParagraphChangedEventArgs`
- `ParagraphDeletedEventArgs`

## Register an event handler

The following code sample registers an event handler for the `onParagraphChanged` event in the document. The code specifies that when content changes in the document, the `handleChange` function runs.

```js
await Word.run(async (context) => {
    eventContext = context.document.onParagraphChanged.add(handleChange);

    await context.sync();
    console.log("Event handler successfully registered for onParagraphChanged event in the document.");
}).catch(errorHandlerFunction);
```

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

The following code sample registers an event handler for the `onParagraphChanged` event in the document and defines the `handleChange` function that runs when the event occurs. It also defines the `deregisterEventHandler()` function that you can call later to remove that event handler. You need the `RequestContext` used to create the event handler to remove it.

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
  // The `RequestContext` used to create the event handler is needed to remove it.
  // In this example, `eventContext` is being used to keep track of that context.
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
  - `onDataChanged`
  - `onDeleted`
  - `onEntered`
  - `onExited`
  - `onSelectionChanged`
- Comment events (preview)
  - `onCommentAdded`
  - `onCommentChanged`
  - `onCommentDeleted`
  - `onCommentDeselected`
  - `onCommentSelected`

The following code sample shows how to register an event handler on each content control. Because you're adding the event to the content controls, `track()` is called on each content control in the collection.

```typescript
let eventContexts = [];

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

The following code sample registers all five preview comment event handlers on the document's body and removes them when they're no longer needed. It retains one tracked `Word.Body` and the registration handles for cleanup. Call `registerCommentEventHandlers()` to register the handlers, then call `deregisterCommentEventHandlers()` to remove them.

For these body comment events, call `body.track()` and then `await context.sync()` **before** adding the handlers. This synchronization resolves the body's host-assigned reference ID, which Word on the web uses to route comment events. Keep the second synchronization after adding the handlers to complete registration.

The example prevents duplicate registrations and overlapping registration and removal calls. Cleanup uses the original request context, removes all retained handlers, and untracks the shared body once. Errors are logged to the console. If registration fails, run cleanup before trying to register again.

```typescript
let commentBody: Word.Body | null = null;
const commentEventHandlers: OfficeExtension.EventHandlerResult<Word.CommentEventArgs>[] = [];
let commentEventsBusy = false;

async function registerCommentEventHandlers() {
  if (commentEventsBusy || commentBody !== null) {
    console.warn("Wait for the current operation and remove existing comment handlers before registering again.");
    return;
  }

  commentEventsBusy = true;
  try {
    await Word.run(async (context) => {
      const body: Word.Body = context.document.body;
      commentBody = body;

      body.track();
      // Resolve the tracked body's host reference before subscribing.
      await context.sync();

      commentEventHandlers.push(
        body.onCommentAdded.add(onEventHandler),
        body.onCommentChanged.add(onChangedHandler),
        body.onCommentDeleted.add(onEventHandler),
        body.onCommentDeselected.add(onEventHandler),
        body.onCommentSelected.add(onEventHandler)
      );
      await context.sync();

      console.log("Comment event handlers registered.");
    });
  } catch (error) {
    console.error(error);
  } finally {
    commentEventsBusy = false;
  }
}

async function deregisterCommentEventHandlers() {
  if (commentEventsBusy) {
    console.warn("Wait for the current comment event operation to finish before removing handlers.");
    return;
  }
  if (commentBody === null) {
    console.warn("There are no comment event registrations to remove.");
    return;
  }

  const body: Word.Body = commentBody;
  commentEventsBusy = true;
  try {
    // Passing the tracked body reuses the request context for all five handlers.
    await Word.run(body, async (context) => {
      for (const handler of commentEventHandlers) {
        handler.remove();
      }
      body.untrack();
      await context.sync();
    });

    commentEventHandlers.length = 0;
    commentBody = null;
    console.log("Comment event handlers removed and body untracked.");
  } catch (error) {
    console.error(error);
  } finally {
    commentEventsBusy = false;
  }
}
```

The callbacks below log the event type, source, and affected comment and reply IDs. The changed-event callback also reports the kind of change. These values are included in the event payload, so logging them doesn't require a `Word.run` batch or `context.sync()`.

```typescript
async function onEventHandler(event: Word.CommentEventArgs) {
  console.log(`${event.type} event. Source: ${event.source}.`);
  if (event.source === Word.EventSource.remote) {
    console.log("A coauthor triggered this event.");
  }
  event.commentDetails.forEach((comment: Word.CommentDetail) => {
    console.log("Affected comment:", { id: comment.id, replyIds: comment.replyIds });
  });
}

async function onChangedHandler(event: Word.CommentEventArgs) {
  console.log(`Comment change: ${event.changeType}.`);
  if (event.changeType === Word.CommentChangeType.edited) {
    console.log("The comment text was edited.");
  }
  await onEventHandler(event);
}
```

For a complete runnable example, try [Comment events (preview)](https://github.com/OfficeDev/office-js-snippets/blob/main/samples/word/99-preview-apis/comment-events.yaml) in [Script Lab](../overview/explore-with-script-lab.md). It provides deterministic setup by replacing the document body with a sample paragraph, plus create, edit, and delete actions to trigger the events. To exercise selection and deselection, select a comment in Word's Comments pane, then click elsewhere in the document. The sample requires the preview Office.js library and TypeScript definitions described in [Events in preview](#events-in-preview), and a Word client that supports these preview APIs.

## See also

- [Word JavaScript object model in Office Add-ins](word-add-ins-core-concepts.md)
- These and other examples are available in our [Script Lab](../overview/explore-with-script-lab.md) tool:
  - [On changing content in paragraphs](https://github.com/OfficeDev/office-js-snippets/blob/prod/samples/word/25-paragraph/onchanged-event.yaml)
  - [On deleting content controls](https://github.com/OfficeDev/office-js-snippets/blob/prod/samples/word/10-content-controls/content-control-ondeleted-event.yaml)
  - [Manage comments](https://github.com/OfficeDev/office-js-snippets/blob/prod/samples/word/99-preview-apis/manage-comments.yaml) (preview)
