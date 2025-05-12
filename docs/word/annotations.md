---
title: Use annotations in your Word add-in
description: Learn to use annotations in your Word add-in.
ms.date: 04/29/2025
ms.localizationpriority: medium
ms.topic: how-to
---

# Use annotations in your Word add-in

**Includes community contributions from:** [Abdulhadi Jarad](https://github.com/abdulhadi-jarad)

You can use annotations to provide feedback about grammar or other aspects of content in a Word document. The user may see colorful underlining that indicates there's an issue or other information. If the user hovers over the affected content, a popup dialog is displayed that shows them what the issue is and possible actions they can take.

APIs for working with annotations were introduced in the [WordApi 1.7 requirement set](/javascript/api/requirement-sets/word/word-api-1-7-requirement-set) and expanded in the [WordApi 1.8 requirement set](/javascript/api/requirement-sets/word/word-api-1-8-requirement-set) as part of supporting writing assistance scenarios like checking spelling and grammar or providing suggestions to improve writing.

In this article, we show how your add-in can insert feedback and critiques using annotations in a document and allow the user to react to them.

> [!IMPORTANT]
> These annotations aren't persisted in the document. This means that when the document is reopened, the annotations need to be regenerated. However, if the user accepts suggested changes, the changes will persist as long as the user saves them before closing the document.

## Prerequisites

The annotation APIs rely on a service that requires a Microsoft 365 subscription. As such, using this feature in Word with a one-time purchase license won't work. The user must be running Word connected to a Microsoft 365 subscription so that your add-in can successfully run the annotation APIs.

## Key annotation APIs

The following are the key annotation APIs.

- [Paragraph.insertAnnotations](/javascript/api/word/word.paragraph#word-word-paragraph-insertannotations-member(1))
- [Paragraph.getAnnotations](/javascript/api/word/word.paragraph#word-word-paragraph-getannotations-member(1))
- Objects:
  - [Annotation](/javascript/api/word/word.annotation): Represents an annotation.
  - [AnnotationCollection](/javascript/api/word/word.annotationcollection): Represents the collection of annotations.
  - [AnnotationSet](/javascript/api/word/word.annotationset): Represents the set of annotations produced by your add-in in this session.
  - [CritiqueAnnotation](/javascript/api/word/word.critiqueannotation): Represents the critique type of annotation.
  - [Critique](/javascript/api/word/word.critique): Represents feedback about an affected area of a paragraph, indicated by a colored underline.
- Annotation events on [Document](/javascript/api/word/word.document):
  - [onAnnotationClicked](/javascript/api/word/word.document#word-word-document-onannotationclicked-member)
  - [onAnnotationHovered](/javascript/api/word/word.document#word-word-document-onannotationhovered-member)
  - [onAnnotationInserted](/javascript/api/word/word.document#word-word-document-onannotationinserted-member)
  - [onAnnotationPopupAction](/javascript/api/word/word.document#word-word-document-onannotationpopupaction-member)
  - [onAnnotationRemoved](/javascript/api/word/word.document#word-word-document-onannotationremoved-member)

## Use annotation APIs

The following sections show how to work with annotation APIs. These examples are based on the [Manage annotations](https://github.com/OfficeDev/office-js-snippets/blob/prod/samples/word/50-document/manage-annotations.yaml) code sample.

Your add-in code should use the feedback or critique results from your service run against the user's content in the document to manage annotations more dynamically.

## Register annotation events

The following code shows how to register event handlers. To learn more about working with events in Word, see [Work with events using the Word JavaScript API](word-add-ins-events.md). For examples of annotation event handlers, see the following sections.

```typescript
let eventContexts = [];

async function registerEventHandlers() {
  // Registers event handlers.
  await Word.run(async (context) => {
    eventContexts[0] = context.document.onParagraphAdded.add(paragraphChanged);
    eventContexts[1] = context.document.onParagraphChanged.add(paragraphChanged);

    eventContexts[2] = context.document.onAnnotationClicked.add(onClickedHandler);
    eventContexts[3] = context.document.onAnnotationHovered.add(onHoveredHandler);
    eventContexts[4] = context.document.onAnnotationInserted.add(onInsertedHandler);
    eventContexts[5] = context.document.onAnnotationRemoved.add(onRemovedHandler);
    eventContexts[6] = context.document.onAnnotationPopupAction.add(onPopupActionHandler);

    await context.sync();

    console.log("Event handlers registered.");
  });
}
```

### onClickedHandler event handler

The following code runs when the registered `onAnnotationClicked` event occurs.

```typescript
async function onClickedHandler(args: Word.AnnotationClickedEventArgs) {
  // Runs when the registered Document.onAnnotationClicked event occurs.
  await Word.run(async (context) => {
    const annotation: Word.Annotation = context.document.getAnnotationById(args.id);
    annotation.load("critiqueAnnotation");

    await context.sync();

    console.log(`AnnotationClicked: ID ${args.id}:`, annotation.critiqueAnnotation.critique);
  });
}
```

### onHoveredHandler event handler

The following code runs when the registered `onAnnotationHovered` event occurs.

```typescript
async function onHoveredHandler(args: Word.AnnotationHoveredEventArgs) {
  // Runs when the registered Document.onAnnotationHovered event occurs.
  await Word.run(async (context) => {
    const annotation: Word.Annotation = context.document.getAnnotationById(args.id);
    annotation.load("critiqueAnnotation");

    await context.sync();

    console.log(`AnnotationHovered: ID ${args.id}:`, annotation.critiqueAnnotation.critique);
  });
}
```

### onInsertedHandler event handler

The following code runs when the registered `onAnnotationInserted` event occurs.

```typescript
async function onInsertedHandler(args: Word.AnnotationInsertedEventArgs) {
  // Runs when the registered Document.onAnnotationInserted event occurs.
  await Word.run(async (context) => {
    const annotations = [];
    for (let i = 0; i < args.ids.length; i++) {
      let annotation: Word.Annotation = context.document.getAnnotationById(args.ids[i]);
      annotation.load("id,critiqueAnnotation");
      annotations.push(annotation);
    }

    await context.sync();

    for (let annotation of annotations) {
      console.log(`AnnotationInserted: ID ${annotation.id}:`, annotation.critiqueAnnotation.critique);
    }
  });
}
```

### onRemovedHandler event handler

The following code runs when the registered `onAnnotationRemoved` event occurs.

```typescript
async function onRemovedHandler(args: Word.AnnotationRemovedEventArgs) {
  // Runs when the registered Document.onAnnotationRemoved event occurs.
  await Word.run(async (context) => {
    for (let id of args.ids) {
      console.log(`AnnotationRemoved: ID ${id}`);
    }
  });
}
```

### onPopupActionHandler event handler

The following code runs when the registered `onAnnotationPopupAction` event occurs.

```typescript
async function onPopupActionHandler(args: Word.AnnotationPopupActionEventArgs) {
  // Runs when the registered Document.onAnnotationPopupAction event occurs.
  await Word.run(async (context) => {
    let message = `AnnotationPopupAction: ID ${args.id} = `;
    if (args.action === "Accept") {
      message += `Accepted: ${args.critiqueSuggestion}`;
    } else {
      message += "Rejected";
    }

    console.log(message);
  });
}
```

## Insert annotations

The following code shows how to insert annotations into the selected paragraph.

```typescript
async function insertAnnotations() {
  // Adds annotations to the selected paragraph.
  await Word.run(async (context) => {
    const paragraph: Word.Paragraph = context.document.getSelection().paragraphs.getFirst();
    const options: Word.CritiquePopupOptions = {
      brandingTextResourceId: "PG.TabLabel",
      subtitleResourceId: "PG.HelpCommand.TipTitle",
      titleResourceId: "PG.HelpCommand.Label",
      suggestions: ["suggestion 1", "suggestion 2", "suggestion 3"]
    };
    const critique1: Word.Critique = {
      colorScheme: Word.CritiqueColorScheme.red,
      start: 1,
      length: 3,
      popupOptions: options
    };
    const critique2: Word.Critique = {
      colorScheme: Word.CritiqueColorScheme.green,
      start: 6,
      length: 1,
      popupOptions: options
    };
    const critique3: Word.Critique = {
      colorScheme: Word.CritiqueColorScheme.blue,
      start: 10,
      length: 3,
      popupOptions: options
    };
    const critique4: Word.Critique = {
      colorScheme: Word.CritiqueColorScheme.lavender,
      start: 14,
      length: 3,
      popupOptions: options
    };
    const critique5: Word.Critique = {
      colorScheme: Word.CritiqueColorScheme.berry,
      start: 18,
      length: 10,
      popupOptions: options
    };
    const annotationSet: Word.AnnotationSet = {
      critiques: [critique1, critique2, critique3, critique4, critique5]
    };

    const annotationIds = paragraph.insertAnnotations(annotationSet);

    await context.sync();

    console.log("Annotations inserted:", annotationIds.value);
  });
}
```

## Get annotations

The following code shows how to get annotations from the selected paragraph.

```typescript
async function getAnnotations() {
  // Gets annotations found in the selected paragraph.
  await Word.run(async (context) => {
    const paragraph: Word.Paragraph = context.document.getSelection().paragraphs.getFirst();
    const annotations: Word.AnnotationCollection = paragraph.getAnnotations();
    annotations.load("id,state,critiqueAnnotation");

    await context.sync();

    console.log("Annotations found:");

    for (let i = 0; i < annotations.items.length; i++) {
      const annotation: Word.Annotation = annotations.items[i];

      console.log(`ID ${annotation.id} - state '${annotation.state}':`, annotation.critiqueAnnotation.critique);
    }
  });
}
```

## Accept an annotation

The following code shows how to accept the first annotation found in the selected paragraph.

```typescript
async function acceptFirst() {
  // Accepts the first annotation found in the selected paragraph.
  await Word.run(async (context) => {
    const paragraph: Word.Paragraph = context.document.getSelection().paragraphs.getFirst();
    const annotations: Word.AnnotationCollection = paragraph.getAnnotations();
    annotations.load("id,state,critiqueAnnotation");

    await context.sync();

    for (let i = 0; i < annotations.items.length; i++) {
      const annotation: Word.Annotation = annotations.items[i];

      if (annotation.state === Word.AnnotationState.created) {
        console.log(`Accepting ID ${annotation.id}...`);
        annotation.critiqueAnnotation.accept();

        await context.sync();
        break;
      }
    }
  });
}
```

## Reject an annotation

The following code shows how to reject the last annotation found in the selected paragraph.

```typescript
async function rejectLast() {
  // Rejects the last annotation found in the selected paragraph.
  await Word.run(async (context) => {
    const paragraph: Word.Paragraph = context.document.getSelection().paragraphs.getFirst();
    const annotations: Word.AnnotationCollection = paragraph.getAnnotations();
    annotations.load("id,state,critiqueAnnotation");

    await context.sync();

    for (let i = annotations.items.length - 1; i >= 0; i--) {
      const annotation: Word.Annotation = annotations.items[i];

      if (annotation.state === Word.AnnotationState.created) {
        console.log(`Rejecting ID ${annotation.id}...`);
        annotation.critiqueAnnotation.reject();

        await context.sync();
        break;
      }
    }
  });
}
```

## Delete annotations

The following code shows how to delete all the annotations found in the selected paragraph.

```typescript
async function deleteAnnotations() {
  // Deletes all annotations found in the selected paragraph.
  await Word.run(async (context) => {
    const paragraph: Word.Paragraph = context.document.getSelection().paragraphs.getFirst();
    const annotations: Word.AnnotationCollection = paragraph.getAnnotations();
    annotations.load("id");

    await context.sync();

    const ids = [];
    for (let i = 0; i < annotations.items.length; i++) {
      const annotation: Word.Annotation = annotations.items[i];

      ids.push(annotation.id);
      annotation.delete();
    }

    await context.sync();

    console.log("Annotations deleted:", ids);
  });
}
```

## Deregister annotation events

The following code shows how to deregister the event handlers using their event contexts you tracked in the `eventContext` variable.

```typescript
async function deregisterEventHandlers() {
  // Deregisters event handlers.
  await Word.run(async (context) => {
    for (let i = 0; i < eventContexts.length; i++) {
      await Word.run(eventContexts[i].context, async (context) => {
        eventContexts[i].remove();
      });
    }

    await context.sync();

    eventContexts = [];
    console.log("Removed event handlers.");
  });
}
```

## See also

- [Annotation APIs need a Microsoft 365 subscription](word-add-ins-troubleshooting.md#annotations-dont-work)
- [Manage annotations code sample](https://github.com/OfficeDev/office-js-snippets/blob/prod/samples/word/50-document/manage-annotations.yaml)
- [Work with events using the Word JavaScript API](word-add-ins-events.md)
