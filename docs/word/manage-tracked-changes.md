---
title: Work with tracked changes in Word
description: Build review workflows that inspect, filter, accept, and reject tracked changes in Word documents.
ms.date: 09/22/2026
ms.localizationpriority: medium
ms.topic: how-to
---

# Work with tracked changes in Word

The Word JavaScript API lets an Office Add-in inspect and process tracked changes in a document. Use these APIs to automate repetitive review tasks and build custom review experiences for contracts, financial documents, legal documents, and other content that goes through multiple revisions. By using these APIs, your add-in can:

- Create a custom list of revisions.
- Filter changes by author, date, type, or location in the document.
- Accept or reject individual changes that match your business rules.
- Accept or reject every change in a selected scope.
- Process the tracked changes produced by a document comparison.

## Turn Track Changes on or off

Use the [Document.changeTrackingMode](/javascript/api/word/word.document#word-word-document-changetrackingmode-member) property to control how Word tracks subsequent edits.

```typescript
async function setChangeTrackingMode(
  mode: Word.ChangeTrackingMode
): Promise<void> {
  await Word.run(async (context) => {
    context.document.changeTrackingMode = mode;
    await context.sync();
  });
}

// Choose the mode that your workflow requires.
// await setChangeTrackingMode(Word.ChangeTrackingMode.trackAll);
// await setChangeTrackingMode(Word.ChangeTrackingMode.trackMineOnly);
// await setChangeTrackingMode(Word.ChangeTrackingMode.off);
```

Setting the mode to `off` doesn't accept, reject, or remove existing tracked changes. It only stops Word from tracking subsequent edits. To process existing revisions, use the tracked-change APIs described in the following sections.

## Get tracked changes from a document location

Call `getTrackedChanges()` on any of the following objects.

| Scope | Method |
| :----- | :----- |
| Document body | [Body.getTrackedChanges](/javascript/api/word/word.body#word-word-body-gettrackedchanges-member(1)) |
| Content control | [ContentControl.getTrackedChanges](/javascript/api/word/word.contentcontrol#word-word-contentcontrol-gettrackedchanges-member(1)) |
| Paragraph | [Paragraph.getTrackedChanges](/javascript/api/word/word.paragraph#word-word-paragraph-gettrackedchanges-member(1)) |
| Range | [Range.getTrackedChanges](/javascript/api/word/word.range#word-word-range-gettrackedchanges-member(1)) |

The method returns a [Word.TrackedChangeCollection](/javascript/api/word/word.trackedchangecollection) for the current version of that scope. To get changes for the main document, use `context.document.body.getTrackedChanges()`. The `Document` object itself doesn't have a `getTrackedChanges()` method.

Each [Word.TrackedChange](/javascript/api/word/word.trackedchange) object provides the following metadata.

| Property | Description |
| :----- | :----- |
| `author` | The name of the person who made the change. |
| `date` | The date and time when the change was created. |
| `text` | The text associated with the change. |
| `type` | The change type: `Added`, `Deleted`, or `Formatted` |

Call `getRange()` on a tracked change to get its location in the document. This information is useful when your review rules depend on where a revision appears, such as inside a contract clause or a designated content control.

## Inspect, filter, and process tracked changes

The following example retrieves tracked changes from the main document body and applies a customer-defined rule. The rule rejects formatting changes made by a specified reviewer after a cutoff date and leaves all other changes unchanged.

```typescript
async function rejectMatchingChanges(
  reviewerName: string,
  cutoffDate: Date
): Promise<void> {

  await Word.run(async (context) => {
    const changes = context.document.body.getTrackedChanges();
    changes.load({
      author: true,
      date: true,
      text: true,
      type: true
    });

    await context.sync();

    const matchingChanges = changes.items.filter((change) => {
      console.log({
        author: change.author,
        date: change.date,
        text: change.text,
        type: change.type
      });

      return (
        change.author === reviewerName &&
        change.date > cutoffDate &&
        change.type === Word.TrackedChangeType.formatted
      );
    });

    matchingChanges.forEach((change) => change.reject());
    await context.sync();

    console.log(`Rejected ${matchingChanges.length} matching changes.`);
  });
}
```

To accept the matching revisions instead, replace `change.reject()` with `change.accept()`.

Use this pattern to create rules such as:

- Accept all additions from an approved reviewer.
- Reject deletions made before or after a specified date.
- Process only revisions within a selected range, paragraph, or content control.
- Present revision metadata in a task pane and let the user choose which changes to accept or reject.

The collection also provides `acceptAll()` and `rejectAll()` methods. Use these methods when every change in the scope should receive the same action.

## Handle empty collections and navigation errors

Loading the `items` property of a collection is the simplest way to process zero or more changes. An empty scope returns an empty `items` array.

If you navigate the collection one object at a time, consider the following error behavior.

- `TrackedChangeCollection.getFirst()` throws an `ItemNotFound` error when the collection is empty.
- `TrackedChange.getNext()` throws an `ItemNotFound` error when the object is the last change.
- `getFirstOrNullObject()` and `getNextOrNullObject()` return an object whose `isNullObject` property is `true` instead of throwing.

## Compare documents and process the differences

Use the `Document.compare` or `Document.compareFromBase64` methods to generate document comparison differences as tracked changes. After you place comparison results in the current document, your add-in can call `getTrackedChanges()` and use the same metadata, filtering, and accept-or-reject workflow described earlier.

| API | Requirement set | Return value |
| :----- | :----- | :----- |
| `Document.compare(filePath, options)` | WordApiDesktop 1.1 | `void` |
| `Document.compareFromBase64(base64File, options)` | WordApiDesktop 1.2 | `void` |

The APIs that initiate a comparison are currently in platform-specific `WordApiDesktop` requirement sets. These requirement sets are production APIs in Word on Windows and Word on Mac, but not available in Word on the web. Your add-in must check for the applicable `WordApiDesktop` requirement set at runtime and provide an alternate workflow for web users.

The following example compares the current document with a Base64-encoded document and places the differences in the current document. It then retrieves the resulting tracked changes.

```typescript
async function compareAndListChanges(base64File: string): Promise<void> {
  if (!Office.context.requirements.isSetSupported("WordApiDesktop", "1.2")) {
    throw new Error("Document comparison isn't supported by this version of Word.");
  }

  await Word.run(async (context) => {
    const options: Word.DocumentCompareOptions = {
      compareTarget: Word.CompareTarget.compareTargetCurrent,
      detectFormatChanges: true,
      removeDateAndTime: false,
      removePersonalInformation: false
    };

    context.document.compareFromBase64(base64File, options);
    await context.sync();

    const changes = context.document.body.getTrackedChanges();
    changes.load({
      author: true,
      date: true,
      text: true,
      type: true
    });
    await context.sync();

    changes.items.forEach((change) => {
      console.log(change.author, change.date, change.type, change.text);
    });
  });
}
```

Use `DocumentCompareOptions` to specify the author name assigned to comparison differences, whether to detect formatting changes, where to display the results, and whether to remove personal information or timestamps. For the complete option list, see [Word.DocumentCompareOptions](/javascript/api/word/word.documentcompareoptions).

> [!NOTE]
> When you use `compareFromBase64()`, the `compareTarget` option can't be `compareTargetSelected`.

## Review workflow ideas

By using tracked-change metadata and object-level accept or reject operations, Office Add-ins offer several solutions for several scenarios. They can:

- Show a custom revision dashboard grouped by reviewer, date, or change type.
- Enforce contract review policies by accepting or rejecting only changes in specific clauses.
- Route financial-document revisions through an approval workflow based on author and creation time.
- Let users review changes in a selected range without processing unrelated parts of the document.
- Apply the same business rules to manually tracked revisions and desktop-generated document comparison results.

## See also

- [Word.TrackedChange class](/javascript/api/word/word.trackedchange)
- [Word.TrackedChangeCollection class](/javascript/api/word/word.trackedchangecollection)
- [Word.ChangeTrackingMode enum](/javascript/api/word/word.changetrackingmode)
- [Understanding platform-specific requirement sets](../develop/platform-specific-requirement-sets.md)
