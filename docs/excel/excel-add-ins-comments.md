---
title: Manage comments in Excel add-ins with Excel JavaScript API
description: Learn how to add, reply to, edit, resolve, and monitor threaded comments in Excel workbooks by using the Excel JavaScript API.
ms.date: 06/03/2026
ms.topic: how-to
ms.localizationpriority: medium
ai-usage: ai-assisted
---

# Manage comments in Excel add-ins by using the Excel JavaScript API

This article shows how to create threaded comments on specific cells, reply to them, edit or delete them, resolve conversations, read author metadata, add mentions, and respond to comment events by using the Excel JavaScript API.

In the Excel JavaScript API, a comment is a thread that starts with one comment and can include replies. Each thread is tied to a single cell. If you need legacy note behavior instead of threaded discussions, see [Work with notes using the Excel JavaScript API](excel-add-ins-notes.md).

## What you can do with comments

Use the Excel comment APIs to:

- Add a new comment thread to a cell.
- Add, edit, and delete replies in an existing thread.
- Resolve or reopen a thread.
- Read author and creation metadata.
- Create comments that include mentions.
- Listen for comment add, change, and delete events.

## Understand the comment object model

The `Workbook.comments` property tracks comments in a workbook. This property returns a [CommentCollection](/javascript/api/excel/excel.commentcollection) that contains both user-created comments and comments created by your add-in. You can also access comments at the [Worksheet](/javascript/api/excel/excel.worksheet) level through the `Worksheet.comments` property.

A [Comment](/javascript/api/excel/excel.comment) object represents the full thread for a single cell. Replies in that thread are stored as [CommentReply](/javascript/api/excel/excel.commentreply) objects in the comment's `replies` collection.

:::image type="content" source="../images/excel-comments.png" alt-text="An Excel comment, labeled 'Comment', with two replies, labeled 'Comment.replies[0]' and 'Comment.replies[1]'.":::

## Add comment threads

Use `CommentCollection.add` to start a threaded conversation on a cell. The method accepts up to three parameters:

- `cellAddress`: The cell where you add the comment. This parameter can be a string or a [Range](/javascript/api/excel/excel.range) object. The range must be a single cell.
- `content`: The comment text. Use a string for plain text comments. Use a [CommentRichContent](/javascript/api/excel/excel.commentrichcontent) object for comments that include [mentions](#mention-users).
- `contentType`: A [ContentType](/javascript/api/excel/excel.contenttype) value that specifies the content type. The default is `ContentType.plain`.

The following sample starts a review thread on cell **A2**. Note that comments that your add-in creates are attributed to the current user.

```js
await Excel.run(async (context) => {
    const comments = context.workbook.comments;
    comments.add("MyWorksheet!A2", "Please confirm the Q2 revenue total.");
    await context.sync();
});
```

> [!NOTE]
> An `InvalidArgument` error is thrown if the range contains multiple cells.

### Add replies to a comment thread

Use `CommentReplyCollection.add` when your add-in needs to continue an existing discussion. Replies are displayed in the order they're added and are also attributed to the current user.

The following sample adds a reply to the thread on cell **A2**.

```js
await Excel.run(async (context) => {
    const comment = context.workbook.comments.getItemByCell("MyWorksheet!A2");
    comment.replies.add("Thanks for the reminder!");
    await context.sync();
});
```

## Edit comment threads

Update `Comment.content` to change the first entry in a thread. Update `CommentReply.content` to change a specific reply.

The following sample updates the main comment on cell **A2**.

```js
await Excel.run(async (context) => {
    const comment = context.workbook.comments.getItemByCell("MyWorksheet!A2");
    comment.content = "Please confirm the Q2 revenue total before we publish this workbook.";
    await context.sync();
});
```

### Edit a reply

Use this pattern when your add-in needs to revise an earlier response in the thread.

```js
await Excel.run(async (context) => {
    const comment = context.workbook.comments.getItemByCell("MyWorksheet!A2");
    const reply = comment.replies.getItemAt(0);

    reply.content = "Thanks. I rechecked the total and it is correct.";
    await context.sync();
});
```

## Delete comment threads

Use `Comment.delete()` to remove an entire thread from a cell. Deleting a comment also deletes all replies in that thread.

```js
await Excel.run(async (context) => {
    const comment = context.workbook.comments.getItemByCell("MyWorksheet!A2");
    comment.delete();
    await context.sync();
});
```

### Delete a reply

Use `CommentReply.delete()` when you need to remove a single reply but keep the rest of the thread.

```js
await Excel.run(async (context) => {
    const comment = context.workbook.comments.getItemByCell("MyWorksheet!A2");
    const reply = comment.replies.getItemAt(0);

    reply.delete();
    await context.sync();
});
```

## Resolve and reopen comment threads

Use the `Comment.resolved` property to track whether a discussion still needs attention. Set the value to `true` to resolve the thread or to `false` to reopen it. `CommentReply.resolved` is read-only and always matches the state of the parent thread.

The following sample resolves the thread on cell **A2**.

```js
await Excel.run(async (context) => {
    const comment = context.workbook.comments.getItemByCell("MyWorksheet!A2");
    comment.resolved = true;
    await context.sync();
});
```

## Read comment metadata

Each comment stores metadata such as the author and creation date. Comments created by your add-in are authored by the current user.

The following sample logs the author email, author name, and creation date for the comment on cell **A2**.

```js
await Excel.run(async (context) => {
    const comment = context.workbook.comments.getItemByCell("MyWorksheet!A2");

    comment.load(["authorEmail", "authorName", "creationDate"]);
    await context.sync();

    console.log(`${comment.creationDate.toDateString()}: ${comment.authorName} (${comment.authorEmail})`);
});
```

### Read reply metadata

Replies store the same metadata as the initial comment. The following sample gets the latest reply in the thread on cell **A2** and logs its author information.

```js
await Excel.run(async (context) => {
    const comment = context.workbook.comments.getItemByCell("MyWorksheet!A2");
    const replyCount = comment.replies.getCount();
    await context.sync();

    if (replyCount.value === 0) {
        console.log("The thread has no replies.");
        return;
    }

    const reply = comment.replies.getItemAt(replyCount.value - 1);
    reply.load(["authorEmail", "authorName", "creationDate"]);
    await context.sync();

    console.log(`Latest reply: ${reply.creationDate.toDateString()}: ${reply.authorName} (${reply.authorEmail})`);
});
```

## Mention users

Use mentions when your add-in needs to tag a colleague in a comment and trigger an email notification. To create a comment with mentions, call `CommentCollection.add` with a [CommentRichContent](/javascript/api/excel/excel.commentrichcontent) object and set the `contentType` parameter to `ContentType.mention`.

Format each mention in the `richContent` string as `<at id="{replyIndex}">{mentionName}</at>`.

Currently, only the mention's exact name can be used as the text of the mention link. Support for shortened versions of a name will be added later.

The following sample adds a comment with a single mention to cell **A1**.

```js
await Excel.run(async (context) => {
    const mention = {
        email: "kakri@contoso.com",
        id: 0,
        name: "Kate Kristensen"
    };

    const commentBody = {
        mentions: [mention],
        richContent: `<at id="0">${mention.name}</at> Can you review the forecast?`
    };

    // An `InvalidArgument` error is thrown if the range contains multiple cells.
    context.workbook.comments.add("MyWorksheet!A1", commentBody, Excel.ContentType.mention);
    await context.sync();
});
```

## Handle comment events

Use comment events when your add-in needs to react to discussions as users update a workbook. [Comment events](/javascript/api/excel/excel.commentcollection#event-details) occur on the `CommentCollection` object.

Register handlers for:

- `onAdded` when a new comment thread is created.
- `onChanged` when a comment or reply is added, edited, deleted, resolved, or reopened.
- `onDeleted` when a comment thread is deleted.

If one operation affects multiple comments, the event arguments contain multiple items in `commentDetails`. The following samples use the first item only for clarity. For general event guidance, see [Work with Events using the Excel JavaScript API](excel-add-ins-events.md).

### Handle comment addition events

The `onAdded` event fires when one or more comments are added to the collection. It doesn't fire when a reply is added to an existing thread.

```js
await Excel.run(async (context) => {
    const comments = context.workbook.worksheets.getActiveWorksheet().comments;

    comments.onAdded.add(commentAdded);
    await context.sync();
});

async function commentAdded(event) {
    await Excel.run(async (context) => {
        const addedComment = context.workbook.comments.getItem(event.commentDetails[0].commentId);
        addedComment.load(["content", "authorName"]);
        await context.sync();

        console.log(`A comment was added. ID: ${event.commentDetails[0].commentId}. Content: ${addedComment.content}. Author: ${addedComment.authorName}`);
    });
}
```

### Handle comment change events

The `onChanged` event fires when:

- A comment's content is updated.
- A comment thread is resolved.
- A comment thread is reopened.
- A reply is added to a comment thread.
- A reply is updated in a comment thread.
- A reply is deleted from a comment thread.

```js
await Excel.run(async (context) => {
    const comments = context.workbook.worksheets.getActiveWorksheet().comments;

    comments.onChanged.add(commentChanged);
    await context.sync();
});

async function commentChanged(event) {
    await Excel.run(async (context) => {
        const changedComment = context.workbook.comments.getItem(event.commentDetails[0].commentId);
        changedComment.load(["content", "authorName"]);
        await context.sync();

        console.log(`A comment was changed. ID: ${event.commentDetails[0].commentId}. Content: ${changedComment.content}. Author: ${changedComment.authorName}`);
    });
}
```

### Handle comment deletion events

The `onDeleted` event fires when a comment is deleted from the collection. After a comment is deleted, its metadata is no longer available. Use the IDs in `CommentDeletedEventArgs.commentDetails` if your add-in needs to track deleted threads.

```js
await Excel.run(async (context) => {
    const comments = context.workbook.worksheets.getActiveWorksheet().comments;

    comments.onDeleted.add(commentDeleted);
    await context.sync();
});

async function commentDeleted(event) {
    console.log(`A comment was deleted. ID: ${event.commentDetails[0].commentId}`);
}
```

## See also

- [Core Excel object model concepts for Office Add-ins](excel-add-ins-core-concepts.md)
- [Manage Excel workbooks with the Excel JavaScript API](excel-add-ins-workbooks.md)
- [Work with Events using the Excel JavaScript API](excel-add-ins-events.md)
- [Work with notes using the Excel JavaScript API](excel-add-ins-notes.md)
- [Coauthoring in Excel add-ins](co-authoring-in-excel-add-ins.md)
- [Insert comments and notes in Excel](https://support.microsoft.com/office/bdcc9f5d-38e2-45b4-9a92-0b2b5c7bf6f8)
