---
title: Work with comments using the Excel JavaScript API
description: Information on using the APIs to add, remove, and edit comments and comment threads.
ms.date: 04/07/2025
ms.localizationpriority: medium
---

# Work with comments using the Excel JavaScript API

This article describes how to add, read, modify, and remove comments in a workbook with the Excel JavaScript API. You can learn more about the comment feature from the [Insert comments and notes in Excel](https://support.microsoft.com/office/bdcc9f5d-38e2-45b4-9a92-0b2b5c7bf6f8) article.

In the Excel JavaScript API, a comment includes both the single initial comment and the connected threaded discussion. It is tied to an individual cell. Anyone viewing the workbook with sufficient permissions can reply to a comment. A [Comment](/javascript/api/excel/excel.comment) object stores those replies as [CommentReply](/javascript/api/excel/excel.commentreply) objects. You should consider a comment to be a thread and that a thread must have a special entry as the starting point.

:::image type="content" source="../images/excel-comments.png" alt-text="An Excel comment, labelled 'Comment' with two replies, labelled 'Comment.replies[0]' and 'Comment.replies[1]'.":::

Comments within a workbook are tracked by the `Workbook.comments` property. This includes comments created by users and also comments created by your add-in. The `Workbook.comments` property is a [CommentCollection](/javascript/api/excel/excel.commentcollection) object that contains a collection of [Comment](/javascript/api/excel/excel.comment) objects. Comments are also accessible at the [Worksheet](/javascript/api/excel/excel.worksheet) level. The samples in this article work with comments at the workbook level, but they can be easily modified to use the `Worksheet.comments` property.

> [!TIP]
> To learn about adding and editing notes with the Excel JavaScript API, see [Work with notes using the Excel JavaScript API](excel-add-ins-notes.md).

## Add comments

Use the `CommentCollection.add` method to add comments to a workbook. This method takes up to three parameters:

- `cellAddress`: The cell where the comment is added. This can either be a string or [Range](/javascript/api/excel/excel.range) object. The range must be a single cell.
- `content`: The comment's content. Use a string for plain text comments. Use a [CommentRichContent](/javascript/api/excel/excel.commentrichcontent) object for comments with [mentions](#mentions).
- `contentType`: A [ContentType](/javascript/api/excel/excel.contenttype) enum specifying type of content. The default value is `ContentType.plain`.

The following code sample adds a comment to cell **A2**.

```js
await Excel.run(async (context) => {
    // Add a comment to A2 on the "MyWorksheet" worksheet.
    let comments = context.workbook.comments;

    // Note that an InvalidArgument error will be thrown if multiple cells passed to `Comment.add`.
    comments.add("MyWorksheet!A2", "TODO: add data.");
    await context.sync();
});
```

> [!NOTE]
> Comments added by an add-in are attributed to the current user of that add-in.

### Add comment replies

A `Comment` object is a comment thread that contains zero or more replies. `Comment` objects have a `replies` property, which is a [CommentReplyCollection](/javascript/api/excel/excel.commentreplycollection) that contains [CommentReply](/javascript/api/excel/excel.commentreply) objects. To add a reply to a comment, use the `CommentReplyCollection.add` method, passing in the text of the reply. Replies are displayed in the order they are added. They are also attributed to the current user of the add-in.

The following code sample adds a reply to the first comment in the workbook.

```js
await Excel.run(async (context) => {
    // Get the first comment added to the workbook.
    let comment = context.workbook.comments.getItemAt(0);
    comment.replies.add("Thanks for the reminder!");
    await context.sync();
});
```

## Edit comments

To edit a comment or comment reply, set its `Comment.content` property or `CommentReply.content` property.

```js
await Excel.run(async (context) => {
    // Edit the first comment in the workbook.
    let comment = context.workbook.comments.getItemAt(0);
    comment.content = "PLEASE add headers here.";
    await context.sync();
});
```

### Edit comment replies

To edit a comment reply, set its `CommentReply.content` property.

```js
await Excel.run(async (context) => {
    // Edit the first comment reply on the first comment in the workbook.
    let comment = context.workbook.comments.getItemAt(0);
    let reply = comment.replies.getItemAt(0);
    reply.content = "Never mind";
    await context.sync();
});
```

## Delete comments

To delete a comment use the `Comment.delete` method. Deleting a comment also deletes the replies associated with that comment.

```js
await Excel.run(async (context) => {
    // Delete the comment thread at A2 on the "MyWorksheet" worksheet.
    context.workbook.comments.getItemByCell("MyWorksheet!A2").delete();
    await context.sync();
});
```

### Delete comment replies

To delete a comment reply, use the `CommentReply.delete` method.

```js
await Excel.run(async (context) => {
    // Delete the first comment reply from this worksheet's first comment.
    let comment = context.workbook.comments.getItemAt(0);
    comment.replies.getItemAt(0).delete();
    await context.sync();
});
```

## Resolve comment threads

A comment thread has a configurable boolean value, `resolved`, to indicate if it is resolved. A value of `true` means the comment thread is resolved. A value of `false` means the comment thread is either new or reopened.

```js
await Excel.run(async (context) => {
    // Resolve the first comment thread in the workbook.
    context.workbook.comments.getItemAt(0).resolved = true;
    await context.sync();
});
```

Comment replies have a read-only `resolved` property. Its value is always equal to that of the rest of the thread.

## Comment metadata

Each comment contains metadata about its creation, such as the author and creation date. Comments created by your add-in are considered to be authored by the current user.

The following sample shows how to display the author's email, author's name, and creation date of a comment at **A2**.

```js
await Excel.run(async (context) => {
    // Get the comment at cell A2 in the "MyWorksheet" worksheet.
    let comment = context.workbook.comments.getItemByCell("MyWorksheet!A2");

    // Load and print the following values.
    comment.load(["authorEmail", "authorName", "creationDate"]);
    await context.sync();
    
    console.log(`${comment.creationDate.toDateString()}: ${comment.authorName} (${comment.authorEmail})`);
});
```

### Comment reply metadata

Comment replies store the same types of metadata as the initial comment.

The following sample shows how to display the author's email, author's name, and creation date of the latest comment reply at **A2**.

```js
await Excel.run(async (context) => {
    // Get the comment at cell A2 in the "MyWorksheet" worksheet.
    let comment = context.workbook.comments.getItemByCell("MyWorksheet!A2");
    let replyCount = comment.replies.getCount();
    // Sync to get the current number of comment replies.
    await context.sync();

    // Get the last comment reply in the comment thread.
    let reply = comment.replies.getItemAt(replyCount.value - 1);
    reply.load(["authorEmail", "authorName", "creationDate"]);

    // Sync to load the reply metadata to print.
    await context.sync();

    console.log(`Latest reply: ${reply.creationDate.toDateString()}: ${reply.authorName} ${reply.authorEmail})`);
    await context.sync();
});
```

## Mentions

[Mentions](https://support.microsoft.com/office/644bf689-31a0-4977-a4fb-afe01820c1fd) are used to tag colleagues in a comment. This sends them notifications with your comment's content. Your add-in can create these mentions on your behalf.

Comments with mentions need to be created with [CommentRichContent](/javascript/api/excel/excel.commentrichcontent) objects. Call `CommentCollection.add` with a `CommentRichContent` containing one or more mentions and specify `ContentType.mention` as the `contentType` parameter. The `content` string also needs to be formatted to insert the mention into the text. The format for a mention is: `<at id="{replyIndex}">{mentionName}</at>`.

> [!NOTE]
> Currently, only the mention's exact name can be used as the text of the mention link. Support for shortened versions of a name will be added later.

The following example shows a comment with a single mention.

```js
await Excel.run(async (context) => {
    // Add an "@mention" for "Kate Kristensen" to cell A1 in the "MyWorksheet" worksheet.
    let mention = {
        email: "kakri@contoso.com",
        id: 0,
        name: "Kate Kristensen"
    };

    // This will tag the mention's name using the '@' syntax.
    // They will be notified via email.
    let commentBody = {
        mentions: [mention],
        richContent: '<at id="0">' + mention.name + "</at> -  Can you take a look?"
    };

    // Note that an InvalidArgument error will be thrown if multiple cells passed to `comment.add`.
    context.workbook.comments.add("MyWorksheet!A1", commentBody, Excel.ContentType.mention);
    await context.sync();
});
```

## Comment events

Your add-in can listen for comment additions, changes, and deletions. [Comment events](/javascript/api/excel/excel.commentcollection#event-details) occur on the `CommentCollection` object. To listen for comment events, register the `onAdded`, `onChanged`, or `onDeleted` comment event handler. When a comment event is detected, use this event handler to retrieve data about the added, changed, or deleted comment. The `onChanged` event also handles comment reply additions, changes, and deletions.

Each comment event only triggers once when multiple additions, changes, or deletions are performed at the same time. All the [CommentAddedEventArgs](/javascript/api/excel/excel.commentaddedeventargs), [CommentChangedEventArgs](/javascript/api/excel/excel.commentchangedeventargs), and [CommentDeletedEventArgs](/javascript/api/excel/excel.commentdeletedeventargs) objects contain arrays of comment IDs to map the event actions back to the comment collections.

See the [Work with Events using the Excel JavaScript API](excel-add-ins-events.md) article for additional information about registering event handlers, handling events, and removing event handlers.

### Comment addition events

The `onAdded` event is triggered when one or more new comments are added to the comment collection. This event is *not* triggered when replies are added to a comment thread (see [Comment change events](#comment-change-events) to learn about comment reply events).

The following sample shows how to register the `onAdded` event handler and then use the `CommentAddedEventArgs` object to retrieve the `commentDetails` array of the added comment.

> [!NOTE]
> This sample only works when a single comment is added.

```js
await Excel.run(async (context) => {
    let comments = context.workbook.worksheets.getActiveWorksheet().comments;

    // Register the onAdded comment event handler.
    comments.onAdded.add(commentAdded);

    await context.sync();
});

async function commentAdded() {
    await Excel.run(async (context) => {
        // Retrieve the added comment using the comment ID.
        // Note: This method assumes only a single comment is added at a time. 
        let addedComment = context.workbook.comments.getItem(event.commentDetails[0].commentId);

        // Load the added comment's data.
        addedComment.load(["content", "authorName"]);

        await context.sync();

        // Print out the added comment's data.
        console.log(`A comment was added. ID: ${event.commentDetails[0].commentId}. Comment content:${addedComment.content}. Comment author:${addedComment.authorName}`);
        await context.sync();
    });
}
```

### Comment change events

The `onChanged` comment event is triggered in the following scenarios.

- A comment's content is updated.
- A comment thread is resolved.
- A comment thread is reopened.
- A reply is added to a comment thread.
- A reply is updated in a comment thread.
- A reply is deleted in a comment thread.

The following sample shows how to register the `onChanged` event handler and then use the `CommentChangedEventArgs` object to retrieve the `commentDetails` array of the changed comment.

> [!NOTE]
> This sample only works when a single comment is changed.

```js
await Excel.run(async (context) => {
    let comments = context.workbook.worksheets.getActiveWorksheet().comments;

    // Register the onChanged comment event handler.
    comments.onChanged.add(commentChanged);

    await context.sync();
});

async function commentChanged() {
    await Excel.run(async (context) => {
        // Retrieve the changed comment using the comment ID.
        // Note: This method assumes only a single comment is changed at a time. 
        let changedComment = context.workbook.comments.getItem(event.commentDetails[0].commentId);

        // Load the changed comment's data.
        changedComment.load(["content", "authorName"]);

        await context.sync();

        // Print out the changed comment's data.
        console.log(`A comment was changed. ID: ${event.commentDetails[0].commentId}. Updated comment content: ${changedComment.content}. Comment author: ${changedComment.authorName}`);
        await context.sync();
    });
}
```

### Comment deletion events

The `onDeleted` event is triggered when a comment is deleted from the comment collection. Once a comment has been deleted, its metadata is no longer available. The [CommentDeletedEventArgs](/javascript/api/excel/excel.commentdeletedeventargs) object provides comment IDs, in case your add-in is managing individual comments.

The following sample shows how to register the `onDeleted` event handler and then use the `CommentDeletedEventArgs` object to retrieve the `commentDetails` array of the deleted comment.

> [!NOTE]
> This sample only works when a single comment is deleted.

```js
await Excel.run(async (context) => {
    let comments = context.workbook.worksheets.getActiveWorksheet().comments;

    // Register the onDeleted comment event handler.
    comments.onDeleted.add(commentDeleted);

    await context.sync();
});

async function commentDeleted() {
    await Excel.run(async (context) => {
        // Print out the deleted comment's ID.
        // Note: This method assumes only a single comment is deleted at a time. 
        console.log(`A comment was deleted. ID: ${event.commentDetails[0].commentId}`);
    });
}
```

## See also

- [Excel JavaScript object model in Office Add-ins](excel-add-ins-core-concepts.md)
- [Work with workbooks using the Excel JavaScript API](excel-add-ins-workbooks.md)
- [Work with Events using the Excel JavaScript API](excel-add-ins-events.md)
- [Insert comments and notes in Excel](https://support.microsoft.com/office/bdcc9f5d-38e2-45b4-9a92-0b2b5c7bf6f8)
