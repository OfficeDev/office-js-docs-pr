---
title: Work with comments using the Excel JavaScript API
description: 'Information on using the APIs to add, remove, and edit comments and comment threads.'
ms.date: 02/11/2020
localization_priority: Normal
---

# Work with comments using the Excel JavaScript API

This article describes how to add, read, modify, and remove comments in a workbook with the Excel JavaScript API. You can learn more about the comment feature from the [Insert comments and notes in Excel](https://support.office.com/article/insert-comments-and-notes-in-excel-bdcc9f5d-38e2-45b4-9a92-0b2b5c7bf6f8) article.

In the Excel JavaScript API, a comment is both the initial note and the connected threaded discussion. It is tied to an individual cell. Anyone viewing the workbook with sufficient permissions can reply to a comment. A [Comment](/javascript/api/excel/excel.comment) object stores those replies as [CommentReply](/javascript/api/excel/excel.commentreply) objects. You should consider a comment to be a thread and that a thread must have a special entry as the starting point.

![An Excel comment, labelled "Comment" with two replies, labelled "Comment.replies[0]" and "Comment.replies[1].](../images/excel-comments.png)

Comments within a workbook are tracked by the `Workbook.comments` property. This includes comments created by users and also comments created by your add-in. The `Workbook.comments` property is a [CommentCollection](/javascript/api/excel/excel.commentcollection) object that contains a collection of [Comment](/javascript/api/excel/excel.comment) objects. Comments are also accessible at the [Worksheet](/javascript/api/excel/excel.worksheet) level. The samples in this article work with comments at the workbook level, but they can be easily modified to use the `Worksheet.comments` property.

## Add comments

Use the `CommentCollection.add` method to add comments to a workbook. This method takes up to three parameters:

- `cellAddress`: The cell where the comment is added. This can either be a string or [Range](/javascript/api/excel/excel.range) object. The range must be a single cell.
- `content`: The comment's content. Use a string for plain text comments. Use a [CommentRichContent](/javascript/api/excel/excel.commentrichcontent) object for comments with [mentions](#mentions-preview).
- `contentType`: A [ContentType](/javascript/api/excel/excel.contenttype) enum specifying type of content. The default value is `ContentType.plain`.

The following code sample adds a comment to cell **A2**.

```js
Excel.run(function (context) {
    // Add a comment to A2 on the "MyWorksheet" worksheet.
    var comments = context.workbook.comments;

    // Note that an InvalidArgument error will be thrown if multiple cells passed to `Comment.add`.
    comments.add("MyWorksheet!A2", "TODO: add data.");
    return context.sync();
});
```

> [!NOTE]
> Comments added by an add-in are attributed to the current user of that add-in.

### Add comment replies

A `Comment` object is a comment thread that contains zero or more replies. `Comment` objects have a `replies` property, which is a [CommentReplyCollection](/javascript/api/excel/excel.commentreplycollection) that contains [CommentReply](/javascript/api/excel/excel.commentreply) objects. To add a reply to a comment, use the `CommentReplyCollection.add` method, passing in the text of the reply. Replies are displayed in the order they are added. They are also attributed to the current user of the add-in.

The following code sample adds a reply to the first comment in the workbook.

```js
Excel.run(function (context) {
    // Get the first comment added to the workbook.
    var comment = context.workbook.comments.getItemAt(0);
    comment.replies.add("Thanks for the reminder!");
    return context.sync();
});
```

## Edit comments

To edit a comment or comment reply, set its `Comment.content` property or `CommentReply.content` property.

```js
Excel.run(function (context) {
    // Edit the first comment in the workbook.
    var comment = context.workbook.comments.getItemAt(0);
    comment.content = "PLEASE add headers here.";
    return context.sync();
});
```

### Edit comment replies

To edit a comment reply, set its `CommentReply.content` property.

```js
Excel.run(function (context) {
    // Edit the first comment reply on the first comment in the workbook.
    var comment = context.workbook.comments.getItemAt(0);
    var reply = comment.replies.getItemAt(0);
    reply.content = "Never mind";
    return context.sync();
});
```

## Delete comments

To delete a comment use the `Comment.delete` method. Deleting a comment also deletes the replies associated with that comment.

```js
Excel.run(function (context) {
    // Delete the comment thread at A2 on the "MyWorksheet" worksheet.
    context.workbook.comments.getItemByCell("MyWorksheet!A2").delete();
    return context.sync();
});
```

### Delete comment replies

To delete a comment reply, use the `CommentReply.delete` method.

```js
Excel.run(function (context) {
    // Delete the first comment reply from this worksheet's first comment.
    var comment = context.workbook.comments.getItemAt(0);
    comment.replies.getItemAt(0).delete();
    return context.sync();
});
```

## Resolve comment threads ([preview](../reference/requirement-sets/excel-preview-apis.md))

A comment thread has a configurable boolean value, `resolved`, to indicate if it is resolved. A value of `true` means the comment thread is resolved. A value of `false` means the comment thread is either new or reopened.

```js
Excel.run(function (context) {
    // Resolve the first comment thread in the workbook.
    context.workbook.comments.getItemAt(0).resolved = true;
    return context.sync();
});
```

Comment replies have a readonly `resolved` property. Its value is always equal to that of the rest of the thread.

## Comment metadata

Each comment contains metadata about its creation, such as the author and creation date. Comments created by your add-in are considered to be authored by the current user.

The following sample shows how to display the author's email, author's name, and creation date of a comment at **A2**.

```js
Excel.run(function (context) {
    // Get the comment at cell A2 in the "MyWorksheet" worksheet.
    var comment = context.workbook.comments.getItemByCell("MyWorksheet!A2");

    // Load and print the following values.
    comment.load(["authorEmail", "authorName", "creationDate"]);
    return context.sync().then(function () {
        console.log(`${comment.creationDate.toDateString()}: ${comment.authorName} (${comment.authorEmail})`);
    });
});
```

### Comment reply metadata

Comment replies store the same types of metadata as the initial comment.

The following sample shows how to display the author's email, author's name, and creation date of the latest comment reply at **A2**.

```js
Excel.run(function (context) {
    // Get the comment at cell A2 in the "MyWorksheet" worksheet.
    var comment = context.workbook.comments.getItemByCell("MyWorksheet!A2");
    var replyCount = comment.replies.getCount();
    // Sync to get the current number of comment replies.
    return context.sync().then(function () {
        // Get the last comment reply in the comment thread.
        var reply = comment.replies.getItemAt(replyCount.value - 1);
        reply.load(["authorEmail", "authorName", "creationDate"]);
        // Sync to load the reply metadata to print.
        return context.sync().then(function () {
            console.log(`Latest reply: ${reply.creationDate.toDateString()}: ${reply.authorName} ${reply.authorEmail})`);
            return context.sync();
        });
    });
});
```

## Mentions ([online-only](../reference/requirement-sets/excel-api-online-requirement-set.md))

> [!NOTE]
> The comment mention APIs are currently available only in public preview. [!INCLUDE [Information about using preview APIs](../includes/using-excel-preview-apis.md)]

> [!IMPORTANT]
> Comment mentions are currently only supported for Excel on the web.

[Mentions](https://support.office.com/article/use-mention-in-comments-to-tag-someone-for-feedback-644bf689-31a0-4977-a4fb-afe01820c1fd) are used to tag colleagues in a comment. This sends them notifications with your comment's content. Your add-in can create these mentions on your behalf.

Comments with mentions need to be created with [CommentRichContent](/javascript/api/excel/excel.commentrichcontent) objects. Call `CommentCollection.add` with a `CommentRichContent` containing one or more mentions and specify `ContentType.mention` as the `contentType` parameter. The `content` string also needs to be formatted to insert the mention into the text. The format for a mention is: `<at id="{replyIndex}">{mentionName}</at>`.

> [NOTE]
> Currently, only the mention's exact name can be used as the text of the mention link. Support for shortened versions of a name will be added later.

The following example shows a comment with a single mention.

```js
Excel.run(function (context) {
    // Add an "@mention" for "Kate Kristensen" to cell A1 in the "MyWorksheet" worksheet.
    var mention = {
        email: "kakri@contoso.com",
        id: 0,
        name: "Kate Kristensen"
    };

    // This will tag the mention's name using the '@' syntax.
    // They will be notified via email.
    var commentBody = {
        mentions: [mention],
        richContent: '<at id="0">' + mention.name + "</at> -  Can you take a look?"
    };

    // Note that an InvalidArgument error will be thrown if multiple cells passed to `comment.add`.
    context.workbook.comments.add("MyWorksheet!A1", commentBody, Excel.ContentType.mention);
    return context.sync();
});
```

## See also

- [Fundamental programming concepts with the Excel JavaScript API](excel-add-ins-core-concepts.md)
- [Work with workbooks using the Excel JavaScript API](excel-add-ins-workbooks.md)
- [Insert comments and notes in Excel](https://support.office.com/article/insert-comments-and-notes-in-excel-bdcc9f5d-38e2-45b4-9a92-0b2b5c7bf6f8)
