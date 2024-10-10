---
title: Troubleshoot Word add-ins
description: Learn how to troubleshoot development errors in Word add-ins.
ms.date: 10/10/2024
ms.topic: troubleshooting
ms.localizationpriority: medium
---

# Troubleshoot Word add-ins

This article discusses troubleshooting issues that are unique to Word. Use the feedback tool at the end of the page to suggest other issues that can be added to this article.

## All selected ranges aren't recognized

If non-contiguous or discontinuous selections are made, the Word API only operates on the last contiguous range in the selection. An unexpected case of this is when you select a column in a table then call, for example, [Document.getSelection](/javascript/api/word/word.document?view=word-js-preview#word-word-document-getselection-member(1)), only the final cell in the selection is returned by the API. Although the selection of a column seems continuous, the API recognizes it as a non-contiguous selection (e.g., a cell per row).

To learn more generally about making non-contiguous selections, see [How to select items that are not next to each other](https://support.microsoft.com/topic/8b9c1be9-cca3-935a-7cbf-94403aa48d2e).

## Body.insertFileFromBase64 doesn't insert header or footer

It's by design that the [Body.insertFileFromBase64](/javascript/api/word/word.body#word-word-body-insertfilefrombase64-member(1)) method excludes any header or footer that was in the source file.

To include any headers or footers from the source file, use [Document.insertFileFromBase64](/javascript/api/word/word.document#word-word-document-insertfilefrombase64-member(1)) instead.

## Get a GeneralException when working with styles

If users are hitting a GeneralException when your add-in calls [Document.insertFileFromBase64](/javascript/api/word/word.document#word-word-document-insertfilefrombase64-member(1)) or Style APIs, it may be that those users are exceeding limits imposed by the Word application. To learn more about these limits, see [Operating parameter limitations and specifications in Word](/office/troubleshoot/word/operating-parameter-limitation).

## Layout breaks when using `insertHtml` while cursor is in content control in header

This issue may occur when the following three conditions are met.

1. Have at least one content control in the header and at least one in the footer of the Word document.
1. Ensure the cursor is inside a content control in the header.
1. Call [insertHtml](/javascript/api/word/word.contentcontrol#word-word-contentcontrol-inserthtml-member(1)) to set a content control in the footer.

The footer is then unexpectedly mixed with the header. To avoid this, clear the content control in the footer before setting it, as shown in the following code sample.

```TypeScript
await Word.run(async (context) => {
    // Credit to https://github.com/barisbikmaz for this version of the workaround.
    // For more information, see https://github.com/OfficeDev/office-js/issues/129.

    // Let's say there are 2 content controls in the header and 1 in the footer.
    const contentControls = context.document.contentControls;
    contentControls.load();

    await context.sync().then(function () {
        // Clear the 2 content controls in the header.
        contentControls.items[0].clear(); 
        contentControls.items[1].clear();

        // Clear the control control in the footer then update it.
        contentControls.items[2].clear();
        contentControls.items[2].insertHtml('<p>New Footer</p>', 'Replace');
    });
});
```

## Lost formatting of last bullet in a list or last paragraph

If the formatting of the last bullet in a list or the last paragraph is lost in the specified body or range, check if you're using [Body.insertFileFromBase64](/javascript/api/word/word.body#word-word-body-insertfilefrombase64-member(1)) or [Range.insertFileFromBase64](/javascript/api/word/word.range#word-word-range-insertfilefrombase64-member(1)). If so, update your code to use [Document.insertFileFromBase64](/javascript/api/word/word.document#word-word-document-insertfilefrombase64-member(1)) instead.

## Meaning of null property values in the response

`null` has special implications in the Word JavaScript APIs. It's used to represent default values or no formatting.

Formatting properties such as [color](/javascript/api/word/word.font#word-word-font-color-member) will contain `null` values in the response when different values exist in the specified [range](/javascript/api/word/word.range). For example, if you retrieve a range and load its `range.font.color` property:

- If all text in the range has the same font color, `range.font.color` specifies that color.
- If multiple font colors are present within the range, `range.font.color` is `null`.

## See also

- [Troubleshoot development errors with Office Add-ins](../testing/troubleshoot-development-errors.md)
- [Troubleshoot user errors with Office Add-ins](../testing/testing-and-troubleshooting.md)
