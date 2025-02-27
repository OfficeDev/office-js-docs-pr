---
title: Troubleshoot Word add-ins
description: Learn how to troubleshoot development errors in Word add-ins.
ms.date: 02/25/2025
ms.topic: troubleshooting
ms.localizationpriority: medium
---

# Troubleshoot Word add-ins

This article discusses troubleshooting issues that are unique to Word. Use the feedback tool at the end of the page to suggest other issues that can be added to this article.

## All selected ranges aren't recognized

If noncontiguous selections are made, the Word API only operates on the last contiguous range in the selection. An unexpected case of this is when you select a column in a table then call, for example, [Document.getSelection](/javascript/api/word/word.document#word-word-document-getselection-member(1)), only the final cell in the selection is returned by the API. Although the selection of a column seems contiguous, the API recognizes it as a noncontiguous selection (e.g., a cell per row).

To learn more generally about making noncontiguous selections, see [How to select items that are not next to each other](https://support.microsoft.com/topic/8b9c1be9-cca3-935a-7cbf-94403aa48d2e).

## Annotations don't work

If the annotation APIs aren't working, it may be because you're not using a Microsoft 365 subscription. If you're using a one-time purchase license, this could be why these APIs aren't working for you.

The annotation APIs rely on a service that requires a Microsoft 365 subscription. Therefore, verify that you're running the add-in in Word connected to a Microsoft 365 subscription license before debugging further.

For more about this problem, see [GitHub issue 4953](https://github.com/OfficeDev/office-js/issues/4953).

## Body.insertFileFromBase64 doesn't insert header or footer

It's by design that the [Body.insertFileFromBase64](/javascript/api/word/word.body#word-word-body-insertfilefrombase64-member(1)) method excludes any header or footer that was in the source file.

To include any headers or footers from the source file, use [Document.insertFileFromBase64](/javascript/api/word/word.document#word-word-document-insertfilefrombase64-member(1)) instead.

## Can't use Mixed to set a property

Several enums in Word offer "Mixed" as a valid value. However, the value can primarily be returned when a get a property or make a get* API call. This is because "Mixed" means that several options are applied to the current selection. If you try to set the option to "Mixed", then it isn't clear which actual value should be applied to the selection.

For example, let's say you're working with the borders around a section of text. Each [border](/javascript/api/word/word.border#word-word-border-width-member) can be set to a different [width](/javascript/api/word/word.borderwidth). If the top border is "Pt025" (that is, 0.25 points), the bottom border is "None", and the left and right borders are "Pt050" (that is, 0.50 points), then when you get the width of the borders, "Mixed" is returned. If you want to change the width of the borders, call the set API on each border using an enum value other than `mixed`.

This behavior also applies for enum values like "Unknown".

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

## My add-in can no longer find the correct Word window

Microsoft Word, like other Windows applications, uses a hierarchy of windows to display documents and UI to users. These windows can be identified by window handles or class names. As of Office Version 2502 (Build 18526.20118), one of the windows in Word's hierarchy was removed.

It's possible that your Word add-in has a rigid dependency on Word's previous window hierarchy and so crashes or no longer works correctly. For an example issue, see [Possibly Microsoft 365 Office Apps updates crashing my Word Addin](https://aka.ms/word-wwf-crash-issue). We recommend that developers not rely on a particular window hierarchical structure. Instead, the current guidance is to search for a window's class name. To find the top-level Word window, search for the "OpusApp" class name. To find the window displaying an open Word document, search for the "_WwG" class name.

The following shows an example of the previous Word window hierarchy.

:::image type="content" source="../images/word-window-hierarchy-before.png" alt-text="Previous Word window hierarchy.":::

The following shows an example of the new window hierarchy. Note that the intermediate window with the "_WwF" class name is no longer present.

:::image type="content" source="../images/word-window-hierarchy-after.png" alt-text="New Word window hierarchy.":::

You can use a debugging tool like [Spy++](/visualstudio/debugger/using-spy-increment) to inspect an application's window hierarchy. However, keep in mind that the hierarchy could further change in the future.

## Native JavaScript APIs don't work with Word.Table

The [Word.Table](/javascript/api/word/word.table) object is different from an [HTML table object](https://developer.mozilla.org/docs/Learn_web_development/Core/Structuring_content/HTML_table_basics). The native JavaScript APIs used to interact with an HTML table can't be used to manage a Word.Table object. Instead, you must use the [Table APIs](/javascript/api/word/word.table) available in the Word Object Model to interact with Word.Table and related objects.

Similarly, don't use Word JavaScript APIs to interact with HTML table objects.

## See also

- [Troubleshoot development errors with Office Add-ins](../testing/troubleshoot-development-errors.md)
- [Troubleshoot user errors with Office Add-ins](../testing/testing-and-troubleshooting.md)
