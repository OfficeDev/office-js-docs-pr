---
title: Work with notes using the Excel JavaScript API
description: Information on using the APIs to add, remove, and edit notes.
ms.date: 06/26/2025
ms.localizationpriority: medium
---

# Work with notes using the Excel JavaScript API

This article describes how to add, change, and remove notes in a workbook with the Excel JavaScript API. You can learn more about notes from the [Insert comments and notes in Excel](https://support.microsoft.com/office/bdcc9f5d-38e2-45b4-9a92-0b2b5c7bf6f8) article. For information about the differences between notes and comments, see [The difference between threaded comments and notes](https://support.microsoft.com/office/the-difference-between-threaded-comments-and-notes-75a51eec-4092-42ab-abf8-7669077b7be3).

Notes are tied to an individual cell. Anyone viewing the workbook with sufficient permissions can view a note. Notes in a workbook are tracked by the `Workbook.notes` property. This includes notes created by users and also notes created by your add-in. The `Workbook.notes` property is a [NoteCollection](/javascript/api/excel/excel.notecollection) object that contains a collection of [Note](/javascript/api/excel/excel.note) objects. Notes are also accessible at the [Worksheet](/javascript/api/excel/excel.worksheet) level.

> [!TIP]
> To learn about adding and editing comments with the Excel JavaScript API, see [Work with comments using the Excel JavaScript API](excel-add-ins-comments.md).

## Add a note

Use the `NoteCollection.add` method to add notes to a workbook. This method takes two parameters:

- `cellAddress`: The cell where the comment is added. This can either be a string or [Range](/javascript/api/excel/excel.range) object. The range must be a single cell.
- `content`: The comment's content, as a string.

The following code sample shows how to add a note to the selected cell in a worksheet.

```js
await Excel.run(async (context) => {
    // This function adds a note to the selected cell.
    const selectedRange = context.workbook.getSelectedRange();

    // Note that an InvalidArgument error is thrown if multiple cells are selected.
    context.workbook.notes.add(selectedRange, "The first note.");
    await context.sync();
});
```

## Change note visibility

By default, the content of a note is hidden unless a user hovers over the cell with the note or sets the workbook to display notes. To display a note, use the [Note.visible](/javascript/api/excel/excel.note#excel-excel-note-visible-member) property. The following code sample shows how to change the visibility of a note.

```js
await Excel.run(async (context) => {
    // This function sets the note on cell A1 to visible.
    const sheet = context.workbook.worksheets.getActiveWorksheet();
    const firstNote = sheet.notes.getItem("A1");

    firstNote.load();
    await context.sync();

    firstNote.visible = true;
});
```

## Edit the content of a note

To edit the content of a note, use the [Note.content](/javascript/api/excel/excel.note#excel-excel-note-content-member) property. The following sample shows how to change the content of the first note in the `NoteCollection`.

```js
await Excel.run(async (context) => {
    // This function changes the content in the first note.
    const sheet = context.workbook.worksheets.getActiveWorksheet();
    const note = sheet.notes.getItemAt(0);

    note.content = "Changing the content of the first note.";
    await context.sync();
});
```

> [!NOTE]
> Use the `Note.authorName` property to get the author of a note. The author name is a read-only property.

## Change the size of a note

To make notes larger or smaller, use the [Note.height](/javascript/api/excel/excel.note#excel-excel-note-height-member) and [Note.width](/javascript/api/excel/excel.note#excel-excel-note-width-member) properties.

The following sample shows how to set the size of the first note in the `NoteCollection`.

```js
await Excel.run(async (context) => {
    // This function changes the height and width of the first note.
    const sheet = context.workbook.worksheets.getActiveWorksheet();
    const note = sheet.notes.getItemAt(0);

    note.width = 400;
    note.height = 200;    

    await context.sync();
});
```

## Delete a note

To delete a note, use the [Note.delete](/javascript/api/excel/excel.note#excel-excel-note-delete-member(1)) method. The following sample shows how to delete the note attached to cell **A2**.

```js
await Excel.run(async (context) => {
    // This function deletes the note from cell A2.
    const sheet = context.workbook.worksheets.getActiveWorksheet();
    const note = sheet.notes.getItem("A2");

    note.delete();
    await context.sync();
});
```

## See also

- [Excel JavaScript object model in Office Add-ins](excel-add-ins-core-concepts.md)
- [Work with workbooks using the Excel JavaScript API](excel-add-ins-workbooks.md)
- [Work with comments using the Excel JavaScript API](excel-add-ins-comments.md)
- [Insert comments and notes in Excel](https://support.microsoft.com/office/bdcc9f5d-38e2-45b4-9a92-0b2b5c7bf6f8)
