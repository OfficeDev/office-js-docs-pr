---
title: Add and manage fields in Word add-ins
description: Use the Word JavaScript API to add, read, update, and delete fields in your Word add-in.
ms.date: 05/21/2026
ms.localizationpriority: medium
ms.topic: how-to
---

# Use fields in your Word add-in

A [field](https://support.microsoft.com/office/c429bbb0-8669-48a7-bd24-bab6ba6b06bb) in Word is a placeholder that displays instructions, generated content, or document metadata instead of fixed text. Use fields when you want a document to update itself, such as a template with a date, a link, or a table of contents.

Word documents support several [field types](https://support.microsoft.com/office/1ad6d91a-55a7-4a8d-b535-cf7888659a51), and many accept parameters that control how the field behaves. Word on the web generally doesn't support adding or editing fields through the UI. For more information, see [Field codes in Word for the web](https://support.microsoft.com/office/d8f46094-13c3-4966-98c3-259748f3caf1). In all platforms, you can get existing fields. In supported platforms, you can insert, update, and delete fields.

This article demonstrates how to use these common field types:

- **[Addin](#addin-field)**: Insert an Addin field to store hidden add-in data.
- **[Date](#date-field)**: Insert a Date field to generate a current date value.
- **[Hyperlink](#hyperlink-field)**: Insert a Hyperlink field to point to a web page or location in the document.
- **[TOC (Table of Contents)](#toc-table-of-contents-field)**: Insert a TOC field to build a table of contents from headings.

## Addin field

The Addin field stores add-in data that's hidden from the Word user interface, even when fields in the document are set to show or hide their content. The Addin field isn't available in the Word UI's **Field** dialog box. Use the API to insert the Addin field type and set the field's data.

The following code sample shows how to insert an Addin field before the cursor location or your selection in the Word document.

```javascript
// Inserts an Addin field before selection.
async function rangeInsertAddinField() {
  await Word.run(async (context) => {
    let range = context.document.getSelection().getRange();
    const field = range.insertField(Word.InsertLocation.before, Word.FieldType.addin);
    field.load("result,code");
    await context.sync();

    if (field.isNullObject) {
      console.log("There are no fields in this document.");
    } else {
      console.log("Code of the field: " + field.code);
      console.log("Result of the field: " + JSON.stringify(field.result));
    }
  });
}
```

The following code sample shows how to get the first Addin field found in a document then set that field's data property.

```javascript
// Gets the first Addin field in the document and sets its data.
async function getFirstAddinFieldAndSetData() {
  await Word.run(async (context) => {
    let myFieldTypes = new Array();
    myFieldTypes[0] = Word.FieldType.addin;
    const addinFields = context.document.body.fields.getByTypes(myFieldTypes);
    let fields = addinFields.load("items");
    await context.sync();

    if (fields.items.length === 0) {
      console.log("No Addin fields in this document.");
    } else {
      fields.load();
      await context.sync();

      const firstAddinField = fields.items[0];
      firstAddinField.load("code,result,data");
      await context.sync();

      console.log("The data of the Addin field before being set:", firstAddinField.data);
      const data = "Insert your data here";
      //const data = $("#input-reference").val(); // Or get user data from your add-in's UI.
      firstAddinField.data = data;
      firstAddinField.load("data");
      await context.sync();

      console.log("The data of the Addin field after being set:", firstAddinField.data);
    }
  });
}
```

## Date field

The Date field inserts the current date in the format you specify. You can toggle between displaying the date or the field code by setting the `showCodes` field property to `false` or `true`, respectively.

The following code sample shows how to insert a Date field before the cursor location or your selection in the Word document.

```javascript
// Inserts a Date field before selection.
async function rangeInsertDateField() {
  await Word.run(async (context) => {
    let range = context.document.getSelection().getRange();
    const field = range.insertField(
      Word.InsertLocation.before,
      Word.FieldType.date,
     '\\@ "M/d/yyyy h:mm am/pm"',
     true
    );
    field.load("result,code");
    await context.sync();

    if (field.isNullObject) {
      console.warn("The field wasn't inserted as expected.");
    } else {
      console.log("Code of the field: " + field.code);
      console.log("Result of the field: " + JSON.stringify(field.result));
    }
  });
}
```

### Further reading for the Date field

- [Manage Fields code sample](https://github.com/OfficeDev/office-js-snippets/blob/prod/samples/word/50-document/manage-fields.yaml)
- [Field codes: Date field](https://support.microsoft.com/office/d0c7e1f1-a66a-4b02-a3f4-1a1c56891306)

## Hyperlink field

The Hyperlink field inserts the address of a location in the same document or an external location such as a web page. When the user selects it, they're taken to the specified location. You can toggle between displaying the hyperlink address or the field code by setting the `showCodes` field property to `false` or `true`, respectively.

The following code sample shows how to insert a Hyperlink field before the cursor location or your selection in the Word document.

```javascript
// Inserts a Hyperlink field before selection.
async function rangeInsertHyperlinkField() {
  await Word.run(async (context) => {
    let range = context.document.getSelection().getRange();
    const field = range.insertField(
      Word.InsertLocation.before,
      Word.FieldType.hyperlink,
      "https://bing.com",
      true
    );
    field.load("result,code");
    await context.sync();

    if (field.isNullObject) {
      console.warn("The field wasn't inserted as expected.");
    } else {
      console.log("Code of the field: " + field.code);
      console.log("Result of the field: " + JSON.stringify(field.result));
    }
  });
}
```

### Further reading for the Hyperlink field

- [Field codes: Hyperlink field](https://support.microsoft.com/office/864f8577-eb2a-4e55-8c90-40631748ef53)

## TOC (Table of Contents) field

The TOC field inserts a table of contents that lists document sections such as headings. You can toggle between displaying the table of contents or the field code by setting the `showCodes` field property to `false` or `true`, respectively.

The following code sample shows how to insert a TOC field at the cursor location or replace your current selection in the Word document.

```javascript
/**
 1. Run setup.
 2. Select "[To place table of contents]" paragraph.
 3. Run rangeInsertTOCField.
 */

// Inserts a TOC (table of contents) field replacing selection.
async function rangeInsertTOCField() {
  await Word.run(async (context) => {
    let range = context.document.getSelection().getRange();
    const field = range.insertField(
      Word.InsertLocation.replace,
      Word.FieldType.toc
    );
    field.load("result,code");
    await context.sync();

    if (field.isNullObject) {
      console.warn("The field wasn't inserted as expected.");
    } else {
      console.log("Code of the field: " + field.code);
      console.log("Result of the field: " + JSON.stringify(field.result));
    }
  });
}

// Prep document so there'll be elements that could be included in a table of contents.
async function setup() {
  await Word.run(async (context) => {
    const body: Word.Body = context.document.body;
    body.clear();
    body.insertParagraph("Document title", "End").styleBuiltIn = Word.BuiltInStyleName.title;
    body.insertParagraph("[To place table of contents]", "End").styleBuiltIn = Word.BuiltInStyleName.normal;
    body.insertParagraph("Introduction", "End").styleBuiltIn = Word.BuiltInStyleName.heading1;
    body.insertParagraph("Paragraph 1", "End").styleBuiltIn = Word.BuiltInStyleName.normal;
    body.insertParagraph("Topic 1", "End").styleBuiltIn = Word.BuiltInStyleName.heading1;
    body.insertParagraph("Paragraph 2", "End").styleBuiltIn = Word.BuiltInStyleName.normal;
    body.insertParagraph("Topic 2", "End").styleBuiltIn = Word.BuiltInStyleName.heading1;
    body.insertParagraph("Paragraph 3", "End").styleBuiltIn = Word.BuiltInStyleName.normal;
  });
}
```

### Further reading for the TOC field

- [Field codes: TOC (Table of Contents) field](https://support.microsoft.com/office/1f538bc4-60e6-4854-9f64-67754d78d05c)

## See also

- [Field codes in Word for the web](https://support.microsoft.com/office/d8f46094-13c3-4966-98c3-259748f3caf1)
- [Insert, edit, and view fields in Word](https://support.microsoft.com/office/c429bbb0-8669-48a7-bd24-bab6ba6b06bb)
