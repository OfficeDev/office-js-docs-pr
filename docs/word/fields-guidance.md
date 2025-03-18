---
title: Use fields in your Word add-in
description: Learn to use fields in your Word add-in.
ms.date: 03/18/2025
ms.localizationpriority: medium
---

# Use fields in your Word add-in

A [field](https://support.microsoft.com/office/c429bbb0-8669-48a7-bd24-bab6ba6b06bb) in a Word document is a placeholder. It allows you to provide instructions for the content instead of the content itself. You can use fields to create and format a Word template. Word documents support a number of [field types](https://support.microsoft.com/office/1ad6d91a-55a7-4a8d-b535-cf7888659a51), many with associated parameters for configuring the field. However, Word on the web generally doesn't support adding or editing fields through the UI. For more information, see [Field codes in Word for the web](https://support.microsoft.com/office/d8f46094-13c3-4966-98c3-259748f3caf1).

Starting from the WordApi 1.5 requirement set, Word JavaScript APIs allow you to manage fields in your Word add-in. In all platforms, you can get existing fields. You can insert, update, and delete fields in platforms that support those capabilities.

The following sections discuss several of the most frequently used field types: Addin, Date, Hyperlink, and TOC (Table of Contents).

## Addin field

The Addin field is meant to store add-in data that's hidden from the Word user interface, regardless of whether fields in the document are set to show or hide its content. The Addin field isn't available in the Word UI's **Field** dialog box. Use the API to insert the Addin field type and set the field's data. Unlike other fields, Word on the web does allow you to update the Addin field.

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
      //const data = $("#input-reference").val(); // Get data from UI.
      const data = "Insert your data here";
      firstAddinField.data = data;
      firstAddinField.load("data");
      await context.sync();

      console.log("The data of the Addin field after being set:", firstAddinField.data);
    }
  });
}
```

## Date field

The Date field inserts the current date according to the format you specify. You can toggle between displaying the date or the field code by setting the `showCodes` field property to `false` or `true` respectively.

```javascript
// Inserts a Date field before selection.
async function rangeInsertDateField() {
  await Word.run(async (context) => {
    let range = context.document.getSelection().getRange();
    const field = range.insertField(Word.InsertLocation.before, Word.FieldType.date, '\\@ "M/d/yyyy h:mm am/pm"', true);
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

[snippet](https://github.com/OfficeDev/office-js-snippets/blob/prod/samples/word/50-document/manage-fields.yaml)

[Field codes: Date field](https://support.microsoft.com/office/d0c7e1f1-a66a-4b02-a3f4-1a1c56891306)

## Hyperlink field

The Hyperlink field inserts the address of either a location in the same document or an external location (for example, a webpage). When the user selects it, they're navigated to the specified location. You can toggle between displaying the hyperlink address or the field code by setting the `showCodes` field property to `false` or `true` respectively.

TODO: Include code example and link to Script Lab snippet

```javascript
```

[Field codes: Hyperlink field](https://support.microsoft.com/office/864f8577-eb2a-4e55-8c90-40631748ef53)

## TOC (Table of Contents) field

The TOC field inserts a table of contents, which lists certain areas of a document, like headings. You can toggle between displaying the table of contents or the field code by setting the `showCodes` field property to `false` or `true` respectively.

TODO: Include code example and link to Script Lab snippet

```javascript
```

[Field codes: TOC (Table of Contents) field](https://support.microsoft.com/office/1f538bc4-60e6-4854-9f64-67754d78d05c)

## See also

- [Field codes in Word for the web](https://support.microsoft.com/office/d8f46094-13c3-4966-98c3-259748f3caf1)
- [Insert, edit, and view fields in Word](https://support.microsoft.com/office/c429bbb0-8669-48a7-bd24-bab6ba6b06bb)
