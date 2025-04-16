---
title: insertAnnotations method (Word JavaScript API)
description: Adds ephemeral critique annotations (e.g., grammar underlines) to Word paragraphs using Office.js without modifying document content.
ms.date: 4/16/2025
ms.topic: reference
---

# insertAnnotations (Word)

> [!NOTE]
> This API is not officially documented as of now, but it has been observed in production use by third-party add-ins. This page is a community contribution based on active experimentation, testing.

## Summary

The `insertAnnotations` method adds non-destructive, in-line critique annotations (such as grammar suggestions) to a Word paragraph. These annotations appear as underlines in the Word UI but do not modify the actual document content and are not saved in the `.docx` file.

Annotations are ephemeral — they disappear when the document is closed and do not appear in exported document XML.

## Syntax

```javascript
paragraph.insertAnnotations({
  critiques: [
    {
      start: number,
      length: number,
      colorScheme: "Red" | "Green" | "Blue" | "Orange" | string
    }
  ]
});
```

## Parameters

| Parameter    | Type    | Description |
|--------------|---------|-------------|
| `start`      | number  | The 0-based index (within the paragraph) where the annotation starts. |
| `length`     | number  | The number of characters to annotate. |
| `colorScheme`| string  | The underline color. Known working values include: `"Red"`, `"Green"`, `"Blue"`, `"Orange"`. |

> [!TIP]
> Although not officially documented, the `colorScheme` property supports a wider range of named colors. Tested values include `"Red"`, `"Green"`, `"Blue"`, and `"Orange"`. Other standard color names may also work.

## Requirements

- Word API requirement set: `WordApi 1.5` (or higher)
- Supported on: Word Desktop (latest builds), Word Online
- Must be called on a valid `Paragraph` object

## Returns

None. The method applies the annotation directly.

## Example

```javascript
await Word.run(async (context) => {
  const paragraph = context.document.getSelection().paragraphs.getFirst();
  paragraph.load("text");
  await context.sync();

  // Add a red underline to the first 5 characters
  paragraph.insertAnnotations({
    critiques: [
      {
        start: 0,
        length: 5,
        colorScheme: "Red"
      }
    ]
  });

  await context.sync();
});
```

## Test Multiple Colors

```javascript
const colors = ['Red', 'Green', 'Blue', 'Orange'];
colors.forEach(async (color) => {
  await Word.run(async (context) => {
    const paragraph = context.document.getSelection().paragraphs.getFirst();
    paragraph.insertAnnotations({
      critiques: [{
        start: 0,
        length: 5,
        colorScheme: color
      }]
    });
    await context.sync();
  });
});
```

## Notes

- The annotation behaves like native Word grammar underlines.
- Hover and click behavior may vary slightly depending on Word version.
- Use in combination with `getAnnotationById()` for tracking or jumping to issues.
- Annotations are not persisted to the document file and will be lost when the document is closed unless re-added.

## Related

- `getAnnotationById(id)`
- `critiqueAnnotation.range`
- [`context.document.annotations`](https://learn.microsoft.com/javascript/api/word/word.document?view=word-js-preview&preserve-view=true#annotations)