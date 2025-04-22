---
title: insertAnnotations method (Word JavaScript API)
description: Adds ephemeral critique annotations (such as grammar underlines) to a paragraph in Word using Office.js.
ms.date: 04/22/2025
ms.topic: reference
---

# insertAnnotations (Word JavaScript API)

> [!NOTE]
> This API is part of the `AnnotationSet` feature in Word JavaScript API and is designed for applying non-persistent, UI-focused annotations. This page provides developer-focused usage, behavior notes, and real-world examples for use in Word add-ins.

## Summary

The `insertAnnotations` method allows developers to programmatically apply ephemeral critique annotations to a Word paragraph. These annotations appear as grammar-style underlines with contextual tooltips and are:

- **Non-destructive**: They don’t modify the document content.
- **Ephemeral**: They disappear once the document is closed or reloaded.
- **UI-oriented**: Ideal for grammar checkers, intelligent suggestions, and contextual UI.

## Syntax

```javascript
paragraph.insertAnnotations({
  critiques: [
    {
      start: number,
      length: number,
      colorScheme: Word.CritiqueColorScheme.red,
      popupOptions: {
        titleResourceId: string,
        subtitleResourceId: string,
        brandingTextResourceId: string,
        suggestions: string[]
      }
    }
  ]
});
```

## Parameters

| Property        | Type                            | Description |
|-----------------|----------------------------------|-------------|
| `start`         | `number`                         | Start index (character offset) within the paragraph text where the critique should begin. |
| `length`        | `number`                         | Number of characters the critique annotation should cover. |
| `colorScheme`   | `Word.CritiqueColorScheme`       | Specifies the color of the underline (e.g., `red`, `green`, `blue`). |
| `popupOptions`  | `object`                         | Metadata for the annotation’s tooltip. See below for its properties. |

### `popupOptions` properties

| Property                   | Type     | Description |
|----------------------------|----------|-------------|
| `titleResourceId`          | `string` | Resource ID for the tooltip title. |
| `subtitleResourceId`       | `string` | Resource ID for the tooltip subtitle. |
| `brandingTextResourceId`   | `string` | Resource ID for the tooltip's branding or footer text. |
| `suggestions`              | `string[]` | A list of suggested alternative texts shown in the tooltip. |

## Returns

Returns a `ClientResult<string[]>` object containing the list of unique annotation IDs created.

You must call `await context.sync()` before accessing the `.value` property to retrieve the IDs:

```javascript
const annotationIds = paragraph.insertAnnotations(annotationSet);
await context.sync();
console.log("Inserted annotation IDs:", annotationIds.value); // Example: ["id1", "id2"]
```

These IDs can be used with other annotation APIs, such as `getAnnotationById()`.

## Example

```javascript
await Word.run(async (context) => {
  const paragraph = context.document.getSelection().paragraphs.getFirst();

  const popupOptions = {
    brandingTextResourceId: "Demo.Branding",
    titleResourceId: "Demo.Title",
    subtitleResourceId: "Demo.Subtitle",
    suggestions: ["suggestion 1", "suggestion 2"]
  };

  const annotationSet = {
    critiques: [
      {
        start: 0,
        length: 4,
        colorScheme: Word.CritiqueColorScheme.red,
        popupOptions
      }
    ]
  };

  const annotationIds = paragraph.insertAnnotations(annotationSet);
  await context.sync();

  console.log("Inserted annotation IDs:", annotationIds.value);
});
```

## Notes

- Annotations are **in-memory only** and do **not persist** in the `.docx` file.
- They are **removed** when the document is closed or reloaded.
- Useful for **real-time UI enhancements** (e.g., spelling/grammar checkers, writing suggestions).
- Use returned IDs to manage critiques (e.g., retrieve or delete specific annotations).

## Related Links

- [AnnotationSet](https://learn.microsoft.com/javascript/api/word/word.annotationset)
- [getAnnotationById()](https://learn.microsoft.com/javascript/api/word/word.document?view=word-js-preview#word-document-getannotationbyidid)
- [setAnnotation()](https://learn.microsoft.com/javascript/api/word/word.range?view=word-js-preview#setannotationoptions-)
