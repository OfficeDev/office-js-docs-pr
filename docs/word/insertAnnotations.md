---
title: insertAnnotations method (Word JavaScript API)
description: Adds ephemeral critique annotations (such as grammar underlines) to a paragraph in Word using Office.js.
ms.date: 4/16/2025
ms.topic: reference
---

# insertAnnotations (Word)

> [!NOTE]
> This API is now supported as part of `AnnotationSet`, but this page provides developer-focused usage guidance, behavior notes, and real-world context for effective use in Word add-ins.

## Summary

The `insertAnnotations` method allows developers to apply non-destructive critique annotations (such as grammar-style underlines) to a Word paragraph. These annotations are applied in-memory, do not modify the document's content, and disappear once the document is closed.

This behavior is particularly useful for real-time grammar checkers and assistive tools where ephemeral, contextual UI is needed.

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

| Parameter | Type     | Description |
|-----------|----------|-------------|
| `start`   | number   | The start index (within the paragraph text) for the critique |
| `length`  | number   | Number of characters the annotation should cover |
| `colorScheme` | enum or string | The color used for the underline (e.g. `Word.CritiqueColorScheme.red`) |
| `popupOptions` | object | Contains metadata for the annotation tooltip |

## Returns

- A list of annotation IDs (via `ClientResult<string[]>`)

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

  paragraph.insertAnnotations(annotationSet);
  await context.sync();
});
```

## Notes

- These annotations are **not persisted** in the `.docx` file
- They are **ephemeral** and will disappear when the document is reloaded
- They are intended for **transient UI**, such as grammar checkers
- You can retrieve and manage them via `getAnnotationById()`

## Related

- [`AnnotationSet`](https://learn.microsoft.com/javascript/api/word/word.annotationset)
- [`getAnnotationById()`](https://learn.microsoft.com/javascript/api/word/word.document?view=word-js-preview#word-document-getannotationbyidid)
- [`setAnnotation()`](https://learn.microsoft.com/javascript/api/word/word.range?view=word-js-preview#setannotationoptions-)
