---
title: Excel JavaScript API online-only requirement set
description: 'Details about the ExcelApiOnline requirement set'
ms.date: 12/05/2019
ms.prod: excel
localization_priority: Normal
---

# Excel JavaScript API online-only requirement set

The `ExcelApiOnline` requirement set is a special requirement set that includes features that are only available for Excel on the web. APIs in this requirement set are considered to be production APIs (not subject to undocumented behavioral or structural changes) for the Excel on the web host. `ExcelApiOnline` are considered to be "preview" APIs for other platforms (Windows, Mac, iOS) and may not be supported by any of those platforms.

When APIs in the `ExcelApiOnline` requirement set are supported across all platforms, they will added to the next released requirement set (`ExcelApi 1.[NEXT]`). Once that new requirement is public, those APIs will be removed from `ExcelApiOnline`. Think of this as a similar promotion process as an API moving from preview to release.

> [!IMPORTANT]
> `ExcelApiOnline` is superset of the latest numbered requirement set.

> [!IMPORTANT]
> `ExcelApiOnline 1.1` is the only version of the online-only APIs. This is because Excel on the web will always have a single version available to users that is the latest version.

## Recommended usage

Because `ExcelApiOnline` APIs are only supported by Excel on the web, your add-in should check if the requirement set is supported before calling these APIs. This avoids calling an online-only API on a different platform.

```js
if (Office.context.requirements.isSetSupported("ExcelApiOnline", "1.1")) {
   // Any API exclusive to the ExcelApiOnline requirement set.
}
```

Once the API is in a cross-platform requirement set, you should remove or edit the `isSetSupported` check. This will enable your add-in's feature on other platforms. Be sure to test the feature on those platforms when making this change.

> [!IMPORTANT]
> Your manifest cannot specify `ExcelApiOnline 1.1` as an activation requirement. It is not a valid value to use in the [Set element](../manifest/set.md).

## API list

The following APIs are currently available for Excel on the web as part of the `ExcelApiOnline 1.1` requirement set.

| Class | Fields | Description |
|:---|:---|:---|
|[Comment](/javascript/api/excel/excel.comment)|[mentions](/javascript/api/excel/excel.comment#mentions)|Gets the entities (e.g. people) that are mentioned in comments.|
||[richContent](/javascript/api/excel/excel.comment#richcontent)|Gets the rich comment content (e.g. mentions in comments). This string is not meant to be displayed to end-users. Your add-in should only use this to parse rich comment content.|
||[updateMentions(contentWithMentions: Excel.CommentRichContent)](/javascript/api/excel/excel.comment#updatementions-contentwithmentions-)|Updates the comment content with a specially formatted string and a list of mentions.|
|[CommentMention](/javascript/api/excel/excel.commentmention)|[email](/javascript/api/excel/excel.commentmention#email)|Gets or sets the email address of the entity that is mentioned in comment.|
||[id](/javascript/api/excel/excel.commentmention#id)|Gets or sets the id of the entity. This is aligned with the id information in `CommentRichContent.richContent`.|
||[name](/javascript/api/excel/excel.commentmention#name)|Gets or sets the name of the entity that is mentioned in comment.|
|[CommentReply](/javascript/api/excel/excel.commentreply)|[mentions](/javascript/api/excel/excel.commentreply#mentions)|Gets the entities (e.g. people) that are mentioned in comments.|
||[richContent](/javascript/api/excel/excel.commentreply#richcontent)|Gets the rich comment content (e.g. mentions in comments). This string is not meant to be displayed to end-users. Your add-in should only use this to parse rich comment content.|
||[updateMentions(contentWithMentions: Excel.CommentRichContent)](/javascript/api/excel/excel.commentreply#updatementions-contentwithmentions-)|Updates the comment content with a specially formatted string and a list of mentions.|
|[CommentRichContent](/javascript/api/excel/excel.commentrichcontent)|[mentions](/javascript/api/excel/excel.commentrichcontent#mentions)|An array containing all the entities (e.g. people) mentioned within the comment.|
||[richContent](/javascript/api/excel/excel.commentrichcontent#richcontent)||
|[Range](/javascript/api/excel/excel.range)|[moveTo(destinationRange: Range \| string)](/javascript/api/excel/excel.range#moveto-destinationrange-)|Moves cell values, formatting, and formulas from current range to the destination range, replacing the old information in those cells.|
|[RangeFormat](/javascript/api/excel/excel.rangeformat)|[adjustIndent(amount: number)](/javascript/api/excel/excel.rangeformat#adjustindent-amount-)|Adjusts the indentation of the range formatting. The indent value ranges from 0 to 250.|

## See also

- [Excel JavaScript API Reference Documentation](/javascript/api/excel?view=excel-js-online)
- [Excel JavaScript preview APIs](./excel-preview-apis.md)
- [Excel JavaScript API requirement sets](./excel-api-requirement-sets.md)