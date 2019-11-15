---
title: Excel JavaScript API online-only requirement set
description: 'Details about the ExcelApi Online requirement set'
ms.date: 11/15/2019
ms.prod: excel
localization_priority: Normal
---

# Excel JavaScript API online-only requirement set

The `ExcelApi Online` requirement set is a special requirement set that includes features that are only available for Excel on the web. APIs in this requirement set are considered to be production APIs (not subject to undocumented behavioral or structural changes) for the Excel on the web host. `ExcelApi Online` are considered to be "preview" APIs for other platforms (Windows, Mac, iOS) and may not be supported by any of those platforms.

APIs in the `ExcelApi Online` requirement set will be moved to a numbered requirement set once they are supported on platforms. The members of the `ExcelApi Online` requirements set will change as features are added across the Excel ecosystem.

> [!IMPORTANT]
> `ExcelApi Online` is superset of the latest numbered requirement set.

## Recommended usage

Because `ExcelApi Online` APIs are only supported by Excel on the web, your add-in should check if the requirement set is supported before calling these APIs. This avoids calling an online-only API on a different platform.

```js
if (Office.context.requirements.isSetSupported("ExcelApi", "Online")) {
   // Any API exclusive to the ExcelApi Online requirement set.
}
```

Once the API is moved to a numbered requirement set, you should remove or edit the isSetSupported check. This will enable your add-in's feature on other platforms. Be sure to test the feature on those platforms when making this change.

## API list

There are currently no online-only APIs. Check back as new features are added to Excel on the web and supported by the Office JavaScript APIs.

## See also

- [Excel JavaScript API Reference Documentation](/javascript/api/excel?view=excel-js-online)
- [Excel JavaScript preview APIs](./excel-preview-apis.md)
- [Excel JavaScript API requirement sets](./excel-api-requirement-sets.md)