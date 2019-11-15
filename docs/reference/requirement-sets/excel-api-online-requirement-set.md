---
title: Excel JavaScript API online-only requirement set
description: 'Details about the ExcelApiOnline requirement set'
ms.date: 11/15/2019
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

## API list

There are currently no online-only APIs. Check back as new features are added to Excel on the web and supported by the Office JavaScript APIs.

## See also

- [Excel JavaScript API Reference Documentation](/javascript/api/excel?view=excel-js-online)
- [Excel JavaScript preview APIs](./excel-preview-apis.md)
- [Excel JavaScript API requirement sets](./excel-api-requirement-sets.md)