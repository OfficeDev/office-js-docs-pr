---
title: Check for API availability at runtime
description: Learn how to verify at runtime that the Office application supports your add-in's API calls.
ms.topic: best-practice
ms.date: 02/12/2025
ms.localizationpriority: medium
---

# Check for API availability at runtime

If your add-in uses a specific extensibility feature for some of its functionality, but has other useful functionality that doesn't require the extensibility feature, you should design the add-in so that it's installable on platform and Office version combinations that don't support the extensibility feature. It can provide a valuable, albeit diminished, experience on those combinations.

When the difference in the two experiences consists entirely of differences in the Office JavaScript Library APIs that are called, and not in any features that are configured in the manifest, then you test at runtime to discover whether the user's Office client supports an [API requirement set](office-versions-and-requirement-sets.md). You can also [test at runtime whether APIs that aren't in a set are supported](#check-for-setless-api-support).

> [!NOTE]
> To provide alternate experiences with features that require manifest configuration, follow the guidance in [Specify Office hosts and API requirements with the unified manifest](specify-office-hosts-and-api-requirements-unified.md) or [Specify Office applications and API requirements with the add-in only manifest](specify-office-hosts-and-api-requirements.md).

## Check for requirement set support

The [isSetSupported](/javascript/api/office/office.requirementsetsupport#office-office-requirementsetsupport-issetsupported-member(1)) method is used to check for requirement set support. Pass the requirement set's name and the minimum version as parameters. If the requirement set is supported, `isSetSupported` returns `true`. The following code shows an example.

```js
if (Office.context.requirements.isSetSupported("WordApi", "1.2")) {
   // Code that uses API members from the WordApi 1.2 requirement set.
} else {
   // Provide diminished experience here.
   // For example, run alternate code when the user's Word is
   // volume-licensed perpetual Word 2016 (which doesn't support WordApi 1.2).
}
```

About this code, note:

- The first parameter is required. It's a string that represents the name of the requirement set. For more information about available requirement sets, see [Office Add-in requirement sets](/javascript/api/requirement-sets/common/office-add-in-requirement-sets).
- The second parameter is optional. It's a string that specifies the minimum requirement set version that the Office application must support in order for the code within the `if` statement to run (for example, "1.9"). If not used, version "1.1" is assumed.

> [!WARNING]
> When calling the `isSetSupported` method, the value of the second parameter (if specified) should be a string, not a number. The JavaScript parser can't differentiate between numeric values such as 1.1 and 1.10, whereas it can for string values such as "1.1" and "1.10".

The following table shows the requirement set names for the application-specific API models.

|Office application|RequirementSetName|
|---|---|
|Excel|ExcelApi|
|OneNote|OneNoteApi|
|Outlook|Mailbox|
|PowerPoint|PowerPointApi|
|Word|WordApi|

The following is an example of using the method with one of the Common API model requirement sets.

```js
if (Office.context.requirements.isSetSupported('CustomXmlParts')) {
    // Run code that uses API members from the CustomXmlParts requirement set.
} else {
    // Run alternate code when the user's Office application doesn't support the CustomXmlParts requirement set.
}
```

> [!NOTE]
> The `isSetSupported` method and the requirement sets for these applications are available in the latest Office.js file on the CDN. If you don't use Office.js from the CDN, your add-in might generate exceptions if you are using an old version of the library in which `isSetSupported` is undefined. For more information, see [Use the latest Office JavaScript API library](specify-office-hosts-and-api-requirements-unified.md#use-the-latest-office-javascript-api-library).

## Check for setless API support 

When your add-in depends on a method that isn't part of a requirement set, called a setless API, use a runtime check to determine whether the method is supported by the Office application, as shown in the following code example. For a complete list of methods that don't belong to a requirement set, see [Office Add-in requirement sets](/javascript/api/requirement-sets/common/office-add-in-requirement-sets#methods-that-arent-part-of-a-requirement-set).

> [!NOTE]
> We recommend that you limit the use of this type of runtime check in your add-in's code.

The following code example checks whether the Office application supports `document.setSelectedDataAsync`.

```js
if (Office.context.document.setSelectedDataAsync) {
    // Run code that uses `document.setSelectedDataAsync`.
}
```

## See also
- [Office requirement sets availability](office-versions-and-requirement-sets.md#office-requirement-sets-availability)
- [Specify Office hosts and API requirements with the unified manifest](specify-office-hosts-and-api-requirements-unified.md)
- [Specify Office hosts and API requirements with the add-in only manifest](specify-office-hosts-and-api-requirements.md)
