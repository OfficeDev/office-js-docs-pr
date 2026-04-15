---
title: Office versions and requirement sets
description: Supported Office.js platforms using JavaScript API.
ms.date: 02/26/2026
ms.localizationpriority: high
---

# Office versions and requirement sets

There are many versions of Office on several platforms, and they don't all support every API in the Office JavaScript Library (Office.js). You may not always have control over the version of Office your users have installed. To handle this situation, we provide a system called "requirement sets" or "capabilities" to help you determine whether an Office application supports the capabilities you need in your Office Add-in. A requirement set is a named group of API members. Some examples are, `BindingEvents`, `AddinCommands`, `ExcelApi 1.5`, and `WordApi 1.3`.

> [!NOTE]
>
> - Office runs across multiple platforms, including Windows, in a browser, Mac, and iPad.
> - Office is available by a Microsoft 365 subscription or perpetual license. The perpetual version is available by volume-licensing agreement or retail.

## How to check your Office version

### Office on Windows

To identify the Office on Windows version, from within an Office application, select the **File** menu, and then choose **Account**. The version of Office appears in the **Product Information** section. For example, the following screenshot indicates Office Version 1802 (Build 9026.1000).

:::image type="content" source="../images/office-version.png" alt-text="A mock-up of the Product Information section of the Account page in Excel on Windows. There is a subsection titled About Excel and below this is text reading 'Version 1802 (Build 9026.1-000 Click-to-run)'":::

> [!NOTE]
> If your version of Office is different from this image, see [What version of Outlook do I have?](https://support.microsoft.com/office/b3a9568c-edb5-42b9-9825-d48d82b2257c) or [About Office: What version of Office am I using?](https://support.microsoft.com/topic/932788b8-a3ce-44bf-bb09-e334518b8b19) to understand how to get this information for your version.

### Office on Mac

To identify the Office on Mac version, follow the guidance at [About Office: What version of Office am I using?](https://support.microsoft.com/en-us/office/about-office-what-version-of-office-am-i-using-932788b8-a3ce-44bf-bb09-e334518b8b19#osversion=macos).

## Deployment

How your add-in is deployed can affect your add-in's availability on the various platforms and clients. To learn more about deployment options, see [Deploy and publish Office Add-ins](../publish/publish.md).

## Office requirement sets availability

Office Add-ins can use API requirement sets to determine whether the Office application supports the API members that it needs to use. Requirement set support varies by Office application, Office application version, and the platform. For detailed information about the platforms, requirement sets, and Common APIs that each Office application supports, see [Office client application and platform availability for Office Add-ins](/javascript/api/requirement-sets). See also the earlier section [How to check your Office version](#how-to-check-your-office-version).

An add-in can only use APIs in requirement sets that are supported by the version of Office application where the add-in is running. Most APIs in the [Common APIs](understand-the-javascript-api-for-office.md#api-models) of the Office JavaScript Library are grouped into requirement sets according to the feature that they support; for example, the `AddinCommands` requirement sets support ribbon customization, and the `DialogApi` requirement set supports the Office Dialog. For information about Common API requirement sets, see [Office Common API requirement sets](/javascript/api/requirement-sets/common/office-add-in-requirement-sets).

Office applications that have [application-specific APIs](understand-the-javascript-api-for-office.md#api-models) (Excel, OneNote, PowerPoint, and Word) also have explicit application-specific requirement sets. For example, the first requirement set for the Excel API was `ExcelApi 1.1` and the first requirement set for the Word API was `WordApi 1.1`. Outlook also has application-specific requirement sets even though it doesn't have application-specific API architecture the way that Excel, OneNote, PowerPoint, and Word do. There are certain subsets of the Common APIs that are only used in Outlook. These are grouped into requirement sets named `Mailbox`, such as `Mailbox 1.1`.

To know exactly which application-specific requirement sets are available for a specific Office application version, refer to the following application-specific requirement set articles.

- [Excel JavaScript API requirement sets](/javascript/api/requirement-sets/excel/excel-api-requirement-sets) (ExcelApi)
- [OneNote JavaScript API requirement sets](/javascript/api/requirement-sets/onenote/onenote-api-requirement-sets) (OneNoteApi)
- [Outlook JavaScript API requirement sets](/javascript/api/requirement-sets/outlook/outlook-api-requirement-sets) (Mailbox)
- [PowerPoint JavaScript API requirement sets](/javascript/api/requirement-sets/powerpoint/powerpoint-api-requirement-sets) (PowerPointApi)
- [Word JavaScript API requirement sets](/javascript/api/requirement-sets/word/word-api-requirement-sets) (WordApi)

The version number of an application-specific requirement set, such as the "1.1" in `ExcelApi 1.1`, is relative to the Office application. The version number of a given requirement set (for example, `ExcelApi 1.1`) doesn't correspond to the version number of Office.js or to requirement sets for other Office applications.  Requirement sets for the different Office applications are released at different rates. For example, `ExcelApi 1.5` was released before the `WordApi 1.3` requirement set.

The Office JavaScript API library (Office.js) includes all the APIs in all the requirement sets. While there is such a thing as requirement sets `ExcelApi 1.3` and `WordApi 1.3`, there is no `Office.js 1.3` requirement set. The latest release of Office.js is maintained as a single Office endpoint delivered via the content delivery network (CDN). For more details around the Office.js CDN, including how versioning and backward compatibility is handled, see [Understanding the Office JavaScript API](../develop/understand-the-javascript-api-for-office.md).

## Specify Office applications and requirement sets

There are various ways to specify which Office applications and requirement sets are required by an add-in, and it is important to understand how Office processes requirement set configuration and runtime checking. We strongly recommend that you read the following articles.

- [Specify Office applications and API requirements with the add-in only manifest](../develop/specify-office-hosts-and-api-requirements.md)
- [Specify Office applications and API requirements with the unified manifest for Microsoft 365](../develop/specify-office-hosts-and-api-requirements-unified.md)
- [How to use the "requirements" property in the unified manifest for Microsoft 365](requirements-property-unified-manifest.md)
- [Understand the logic of API requirement configuration](understand-requirement-configuration.md)
- [Check for API availability at runtime](specify-api-requirements-runtime.md)

## See also

- [Install the latest version of Office](../develop/install-latest-office-version.md)
- [Overview of update channels for Microsoft 365 Apps](/deployoffice/overview-of-update-channels-for-office-365-proplus)
- [Reimagine productivity with Microsoft 365 and Microsoft Teams](https://www.microsoft.com/microsoft-365/buy/compare-all-microsoft-365-products)
