---
title: Office versions and requirement sets
description: Supported Office.js platforms using JavaScript API.
ms.date: 05/26/2022
ms.localizationpriority: high
---

# Office versions and requirement sets

There are many versions of Office on several platforms, and they don't all support every API in Office JavaScript API (Office.js). You may not always have control over the version of Office your users have installed.  To handle this situation, we provide a system called requirement sets to help you determine whether an Office application supports the capabilities you need in your Office Add-in.

> [!NOTE]
>
> - Office runs across multiple platforms, including Windows, in a browser, Mac, and iPad.
> - Examples of Office applications are Office Products: Excel, Word, PowerPoint, Outlook, OneNote, and so forth.  
> - A requirement set is a named group of API members e.g., `ExcelApi 1.5`, `WordApi 1.3`, and so on.  

## How to check your Office version

To identify the Office version that you're using, from within an Office application, select the **File** menu, and then choose **Account**. The version of Office will appear in the **Product Information** section. For example, the following screenshot indicates Office Version 1802 (Build 9026.1000).

![Checking your Office version.](../images/office-version.png)

## Office requirement sets availability

Office Add-ins can use API requirement sets to determine whether the Office application supports the API members that it need to use. Requirement set support varies by Office application and the Office application version (see previous section).

Some Office applications have their own API requirement sets. For example, the first requirement set for the Excel API was `ExcelApi 1.1` and the first requirement set for the Word API was `WordApi 1.1`. Since then, multiple new ExcelApi requirement sets and WordApi requirement sets have been added to provide additional API functionality.

In addition, other functionality such as add-in commands (ribbon extensibility) and the ability to launch dialog boxes (Dialog API) were added to the Common API. Add-in commands and Dialog API requirement sets are examples of API sets that the various Office applications share in common.

An add-in can only use APIs in requirement sets that are supported by the version of Office application where the add-in is running. To know exactly which requirement sets are available for a specific Office application version, refer to the following application-specific requirement set articles.

- [Excel JavaScript API requirement sets](/javascript/api/requirement-sets/excel/excel-api-requirement-sets) (ExcelApi)
- [Word JavaScript API requirement sets](/javascript/api/requirement-sets/word/word-api-requirement-sets) (WordApi)
- [OneNote JavaScript API requirement sets](/javascript/api/requirement-sets/onenote/onenote-api-requirement-sets) (OneNoteApi)
- [PowerPoint JavaScript API requirement sets](/javascript/api/requirement-sets/powerpoint/powerpoint-api-requirement-sets) (PowerPointApi)
- [Understanding Outlook API requirement sets](/javascript/api/requirement-sets/outlook/outlook-api-requirement-sets) (Mailbox)

Some requirement sets contain APIs that can be used by any Office application. For information about these requirement sets, refer to the following articles.

- [Office common requirement sets](/javascript/api/requirement-sets/common/office-add-in-requirement-sets)
- [Add-in commands requirement sets](/javascript/api/requirement-sets/common/add-in-commands-requirement-sets)
- [Dialog API requirement sets](/javascript/api/requirement-sets/common/dialog-api-requirement-sets)
- [Dialog Origin requirement sets](/javascript/api/requirement-sets/common/dialog-origin-requirement-sets)
- [Identity API requirement sets](/javascript/api/requirement-sets/common/identity-api-requirement-sets)
- [Image Coercion requirement sets](/javascript/api/requirement-sets/common/image-coercion-requirement-sets)
- [Keyboard Shortcuts requirement sets](/javascript/api/requirement-sets/common/keyboard-shortcuts-requirement-sets)
- [Open Browser Window requirement sets](/javascript/api/requirement-sets/common/open-browser-window-api-requirement-sets)
- [Ribbon API requirement sets](/javascript/api/requirement-sets/common/ribbon-api-requirement-sets)
- [Shared Configure your Office Add-in to use a shared runtime requirement sets](/javascript/api/requirement-sets/common/shared-Configure your Office Add-in to use a shared runtime-requirement-sets)

The version number of a requirement set, such as the "1.1" in `ExcelApi 1.1`, is relative to the Office application. The version number of a given requirement set (e.g., `ExcelApi 1.1`) does not correspond to the version number of Office.js or to requirement sets for other Office applications (e.g., Word, Outlook, etc.).  Requirement sets for the different Office applications are released at different rates. For example, `ExcelApi 1.5` was released before the `WordApi 1.3` requirement set.

The Office JavaScript API library (Office.js) includes all requirement sets that are currently available. While there is such a thing as requirement sets `ExcelApi 1.3` and `WordApi 1.3`, there is no `Office.js 1.3` requirement set. The latest release of Office.js is maintained as a single Office endpoint delivered via the content delivery network (CDN). For more details around the Office.js CDN, including how versioning and backward compatibility is handled, see [Understanding the Office JavaScript API](../develop/understanding-the-javascript-api-for-office.md).

## Specify Office applications and requirement sets

There are various ways to specify which Office applications and requirement sets are required by an add-in.  For detailed information, see [Specify Office applications and API requirements](../develop/specify-office-hosts-and-api-requirements.md)

## See also

- [Specify Office applications and API requirements](../develop/specify-office-hosts-and-api-requirements.md)
- [Install the latest version of Office](../develop/install-latest-office-version.md)
- [Overview of update channels for Microsoft 365 Apps](/deployoffice/overview-of-update-channels-for-office-365-proplus)
- [Reimagine productivity with Microsoft 365 and Microsoft Teams](https://products.office.com/compare-all-microsoft-office-products?tab=2)
