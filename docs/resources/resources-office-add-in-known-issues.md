---
title: Office Add-ins known issues
description: This article documents active and resolved issues with Office Add-ins.
ms.date: 09/23/2025
ms.localizationpriority: medium
---

# Office Add-ins known issues

_Last updated 09/23/2025_

This article provides information about current known issues with Office Add-ins. For more information about common error messages you might encounter, see [Troubleshoot user errors with Office Add-ins](/office/dev/add-ins/testing/testing-and-troubleshooting) or contact the add-in developer on the **Details + support** tab on the add-in's detail page in [AppSource](https://appsource.microsoft.com).

## Active issues in Office Add-ins

### Outlook: Delays loading inline images in email signatures in the new Outlook for Windows and Outlook for the web

#### ISSUE

We're currently investigating reports from Outlook users who are experiencing loading delays of inline images in email signatures when using the new Outlook for Windows and Outlook for the web. Our findings indicate that this is a server-side performance issue that affects rendering of all inline images. Attempting to send messages while the images are not yet loaded results in the following dialog box.

 ![Outlook images still loading error message.](../images/outlook-images-still-loading-error.png)

Tracking ID: 678890927.

#### STATUS

We're actively investigating this issue with high priority. Because it stems from a server-side performance delay, the impact varies by customer and region. While not all users will experience the issue, those affected may see delays when loading inline images—particularly in scenarios involving signature add-ins.

#### WORKAROUND

1. Remove inline images from signature.
1. Wait for images to load before sending the file. 
1. Switch to classic Outlook for Windows or Outlook for Mac.

### Excel: Centrally deployed add-in error "You don't have permission to use this add-in"

#### ISSUE

Numerous customers report that after updating Office from 2505 to 2507 their add-in will not load and an error is displayed "You don't have permission to use this add-in. Contact your system administrator." Any add-in may reproduce this issue; it is not specific to a single add-in.

 ![Excel web add-in permissions error message.](../images/excel-web-add-in-permission-error.png)

Tracking ID: 667052546.

Version affected: Office Monthly Enterprise 2507.

#### STATUS

We're currently working on a fix.

#### WORKAROUND

1. Roll back Office to previous version 2505.

### Excel: Increased frequency of RichApi.Error: Error code: 0xF532001

#### ISSUE

Since late August, customers are seeing an increase of RichApi.Error 0xF532001 in their error telemetry.

Tracking ID: 679969584.

#### STATUS

We're currently working on a fix.

#### WORKAROUND

1. Roll back Office to version 2505.

## Recently resolved issues in Office Add-ins

For more information about resolved issues in Office Add-ins, see the [Office-js closed issues in GitHub](https://github.com/OfficeDev/office-js/issues?q=is%3Aissue%20state%3Aclosed).
