---
title: Office Add-ins known issues
description: This article documents active and resolved issues with Office Add-ins.
ms.date: 11/26/2025
ms.localizationpriority: medium
---

# Office Add-ins known issues

_Last updated November 26, 2025_

This article provides information about current known issues with Office Add-ins. For more information about common error messages you might encounter, see [Troubleshoot user errors with Office Add-ins](/office/dev/add-ins/testing/testing-and-troubleshooting) or contact the add-in developer on the **Details + support** tab on the add-in's detail page in [Microsoft Marketplace](https://marketplace.microsoft.com).

## Active issues in Office add-ins

### Outlook: Outlook add-ins disappearing from ribbon

Customers are reporting their Outlook add-ins are missing from the ribbon and are also reporting difficulties deploying them in some tenants. 

#### STATUS

A fix is been released to affected tenants. 
 
#### START TIME

Nov 17 2025 5:38pm PST

### Outlook: Users may experience delays of up to ten seconds loading signature add-in images in Exchange Online

#### STATUS

We've determined that a recent update to an authentication component of attachment logging introduced a regression which is resulting in impact. We're reverting this update to resolve the issue.

Tracking ID: 706911563

#### IMPACT

Some users may experience delays of up to ten seconds loading images in Exchange Online. This section will be updated as our investigation continues.

#### START TIME

Monday, November 3, 2025, at 4:31 PM UTC

### Outlook: Delays loading inline images in email signatures in the new Outlook for Windows and Outlook for the web

We're currently investigating reports from Outlook users who are experiencing loading delays of inline images in email signatures when using the new Outlook for Windows and Outlook for the web. Our findings indicate that this is a server-side performance issue that affects rendering of all inline images. Attempting to send messages while the images are not yet loaded results in the following dialog box.

:::image type="content" source="../images/outlook-images-still-loading-error.png" alt-text="Outlook images still loading error message.":::

Tracking ID: 678890927

Client version: 20250822005.18

#### STATUS

We're still receiving isolated reports from some users regarding this previously resolved issue. While the issue has been largely mitigated, certain users in specific regions are still experiencing inline signature images loading slowly and the blocking dialog during email send. Because this stems from a server-side performance delay, the impact varies by customer and region. Those affected may see delays when loading inline images—particularly in scenarios involving signature add-ins. We're actively investigating this issue with highest priority.

#### WORKAROUND

Options:

1. Remove inline images from signature.
1. Wait for images to load before sending the file.
1. Switch to classic Outlook for Windows or Outlook for Mac.

### Centrally deployed add-in error "You don't have permission to use this add-in"

Numerous customers report that after updating Office from 2505 to 2507 their add-in will not load and an error is displayed "You don't have permission to use this add-in. Contact your system administrator." Any add-in may reproduce this issue; it is not specific to a single add-in.

 :::image type="content" source="../images/excel-web-add-in-permission-error.png" alt-text="Excel web add-in permissions error message.":::

Tracking ID: 667052546

Version affected: Office Monthly Enterprise 2507

#### STATUS

A fix is being deployed.

| Channel | Release timeline |
| --- | --- |
| Insiders | Available as of September 30th, 2025 |
| Current Channel | Available to install on October 7th, 2025 |
| Monthly Enterprise Channel | Available to install on October Patch Tuesday, October 14th, 2025 |

#### WORKAROUND

**Option 1: Refresh admin-managed add-ins**

1. Select **Home** > **Add-ins** in the ribbon.
1. Select **More add-ins**.
1. Go to the **Admin Managed** tab.
1. Select the **Refresh** button in top right.
1. The add-in should reappear. Open it to reload the add-in.

**Option 2: Forced admin refresh**

IT admins can force the add-ins to refresh by creating the following registry key.

Key: `HKCU\Software\Microsoft\Office\16.0\WEF\TrustedCatalogs\ClearInstalledExtensions`
Value: `DWORD = 1`

### Excel: Increased frequency of RichApi.Error: Error code: 0xF5320001

Date reported: 09/04/2025

Since late August, customers are seeing an increase in `RichApi.Error 0xF532001` in their error telemetry. This error happens only when the `Office.ribbon.requestUpdate` API is called immediately after the `Office.ribbon.requestCreateControls` API is called.

Tracking ID: 10529994

GitHub issue: [Increased frequency of RichApi.Error code 0xF5320001](https://github.com/OfficeDev/office-js/issues/6072)

#### STATUS

We're currently working on a fix.

#### WORKAROUND

Options:

1. When you make the initial `requestCreateControls` call, include the enabled/disabled state, if known. Instead of making two calls one right after the other, do it in one call.
1. Roll back Office from version 2508 to 2507.

## Recently resolved issues in Office Add-ins

### Excel: RichApi.Error code 0x8002802B known as hrNotFound is occurring more frequently when not expected

Date reported: 09/17/2025

Users might experience failures when executing Excel grid operations initiated through add-in commands on the ribbon or context menu. This issue occurs primarily when users have Custom Functions.

Platform affected: Windows Desktop

#### STATUS

Date fixed: 09/26/2025

Users should upgrade Excel to 2508 (19127.20264) or later for the fix.

### See also

- For more information about resolved issues in Office Add-ins, see the [Office-js closed issues in GitHub](https://github.com/OfficeDev/office-js/issues?q=is%3Aissue%20state%3Aclosed).
