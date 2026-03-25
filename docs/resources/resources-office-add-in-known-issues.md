---
title: Office Add-ins known issues
description: This article documents active and resolved issues with Office Add-ins.
ms.date: 03/23/2026
ms.localizationpriority: medium
---

# Office Add-ins known issues

_Last updated 03/23/2026_

This article provides information about current known issues with Office Add-ins. For more information about common error messages you might encounter, see [Troubleshoot user errors with Office Add-ins](/office/dev/add-ins/testing/testing-and-troubleshooting) or contact the add-in developer on the **Details + support** tab on the add-in's detail page in [Microsoft Marketplace](https://marketplace.microsoft.com).

<!-------------Copy and paste this line and the following ***. Paste between each issue for readability. (Inserts a line in topic) ---------------------------------------------------->
***

## ISSUE: Intermittent failure to load or deploy Office Add-ins

Some users are experiencing issues where Office Add-ins appear missing when deployed through the Microsoft 365 Admin Center. In affected scenarios, add-ins are visible in the admin experience but do not render or appear correctly in Office & Outlook clients.  

### STATUS

We are currently investigating the cause.  

### START DATE

Date reported: 03/23/2026 

<!-------------Copy and paste this line and the following ***. Paste between each issue for readability. (Inserts a line in topic) ---------------------------------------------------->
***

## ISSUE: PIM-enabled tenants are unable to deploy or manage Office Add-ins via centralized deployment

When using Azure AD Privileged Identity Management (PIM) to activate admin roles, there is a regression where PIM-enabled admin roles are not correctly honored. During centralized deployment, role-based access control (RBAC) authorization fails and leads to false permission denials during add‑in deployment and management flows.

### STATUS

Open; tracking id: 11126536

### IMPACT

Admins are unable to deploy Office Add-ins via central deployment when using PIM-enabled admin roles.

### WORKAROUND

Don't use PIM-enabled roles if you are blocked by this issue.

<!-------------Copy and paste this line and the following ***. Paste between each issue for readability. (Inserts a line in topic) ---------------------------------------------------->
***

## OUTLOOK ISSUE: Users unable to access the My Templates add-in in Exchange Online across all Outlook clients

Users report that **My Templates** add-in is missing and undiscoverable across all Outlook surfaces. The add-in subscription exists on affected mailboxes and Centralized Deployment returns it correctly, but client-side discovery within Outlook and Exchange services fails to surface the add-in to end users. Users cannot find it in the toolbar, ribbon, **Add Apps** search, All Apps, Integrated Apps in Admin Center, or via PowerShell Get-App in some cases. The issue presents as a service-side discovery or authentication regression, rather than an admin configuration or Centralized Deployment failure.

### STATUS

We're currently working on a fix.

### START DATE

Date reported: 01/20/2026

### DETAILS

Impacted add-ins: My Templates (primary); Viva Insights (confirmed also impacted as of March 3, 2026); other default add-ins (Bing Maps, Unsubscribe, Common Actions) intermittently affected.

Severity level: High

Affected platforms/clients: Outlook Classic (Desktop, Windows),  New Outlook (Desktop, Windows), Outlook on the web, Outlook mobile

### USER IMPACT

Widespread, multi-tenant impact. Impact is tenant-wide in most cases.

### CAUSE

Partially identified. Engineering has confirmed two contributing factors:

1. A recent backend change that switched authentication from Exchange Web Services (EWS) to REST for the My Templates add-in caused access errors. The REST auth change was rolled back on March 3, 2026. This produced a significant drop in errors, but full remediation has not been achieved. The subscription is present on the mailbox, but add-in information is not returned to clients.
2. Historical/recurring root cause: A prior wave was resolved via rollback + cache resets in December 2025 — but some tenants never fully recovered.

### WORK AROUND (steps to mitigate)

No reliable universal workaround exists. The following steps have been attempted by support teams with limited/inconsistent success:

1. **Global Admin PowerShell — re-enable the add-in org-wide** (may take up to 72 hours to reflect; some tenants encounter 401 errors):
   ```PowerShell

   Set-App -Identity a216ceed-7791-4635-a752-5a4ac0a5eb93 -OrganizationApp -Enabled $true

   ```
1. **Verify the add-in status**:
   ```PowerShell

   Get-App -Identity a216ceed-7791-4635-a752-5a4ac0a5eb93

   ```
1. **Refresh the Outlook client** — In some cases, a page refresh or Outlook restart triggered the add-in to reappear temporarily.
1. **Submit in-app feedback with diagnostic logs** — Go to **Help** > **Feedback** > **Report a Problem in Outlook** and share the Session ID / User ID with support so engineering can pull diagnostics.
1. **Reference the public support article** — See [My Templates are missing from Outlook](https://support.microsoft.com/office/34967a7a-7a80-4d72-bb45-a43ecdc93678).

### NOTES TO ADMIN

Re-enabling the add-in via PowerShell or the Admin Center does not guarantee resolution while the service-side issue is active. Engineering is working on a fix and will post updates to the Service Health Dashboard (SHD).

<!-------------Copy and paste this line and the following ***. Paste between each issue for readability. (Inserts a line in topic) ---------------------------------------------------->
***

## ISSUE: Intermittent failure to load or deploy Office Add-ins due to Exchange authentication changes

Some users experience issues where Office add-ins failed to load, appeared missing, or could not be deployed through the Microsoft 365 admin center. In affected scenarios, add-ins were visible in the admin experience or store but did not render or appear correctly in Outlook or other Office clients.

### START DATE

Reported by: Microsoft Support / Customer Reports on: 02/25/2026

### DETAILS

Impacted add-ins: Admin-deployed and organization-scoped Office add-ins (including third‑party add-ins)

Severity level: Medium

Current status: Open, mitigation has begun rolling out.

Affected platforms/clients: Office clients, Microsoft 365 admin center (centralized deployment experience)

### USER IMPACT

Intermittent authentication failures cause Office add-ins to appear missing or fail to deploy. This primarily affects tenants impacted by recent Exchange Web Services (EWS) security enforcement changes.

### CAUSE

As part of ongoing Exchange Web Services (EWS) security improvements, Microsoft enforced stricter authentication requirements that no longer allow certain legacy authentication methods. Some add-in service calls were still relying on these legacy paths, causing add-in metadata retrieval requests to be rejected. As a result, affected add-ins could not be loaded or displayed correctly for users.

### WORK AROUND (steps to mitigate)
No customer action was required. Microsoft applied targeted mitigations to restore compatibility while a longer-term fix is validated. Customers who continue to experience issues are advised to contact Microsoft Support for assistance.

### SEE ALSO

For more information, see [Deprecation of Exchange Web Services (EWS) in Exchange Online](/exchange/clients-and-mobile-in-exchange-online/deprecation-of-ews-exchange-online).

<!-------------Copy and paste this line and the following ***. Paste between each issue for readability. (Inserts a line in topic) ---------------------------------------------------->
***

## OUTLOOK ISSUE: Users may experience delays of up to ten seconds loading signature add-in images in Exchange Online

### STATUS

We've determined that a recent update to an authentication component of attachment logging introduced a regression which is resulting in impact. We're reverting this update to resolve the issue.

Tracking ID: 706911563

### IMPACT

Some users may experience delays of up to ten seconds loading images in Exchange Online. This section will be updated as our investigation continues.

### START DATE

Monday, 11/03/2025, at 4:31 PM UTC

<!-------------Copy and paste this line and the following ***. Paste between each issue for readability. (Inserts a line in topic) ---------------------------------------------------->
***

## OUTLOOK ISSUE: Delays loading inline images in email signatures in the new Outlook for Windows and Outlook for the web

We're currently investigating reports from Outlook users who are experiencing loading delays of inline images in email signatures when using the new Outlook for Windows and Outlook for the web. Our findings indicate that this is a server-side performance issue that affects rendering of all inline images. Attempting to send messages while the images are not yet loaded results in the following dialog box.

:::image type="content" source="../images/outlook-images-still-loading-error.png" alt-text="Outlook images still loading error message.":::

Tracking ID: 678890927

Client version: 20250822005.18

### STATUS

We're still receiving isolated reports from some users regarding this previously resolved issue. While the issue has been largely mitigated, certain users in specific regions are still experiencing inline signature images loading slowly and the blocking dialog during email send. Because this stems from a server-side performance delay, the impact varies by customer and region. Those affected may see delays when loading inline images—particularly in scenarios involving signature add-ins. We're actively investigating this issue with highest priority.

### WORKAROUND

Options:

1. Remove inline images from signature.
1. Wait for images to load before sending the file.
1. Switch to classic Outlook for Windows or Outlook for Mac.

<!-------------Copy and paste this line and the following ***. Paste between each issue for readability. (Inserts a line in topic) ---------------------------------------------------->
***

## ISSUE: Centrally deployed add-in error "You don't have permission to use this add-in"

Numerous customers report that after updating Office from 2505 to 2507 their add-in will not load and an error is displayed "You don't have permission to use this add-in. Contact your system administrator." Any add-in may reproduce this issue; it is not specific to a single add-in.

 :::image type="content" source="../images/excel-web-add-in-permission-error.png" alt-text="Excel web add-in permissions error message.":::

Tracking ID: 667052546

Version affected: Office Monthly Enterprise 2507

### STATUS

A fix is being deployed.

| Channel | Release timeline |
| --- | --- |
| Insiders | Available as of September 30th, 2025 |
| Current Channel | Available to install on October 7th, 2025 |
| Monthly Enterprise Channel | Available to install on October Patch Tuesday, October 14th, 2025 |

### WORKAROUND

#### Option 1: Refresh admin-managed add-ins

1. Select **Home** > **Add-ins** in the ribbon.
1. Select **More add-ins**.
1. Go to the **Admin Managed** tab.
1. Select the **Refresh** button in top right.
1. The add-in should reappear. Open it to reload the add-in.

#### Option 2: Forced admin refresh

IT admins can force the add-ins to refresh by creating the following registry key.

Key: `HKCU\Software\Microsoft\Office\16.0\WEF\TrustedCatalogs\ClearInstalledExtensions`
Value: `DWORD = 1`

<!-------------Copy and paste this line and the following ***. Paste between each issue for readability. (Inserts a line in topic) ---------------------------------------------------->
***

## EXCEL ISSUE: Increased frequency of RichApi.Error: Error code: 0xF5320001

Since late August, customers are seeing an increase in `RichApi.Error 0xF532001` in their error telemetry. This error happens only when the `Office.ribbon.requestUpdate` API is called immediately after the `Office.ribbon.requestCreateControls` API is called.

Tracking ID: 10529994

GitHub issue: [Increased frequency of RichApi.Error code 0xF5320001](https://github.com/OfficeDev/office-js/issues/6072)

### STATUS

We're currently working on a fix.

### START DATE

Reports began late August 2025. Date opened: 09/04/2025

### WORKAROUND

Options:

1. When you make the initial `requestCreateControls` call, include the enabled/disabled state, if known. Instead of making two calls one right after the other, do it in one call.
1. Roll back Office from version 2508 to 2507.

<!-- --------RESOLVED SECTION: Move resolved issues to the top of this section. Delete after 90 days.-------- -->

<!-------------Copy and paste this line and the following ***. Paste between each issue for readability. ADD "RESOLVED:" to H2---------------------------------------------------->
***

## RESOLVED: Microsoft Marketplace: Issues installing add-ins from the Marketplace

Some users may experienced failures when installing add-ins from the Microsoft Marketplace. During the installation flow, the process may not complete successfully, and users may see a 50x server-related error.

### STATUS

The issue is now resolved.

### START TIME

Sunday, 02/08/2026

### RESOLUTION TIME

Friday, 02/13/2026

<!-------------Copy and paste this line and the following ***. Paste between each issue for readability. ADD "RESOLVED:" to H2---------------------------------------------------->
***

## RESOLVED: Outlook for Mac: Signatures not inserted using add-ins and user with Smart Alerts add-ins not able to send email

A disruption in processing `LaunchEvent` caused the following issues:

- Signatures were not stamped on outgoing emails.
- Users with Smart Alerts add-ins were unable to send emails in some cases.

### STATUS

Resolved. This issue was caused by a temporary configuration issue during a backend change management update for event-based activation support. For a subset of users having event-based add-ins, this resulted in add‑ins not initializing as expected, which in turn blocked sending emails. The configuration has now been corrected.

Note: Because these settings are cached locally and sync asynchronously, some users may need to restart Outlook more than once to pick up the updated configuration.

Tracking ID: 734492427

### START TIME

Thursday, 01/15/2026 5:45am PST

### RESOLUTION TIME

The fix was released Thursday, 01/15/2026 7:00am PST

<!-------------Copy and paste this line and the following ***. Paste between each issue for readability. ADD "RESOLVED:" to H2---------------------------------------------------->
***

## RESOLVED: EXCEL: RichApi.Error code 0x8002802B known as hrNotFound is occurring more frequently when not expected

Users experienced failures when executing Excel grid operations initiated through add-in commands on the ribbon or context menu. This issue occured primarily when users have Custom Functions.

Platform affected: Windows Desktop

### STATUS

Users should upgrade Excel to 2508 (19127.20264) or later for the fix.

### START DATE

Date reported: SEP 17, 2025

### RESOLUTION DATE

Date fixed: 09/26/2025

<!------------LEAVE SEE ALSO---------------------------------------------------->
***

#### SEE ALSO

[Fixes or workaround for recent issues in classic Outlook for Windows](https://support.microsoft.com/office/fixes-or-workarounds-for-recent-issues-in-classic-outlook-for-windows-ecf61305-f84f-4e13-bb73-95a214ac1230)
[Office-js resolved issues in GitHub](https://github.com/OfficeDev/office-js/issues?q=is%3Aissue%20state%3Aclosed)
[Deprecation of Exchange Web Services (EWS) in Exchange Online](/exchange/clients-and-mobile-in-exchange-online/deprecation-of-ews-exchange-online)
