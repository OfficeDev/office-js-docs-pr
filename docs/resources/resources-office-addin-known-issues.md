---
title: Office Add-ins known issues
description: This article documents active and resolved issues with Office Add-ins.
ms.date: 09/25/2025
ms.localizationpriority: medium
---

# Troubleshoot Office Add-ins

_Last updated 09/25/2025_

Use information in this article to help you resolve current known issues in Office Add-ins.  For more information about common error messages you might encounter, see [Troubleshoot user errors with Office Add-ins](./testing/testing-and-troubleshooting) or contact the add-in developer on the **Details + support** tab on the add-in's detail page in [AppSource](https://appsource.microsoft.com).

## Active issues in Office Add-ins

### Outlook: Delay in sending email in New Outlook for Windows

**ISSUE**
Outlook customers report that there is an ongoing issue where emails composed in New Outlook for Windows are stuck in the Outbox and not sent. Our investigations indicate that this Outlook issue affects signature add-ins, including CodeTwo, and causes delays in sending emails due to slow inline image loading.

 ![Outlook images still loading error message.](../images/outlook-images-still-loading-error.png)

Tracking ID: 678890927.

The versions affected are 20250829003.06 and 20250829003.07.

**STATUS**
The Outlook team has deployed a fix to Dogfood and is validating it. Rollout to production rollout expected to start soon (week of September 22 or sooner.)

**WORKAROUND**
Uninstall signature add-ins and/or remove inline images from signature.

## Resolved issues in Office Add-ins
