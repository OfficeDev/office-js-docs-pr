---
title: Contextual tabs in Office Add-ins
description: 'Learn how to add contextual tabs to your Office Add-in.'
ms.date: 11/09/2020
localization_priority: Normal
---

# Contextual tabs in Office Add-ins (preview)

A contextual tab is a hidden tab control in the Office ribbon that is displayed in the tab row when an object in the Office document, such as an image or a table has focus; for example, the **Table Design** tab that appears on the Excel ribbon when a table is selected. You can include custom contextual tabs in your Office add-in and specify when they are visible or hidden.

> [!IMPORTANT]
> Contextual tabs are in preview. Please experiment with them in a development or testing environment but don't add them to a production add-in.
>
> Contextual tabs are currently only supported on Excel and only on these platforms and builds:
>
>* Excel on Windows (Microsoft 365 only, not perpetual license): Version ???? (Build ?????.?????) Your Microsoft 365 subscription may need to be on the [Current Channel (Preview)](https://insider.office.com/en-us/join/windows) formerly called "Monthly Channel (Targeted)" or "Insider Slow".

> [!NOTE]
> Contextual tabs work only on platforms that support the following requirement sets. For more about requirement sets and how to work with them, see [Specify Office applications and API requirements](../develop/specify-office-hosts-and-api-requirements.md).
>
> - [SharedRuntime 1.1](../reference/requirement-sets/shared-runtime-requirement-sets.md)

The following are the major steps for including a contextual tab in an add-in

1. Define the groups and controls that appear on the tab.
1. Specify the circumstances when the tab will be visible.


### Configure the add-in to use a shared runtime

Adding custom keyboard shortcuts requires your add-in to use the shared runtime. For more information, [Configure an add-in to use a shared runtime](../excel/configure-your-add-in-to-use-a-shared-runtime.md).

