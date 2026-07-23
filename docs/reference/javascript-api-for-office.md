---
title: Office JavaScript API reference documentation
description: Find the Office JavaScript API reference for each Office host and choose the right API model for your add-in.
ms.topic: overview
ms.date: 07/23/2026
ms.localizationpriority: high
---

# Office JavaScript API reference documentation

Use this page to quickly find the Office JavaScript API reference for the Office host from which you're building an add-in.

## Choose the right API model

- Use **application-specific** APIs when you're building for Excel, Outlook, PowerPoint, Word, or OneNote and need access to host-specific objects and features.
- Use **common** APIs for cross-host features such as UI, dialogs, and client settings.

Start with application-specific APIs whenever possible. Use common APIs for scenarios that application-specific APIs don't support. For more information about these API models, see [Develop Office Add-ins](../develop/develop-overview.md#api-models).

Project add-ins currently use common APIs because there isn't an application-specific JavaScript API for Project.

## API reference by host

:::row:::
   :::column span="":::
        :::image type="content" source="../images/m365-app-excel.png" alt-text="Excel API reference" border="false":::
        **Excel API reference**
        [JavaScript APIs for building Excel add-ins](/javascript/api/excel)
   :::column-end:::
   :::column span="":::
        :::image type="content" source="../images/m365-app-outlook.png" alt-text="Outlook API reference":::
        **Outlook API reference**
        [JavaScript APIs for building Outlook add-ins](/javascript/api/outlook)
   :::column-end:::
   :::column span="":::
        :::image type="content" source="../images/m365-app-word.png" alt-text="Word API reference" border="false":::
        **Word API reference**
        [JavaScript APIs for building Word add-ins](/javascript/api/word)
   :::column-end:::
:::row-end:::
:::row:::
   :::column span="":::
        :::image type="content" source="../images/m365-app-powerpoint.png" alt-text="PowerPoint API reference" border="false":::
        **PowerPoint API reference**
        [JavaScript APIs for building PowerPoint add-ins](/javascript/api/powerpoint)
   :::column-end:::
   :::column span="":::
        :::image type="content" source="../images/m365-app-onenote.png" alt-text="OneNote API reference" border="false":::
        **OneNote API reference**
        [JavaScript APIs for building OneNote add-ins](/javascript/api/onenote)
   :::column-end:::
   :::column span="":::
   :::image type="content" source="../images/m365-app-office.png" alt-text="Common API reference" border="false":::
        **Common API reference**
        [JavaScript APIs that can be used by any Office Add-in](/javascript/api/office)
   :::column-end:::
:::row-end:::
