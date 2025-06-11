---
title: Office JavaScript API reference documentation
description: Learn about the Office JavaScript APIs.
ms.author: lindalu
ms.topic: overview
ms.date: 06/11/2025
ms.localizationpriority: high
---

# Office JavaScript API reference documentation

An add-in can use the Office JavaScript APIs to interact with objects in Office client applications.

- **Application-specific** APIs provide strongly-typed objects that can be used to interact with objects that are native to a specific Office application.
- **Common** APIs can be used to access features such as UI, dialogs, and client settings that are common across multiple Office applications.

You should use application-specific APIs whenever feasible, and use Common APIs only for scenarios that aren't supported by application-specific APIs. For more detailed information about these two API models, see [Develop Office Add-ins](../develop/develop-overview.md#api-models).

## API reference

:::row:::
   :::column span="":::
        :::image type="content" source="../images/m365-app-excel.png" alt-text="Excel API reference" border="false":::
        <br>**Excel API reference**<br>[JavaScript APIs for building Excel add-ins](/javascript/api/excel)
   :::column-end:::
   :::column span="":::
        :::image type="content" source="../images/m365-app-outlook.png" alt-text="Outlook API reference":::
        <br>**Outlook API reference**<br>[JavaScript APIs for building Outlook add-ins](/javascript/api/outlook)
   :::column-end:::
   :::column span="":::
        :::image type="content" source="../images/m365-app-word.png" alt-text="Word API reference" border="false":::
        <br>**Word API reference**<br>[JavaScript APIs for building Word add-ins](/javascript/api/word)
   :::column-end:::
:::row-end:::
:::row:::
   :::column span="":::
        :::image type="content" source="../images/m365-app-powerpoint.png" alt-text="PowerPoint API reference" border="false":::
        <br>**PowerPoint API reference**<br>[JavaScript APIs for building PowerPoint add-ins](/javascript/api/powerpoint)
   :::column-end:::
   :::column span="":::
        :::image type="content" source="../images/m365-app-onenote.png" alt-text="OneNote API reference" border="false":::
        <br>**OneNote API reference**<br>[JavaScript APIs for building OneNote add-ins](/javascript/api/onenote)
   :::column-end:::
   :::column span="":::
   :::image type="content" source="../images/m365-app-office.png" alt-text="Common API reference" border="false":::
        <br>**Common API reference**<br>[JavaScript APIs that can be used by any Office Add-in](/javascript/api/office)
   :::column-end:::
:::row-end:::

**Note**: There's currently no application-specific JavaScript API for Project; you'll use Common APIs to create Project add-ins.
