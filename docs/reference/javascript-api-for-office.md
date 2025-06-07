---
title: Office JavaScript API reference documentation
description: Learn about the Office JavaScript APIs.
author: lindalu-msft
ms.author: lindalu
ms.topic: overview
ms.date: 06/06/25
ms.localizationpriority: high
---

# Office JavaScript API reference documentation

An add-in can use the Office JavaScript APIs to interact with objects in Office client applications.

- **Application-specific** APIs provide strongly-typed objects that can be used to interact with objects that are native to a specific Office application.
- **Common** APIs can be used to access features such as UI, dialogs, and client settings that are common across multiple Office applications.

You should use application-specific APIs whenever feasible, and use Common APIs only for scenarios that aren't supported by application-specific APIs. For more detailed information about these two API models, see [Develop Office Add-ins](../develop/develop-overview.md#api-models).

## API reference

|                           |                      |           |
| :------------------------ | -------------------- | ----------------|
| <img src="../images/index/logo-excel.svg" width="100"> </br>**Excel API reference**</br>[JavaScript APIs for building Excel add-ins](/javascript/api/excel).  | <img src="../images/index/logo-outlook.svg" width="100"> </br>**Outlook API reference**</br>[JavaScript APIs for building Outlook add-ins](/javascript/api/outlook). | <img src="../images/index/logo-word.svg" width="100"> </br>**Word API reference**</br>[JavaScript APIs for building Word add-ins](/javascript/api/word). |
| <img src="../images/index/logo-powerpoint.svg" width="100"> </br>**PowerPoint API reference**</br>[JavaScript APIs for building PowerPoint add-ins](/javascript/api/powerpoint).  | <img src="../images/index/logo-onenote.svg" width="100"> </br>**OneNote API reference**</br>[JavaScript APIs for building OneNote add-ins](/javascript/api/onenote). | <img src="../images/index/logo-office.png" width="100"> </br>**Common API reference**</br>[JavaScript APIs that can be used by any Office Add-in](/javascript/api/office). |

**Note**: There's currently no application-specific JavaScript API for Project; you'll use Common APIs to create Project add-ins.
