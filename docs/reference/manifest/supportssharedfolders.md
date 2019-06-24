---
title: SupportsSharedFolders element in the manifest file
description: ''
ms.date: 03/01/2019
localization_priority: Normal
---

# SupportsSharedFolders element

Defines whether the Outlook add-in is available in delegate scenarios. The **SupportsSharedFolders** element is a child element of [DesktopFormFactor](desktopformfactor.md). It is set to *false* by default.

> [!IMPORTANT]
> Delegate access for Outlook add-ins is currently in preview and only supported in clients that run against Exchange Online. Add-ins that use this element cannot be published to AppSource or deployed via centralized deployment.

The following is an example of the  **SupportsSharedFolders** element.

```XML
<DesktopFormFactor>
  <FunctionFile resid="residDesktopFuncUrl" />
  <SupportsSharedFolders>true</SupportsSharedFolders>
  <ExtensionPoint xsi:type="MessageReadCommandSurface">
    <!-- configure selected extension point -->
  </ExtensionPoint>

  <!-- You can define more than one ExtensionPoint element as needed -->

</DesktopFormFactor>
```
