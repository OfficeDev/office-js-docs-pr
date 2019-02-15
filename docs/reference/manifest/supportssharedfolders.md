---
title: SupportsSharedFolders element in the manifest file
description: ''
ms.date: 10/09/2018
localization_priority: Normal
---

# SupportsSharedFolders element

Defines whether the Outlook add-in is available in delegate scenarios. The **SupportsSharedFolders** element is a child element of [DesktopFormFactor](desktopformfactor.md). It is set to *false* by default.

> [!IMPORTANT]
> This element is only available in the [Outlook add-ins Preview requirement set](../objectmodel/preview-requirement-set/outlook-requirement-set-preview.md) against Exchange Online. Add-ins that use this element cannot be published to AppSource or deployed via centralized deployment.

The following is an example of the  **SupportsSharedFolders** element.

```XML
<DesktopFormFactor>
  <FunctionFile resid="residDesktopFuncUrl" />
  <SupportsSharedFolders>true</SupportsSharedFolders>
  <ExtensionPoint xsi:type="PrimaryCommandSurface">
    <!-- information about this extension point -->
  </ExtensionPoint>

  <!-- You can define more than one ExtensionPoint element as needed -->

</DesktopFormFactor>
```
