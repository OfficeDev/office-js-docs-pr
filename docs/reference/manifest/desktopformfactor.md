---
title: DesktopFormFactor element in the manifest file
description: Specifies the settings for an add-in for the desktop form factor.
ms.date: 09/29/2021
ms.localizationpriority: medium
---

# DesktopFormFactor element

Specifies the settings for an add-in for the desktop form factor. The desktop form factor includes Office on the web, Windows, and Mac. It contains all the add-in information for the desktop form factor except for the **Resources** node.

Each DesktopFormFactor definition contains the **FunctionFile** element and one or more **ExtensionPoint** elements. For more information, see [FunctionFile element](functionfile.md) and [ExtensionPoint element](extensionpoint.md).

**Add-in type:** Task pane, Mail

**Valid only in these VersionOverrides schemas**:

- Task pane 1.0
- Mail 1.0
- Mail 1.1

For more information, see [Version overrides in the manifest](../../develop/add-in-manifests.md#version-overrides-in-the-manifest).

## Child elements

| Element                               | Required | Description  |
|:--------------------------------------|:--------:|:-------------|
| [ExtensionPoint](extensionpoint.md)   | Yes      | Defines where an add-in exposes functionality. |
| [FunctionFile](functionfile.md)       | Yes      | A URL to a file that contains JavaScript functions.|
| [GetStarted](getstarted.md)           | No       | Defines the callout that appears when installing the add-in in Word, Excel, or PowerPoint. If omitted, the callout uses the values from the [DisplayName](displayname.md) and [Description](description.md) elements instead. |
| [SupportsSharedFolders](supportssharedfolders.md) | No | Defines whether the Outlook add-in is available in shared mailbox (now in preview) and shared folders (that is, delegate access) scenarios. Set to *false* by default. |

## DesktopFormFactor example

```xml
...
<Hosts>
  <Host xsi:type="Presentation">
    <DesktopFormFactor>
      <FunctionFile resid="residDesktopFuncUrl" />
      <GetStarted>
        <!-- GetStarted callout -->
      </GetStarted>
      <ExtensionPoint xsi:type="PrimaryCommandSurface">
        <!-- Information on this extension point. -->
      </ExtensionPoint>
      <!-- Possibly more ExtensionPoint elements. -->
    </DesktopFormFactor>
  </Host>
</Hosts>
...
```
