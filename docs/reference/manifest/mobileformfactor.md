---
title: MobileFormFactor element in the manifest file
description: The MobileFormFactor element specifies the mobile form factor settings for an add-in.
ms.date: 02/25/2022
ms.localizationpriority: medium
---

# MobileFormFactor element

Specifies the settings for an add-in for the mobile form factor. It contains all the add-in information for the mobile form factor except for the **Resources** node.

Each **MobileFormFactor** definition contains the **FunctionFile** element and one or more **ExtensionPoint** elements. For more information, see [FunctionFile element](functionfile.md) and [ExtensionPoint element](extensionpoint.md).

The **MobileFormFactor** element is defined in VersionOverrides schema 1.1. The containing [VersionOverrides](versionoverrides.md) element must have an `xsi:type` attribute value of `VersionOverridesV1_1`.

**Add-in type:** Mail

**Valid only in these VersionOverrides schemas**:

- Mail 1.1

For more information, see [Version overrides in the manifest](../../develop/add-in-manifests.md#version-overrides-in-the-manifest).

## Child elements

| Element                             | Required | Description  |
|:------------------------------------|:--------:|:-------------|
| [ExtensionPoint](extensionpoint.md) | Yes      | Defines where an add-in exposes functionality. |
| [FunctionFile](functionfile.md)     | Yes      | A URL to a file that contains JavaScript functions.|

## MobileFormFactor example

```xml
...
<Hosts>
  <Host xsi:type="MailHost">
    ...
    <MobileFormFactor>
      <FunctionFile resid="residUILessFunctionFileUrl" />
      <ExtensionPoint xsi:type="MobileMessageReadCommandSurface">
        <!-- information on this extension point -->
      </ExtensionPoint>
      <!-- possibly more ExtensionPoint elements -->
    </MobileFormFactor>
  </Host>
</Hosts>
...
```
