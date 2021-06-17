---
title: SupportsSharedFolders element in the manifest file
description: The SupportsSharedFolders element defines whether the Outlook add-in is available in shared folders and shared mailbox scenarios.
ms.date: 06/15/2021
localization_priority: Normal
---

# SupportsSharedFolders element

Defines whether the Outlook add-in is available in shared mailbox (now in preview) and shared folders (that is, delegate access) scenarios. The **SupportsSharedFolders** element is a child element of [DesktopFormFactor](desktopformfactor.md). It is set to *false* by default.

> [!IMPORTANT]
> Support for this element was introduced in requirement set 1.8. See [clients and platforms](../../reference/requirement-sets/outlook-api-requirement-sets.md#requirement-sets-supported-by-exchange-servers-and-outlook-clients) that support this requirement set.

The following is an example of the **SupportsSharedFolders** element.

```XML
...
<VersionOverrides xmlns="http://schemas.microsoft.com/office/mailappversionoverrides" xsi:type="VersionOverridesV1_0">
  <VersionOverrides xmlns="http://schemas.microsoft.com/office/mailappversionoverrides/1.1" xsi:type="VersionOverridesV1_1">
    ...
    <Hosts>
      <Host xsi:type="MailHost">
        <DesktopFormFactor>
          <SupportsSharedFolders>true</SupportsSharedFolders>
          <FunctionFile resid="residDesktopFuncUrl" />
          <ExtensionPoint xsi:type="MessageReadCommandSurface">
            <!-- Configure selected extension point. -->
          </ExtensionPoint>

          <!-- You can define more than one ExtensionPoint element as needed. -->

        </DesktopFormFactor>
      </Host>
    </Hosts>
    ...
  </VersionOverrides>
</VersionOverrides>
...
```
