---
title: ExtendedPermissions element in the manifest file
description: ''
ms.date: 03/04/2020
localization_priority: Normal
---

# ExtendedPermissions element

Defines the extended permissions the add-in needs to work. The `ExtendedPermissions` element is a child element of [VersionOverrides](versionoverrides.md).

> [!IMPORTANT]
> This element is only available in the [Outlook add-ins preview requirement set](../objectmodel/preview-requirement-set/outlook-requirement-set-preview.md) against Exchange Online. Add-ins that use this element cannot be published to AppSource or deployed via centralized deployment.

## Child elements

|  Element |  Required  |  Description  |
|:-----|:-----|:-----|
|  [ExtendedPermission](extendedpermission.md)    |  Yes   | Defines the extended permission needed for the add-in to work. At least one entry should be included in the `ExtendedPermissions` collection. |

## `ExtendedPermissions` example

The following is an example of the `ExtendedPermissions` element.

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
    <ExtendedPermissions>
      <ExtendedPermission>AppendOnSend</ExtendedPermission>
    </ExtendedPermissions>
  </VersionOverrides>
</VersionOverrides>
...
```

## Contained in

[VersionOverrides](versionoverrides.md)

## Can contain

[ExtendedPermission](extendedpermission.md)
