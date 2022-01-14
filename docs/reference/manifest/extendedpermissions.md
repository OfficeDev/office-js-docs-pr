---
title: ExtendedPermissions element in the manifest file
description: Defines the collection of extended permissions the add-in needs to access associated APIs or features.
ms.date: 01/04/2022
ms.localizationpriority: medium
---

# ExtendedPermissions element

Defines the collection of extended permissions the add-in needs to access associated APIs or features. The `ExtendedPermissions` element is a child element of [VersionOverrides](versionoverrides.md).

> [!IMPORTANT]
> Support for this element was introduced in requirement set 1.9. See [clients and platforms](../../reference/requirement-sets/outlook-api-requirement-sets.md#requirement-sets-supported-by-exchange-servers-and-outlook-clients) that support this requirement set.

**Add-in type:** Mail

**Valid only in these VersionOverrides schemas**:

- Mail 1.1

For more information, see [Version overrides in the manifest](../../develop/add-in-manifests.md#version-overrides-in-the-manifest).

**Associated with these requirement sets**:

- [Mailbox 1.9](../../reference/objectmodel/requirement-set-1.9/outlook-requirement-set-1.9.md)

## Child elements

|  Element |  Required  |  Description  |
|:-----|:-----:|:-----|
|  [ExtendedPermission](extendedpermission.md)    |  No   | Defines an extended permission needed for the add-in to access the associated API or feature. |

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
