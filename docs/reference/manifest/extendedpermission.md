---
title: ExtendedPermission element in the manifest file
description: Defines an extended permission the add-in needs to access the associated API or feature.
ms.date: 03/05/2020
localization_priority: Normal
---

# `ExtendedPermission` element

Defines an extended permission the add-in needs to access the associated API or feature. The `ExtendedPermission` element is a child element of [ExtendedPermissions](extendedpermissions.md).

> [!IMPORTANT]
> This element is only available in the [Outlook add-ins preview requirement set](../objectmodel/preview-requirement-set/outlook-requirement-set-preview.md) against Exchange Online. Add-ins that use this element cannot be published to AppSource or deployed via centralized deployment.

## Available extended permissions

The following are the available values.

|Available value|Description|Hosts|
|---|---|---|
|`AppendOnSend`|Declares that the add-in is using the [Office.Body.appendOnSendAsync](/javascript/api/outlook/office.body?view=outlook-js-preview#appendonsendasync-data--options--callback-) API.|Outlook|

## `ExtendedPermission` example

The following is an example of the `ExtendedPermission` element.

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

[ExtendedPermissions](extendedpermissions.md)
