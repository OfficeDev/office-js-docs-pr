---
title: ExtendedPermission element in the manifest file
description: Defines an extended permission the add-in needs to access the associated API or feature.
ms.date: 10/15/2020
localization_priority: Normal
---

# `ExtendedPermission` element

Defines an extended permission the add-in needs to access the associated API or feature. The `ExtendedPermission` element is a child element of [ExtendedPermissions](extendedpermissions.md).

> [!IMPORTANT]
> Support for this element was introduced in requirement set 1.9. See [clients and platforms](../../reference/requirement-sets/outlook-api-requirement-sets.md#requirement-sets-supported-by-exchange-servers-and-outlook-clients) that support this requirement set.

## Available extended permissions

The following are the available values.

|Available value|Description|Hosts|
|---|---|---|
|`AppendOnSend`|Declares that the add-in is using the [Office.Body.appendOnSendAsync](/javascript/api/outlook/office.body?view=outlook-js-preview&preserve-view=true#appendonsendasync-data--options--callback-) API.|Outlook|

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
