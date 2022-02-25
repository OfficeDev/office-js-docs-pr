---
title: Scopes element in the manifest file
description: The Scopes element contains permissions the add-in needs to connect to an external resource.
ms.date: 02/25/2022
ms.localizationpriority: medium
---

# Scopes element

Contains permissions that the add-in needs to an external resource, such as Microsoft Graph. When Microsoft Graph is the resource, AppSource uses the Scopes element to create a consent dialog box. When users install the add-in from the Store, they are prompted to grant the add-in the specified permissions to the user's Microsoft Graph data.

**Add-in type:** Task pane, Mail, Content

**Valid only in these VersionOverrides schemas**:

- Task pane 1.0
- Content 1.0
- Mail 1.0
- Mail 1.1

For more information, see [Version overrides in the manifest](../../develop/add-in-manifests.md#version-overrides-in-the-manifest).

**Associated with these requirement sets**:

- [IdentityAPI 1.3](../requirement-sets/identity-api-requirement-sets.md)

**Scopes** is a child element of the [WebApplicationInfo](webapplicationinfo.md) element in the manifest.

## Child elements

|  Element |  Required  |  Description  |
|:-----|:-----|:-----|
|  **Scope**                |  Yes     |   The name of a permission; for example, Files.Read.All or profile. |

## Example

```xml
<OfficeApp>
...
  <VersionOverrides xmlns="http://schemas.microsoft.com/office/mailappversionoverrides" xsi:type="VersionOverridesV1_0">
    ...
    <WebApplicationInfo>
      <Id>12345678-abcd-1234-efab-123456789abc</Id>
      <Resource>api://contoso.com/12345678-abcd-1234-efab-123456789abc<Resource>
      <Scopes>
        <Scope>Files.Read.All</Scope>
        <Scope>offline_access</Scope>
        <Scope>openid</Scope>
        <Scope>profile</Scope>
      </Scopes>
    </WebApplicationInfo>
  </VersionOverrides>
...
</OfficeApp>
```
