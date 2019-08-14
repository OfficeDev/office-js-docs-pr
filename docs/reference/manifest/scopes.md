---
title: Scopes element in the manifest file
description: ''
ms.date: 08/12/2019
localization_priority: Normal
---

# Scopes element

Contains permissions that the add-in needs to an external resource, such as Microsoft Graph. When Microsoft Graph is the resource, AppSource uses the Scopes element to create a consent dialog box. When users install the add-in from the Store, they are prompted to grant the add-in the specified permissions to the user's Microsoft Graph data.

**Scopes** is a child element of the [WebApplicationInfo](webapplicationinfo.md) and [Authorization](authorization.md) elements in the manifest.

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
      <Resource>api://myDomain.com/12345678-abcd-1234-efab-123456789abc<Resource>
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
