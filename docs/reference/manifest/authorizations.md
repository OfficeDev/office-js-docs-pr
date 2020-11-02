---
title: Authorizations element in the manifest file
description: Specifies the external resources that the add-in's web application needs authorization to and the required permissions.
ms.date: 08/12/2019
localization_priority: Normal
---

# Authorizations element

Specifies the external resources that the add-in's web application needs authorization to and the required permissions.

**Authorizations** is a child element of the [WebApplicationInfo](webapplicationinfo.md) element in the manifest.

## Child elements

|  Element |  Required  |  Description  |
|:-----|:-----|:-----|
|  [Authorization](authorization.md)                |  Yes     |   Identifies an external resource that the add-in's web application needs authorization to, and the scopes (permissions) that it needs. |

## Example

```xml
<OfficeApp>
...
  <VersionOverrides xmlns="http://schemas.microsoft.com/office/mailappversionoverrides" xsi:type="VersionOverridesV1_0">
    ...
    <WebApplicationInfo>
      <Id>12345678-abcd-1234-efab-123456789abc</Id>
      <Resource>api://myDomain.com/12345678-abcd-1234-efab-123456789abc</Resource>
      <Scopes>
        <Scope>Files.Read.All</Scope>
        <Scope>offline_access</Scope>
        <Scope>openid</Scope>
        <Scope>profile</Scope>
      </Scopes>
      <Authorizations>
        <Authorization>
          <Resource>https://api.contoso.com</Resource>
            <Scopes>
              <Scope>profile</Scope>
          </Scopes>
        </Authorization>
      </Authorizations>
    </WebApplicationInfo>
  </VersionOverrides>
...
</OfficeApp>
```
