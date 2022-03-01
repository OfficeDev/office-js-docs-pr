---
title: WebApplicationInfo element in the manifest file
description: Reference documentation of the WebApplicationInfo element for Office Add-ins manifest (XML) files.
ms.date: 02/25/2022
ms.localizationpriority: medium
---

# WebApplicationInfo element

Supports single sign-on (SSO) in Office Add-ins. This element contains information for the add-in as both:

- An OAuth 2.0 *resource* to which the Office client application might need permissions.
- An OAuth 2.0 *client* that might need permissions to Microsoft Graph.

**Add-in type:** Task pane, Mail, Content

**Valid only in these VersionOverrides schemas**:

- Task pane 1.0
- Content 1.0
- Mail 1.0
- Mail 1.1

For more information, see [Version overrides in the manifest](../../develop/add-in-manifests.md#version-overrides-in-the-manifest).

**Associated with these requirement sets**:

- [IdentityAPI 1.3](../requirement-sets/identity-api-requirement-sets.md)

> [!NOTE]
> The single sign-on API is currently supported for Word, Excel, Outlook, and PowerPoint. For more information about where the single sign-on API is currently supported, see [IdentityAPI requirement sets](../requirement-sets/identity-api-requirement-sets.md). If you are working with an Outlook add-in, be sure to enable Modern Authentication for the Microsoft 365 tenancy. To learn how to do this, see [Exchange Online: How to enable your tenant for modern authentication](https://social.technet.microsoft.com/wiki/contents/articles/32711.exchange-online-how-to-enable-your-tenant-for-modern-authentication.aspx).

**WebApplicationInfo** is a child element of the [VersionOverrides](versionoverrides.md) element in the manifest.  

## Child elements

|  Element |  Required  |  Description  |
|:-----|:-----|:-----|
|  **Id**    |  Yes   |  The **Application Id** of the add-in's associated service as registered in the Azure Active Directory v 2.0 endpoint.|
|  **Resource**  |  Yes   |  Specifies the **Application ID URI** of the add-in as registered in the Azure Active Directory v 2.0 endpoint.|
|  [Scopes](scopes.md)                |  Yes  |  Specifies the permissions that the add-in needs to a resource, such as Microsoft Graph.  |

## WebApplicationInfo example

```xml
<OfficeApp>
...
  <VersionOverrides xmlns="http://schemas.microsoft.com/office/mailappversionoverrides" xsi:type="VersionOverridesV1_0">
    ...
    <WebApplicationInfo>
      <Id>12345678-abcd-1234-efab-123456789abc</Id>
      <Resource>api://contoso.com/12345678-abcd-1234-efab-123456789abc</Resource>
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
