---
title: WebApplicationInfo element in the manifest file
description: 'Reference documentation of the VersionOverrides element for Office Add-ins manifest (XML) files.'
ms.date: 08/12/2019
localization_priority: Normal
---

# WebApplicationInfo element

Supports single sign-on (SSO) in Office Add-ins. This element contains information for the add-in as both:

- An OAuth 2.0 *resource* to which the Office host application might need permissions.
- An OAuth 2.0 *client* that might need permissions to Microsoft Graph.

> [!NOTE]
> The single sign-on API is currently supported in preview for Word, Excel, Outlook, and PowerPoint. For more information about where the single sign-on API is currently supported, see [IdentityAPI requirement sets](../requirement-sets/identity-api-requirement-sets.md). If you are working with an Outlook add-in, be sure to enable Modern Authentication for the Office 365 tenancy. To learn how to do this, see [Exchange Online: How to enable your tenant for modern authentication](https://social.technet.microsoft.com/wiki/contents/articles/32711.exchange-online-how-to-enable-your-tenant-for-modern-authentication.aspx).

**WebApplicationInfo** is a child element of the [VersionOverrides](versionoverrides.md) element in the manifest.  

## Child elements

|  Element |  Required  |  Description  |
|:-----|:-----|:-----|
|  **Id**    |  Yes   |  The **Application Id** of the add-in's associated service as registered in the Azure Active Directory v 2.0 endpoint.|
|  **MsaId**    |  No   |  The client ID of your add-in's web application for MSA as registered in msm.live.com.|
|  **Resource**  |  Yes   |  Specifies the **Application ID URI** of the add-in as registered in the Azure Active Directory v 2.0 endpoint.|
|  [Scopes](scopes.md)                |  Yes  |  Specifies the permissions that the add-in needs to a resource, such as Microsoft Graph.  |
|  [Authorizations](authorizations.md)  |  No   | Specifies the external resources that the add-in's web application needs authorization to and the required permissions.|

## WebApplicationInfo example

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
