# WebApplicationInfo element

Supports single sign-on (SSO) in Office Add-ins. This element contains information for the add-in as both:

- An OAuth 2.0 *resource* to which the Office host application might need permissions.
- An OAuth 2.0 *client* that might need permissions to Microsoft Graph.

**WebApplicationInfo** is a child element of the [VersionOverrides](versionoverrides.md) element in the manifest.  

## Child elements

|  Element |  Required  |  Description  |
|:-----|:-----|:-----|
|  **Id**    |  Yes   |  The **Application Id** of the add-in's associated service as registered in the Azure Active Directory v 2.0 endpoint.|
|  **Resource**  |  Yes   |  Specifies the **Application ID URI** of the add-in as registered in the Azure Active Directory v 2.0 endpoint.|
|  [Scopes](scopes.md)                |  No  |  Specifies the permissions that the add-in needs to Microsoft Graph.  |

> [!NOTE] 
> Currently, it's necessary that your add-in's Resource matches its Host. Office will not request a Token for an add-in unless it can prove ownership, and today this is done by hosting the add-in under the Resource's fully-qualified domain name.

## WebApplicationInfo example

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
