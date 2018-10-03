# Scopes element

Contains permissions to Microsoft Graph that the add-in needs. The Office Store uses the Scopes element to create a consent dialog box. When users install the add-in from the Store, they are prompted to grant the add-in the specified permissions to the user's Microsoft Graph data.

## Child elements

|  Element |  Type  |  Description  |
|:-----|:-----|:-----|
|  **Scope**                |  string     |   The name of a permission to Microsoft Graph; for example, Files.Read.All. |

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
