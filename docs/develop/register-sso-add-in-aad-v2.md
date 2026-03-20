---
title: Register an Office Add-in that uses legacy Office SSO with the Microsoft identity platform
description: Learn how to register an Office Add-in with the Microsoft identity platform to use legacy Office SSO with Word, Excel, PowerPoint, and Outlook.
ms.date: 05/25/2025
ms.localizationpriority: medium
---

# Register an Office Add-in that uses legacy Office SSO with the Microsoft identity platform

> [!NOTE]
> This article describes legacy Office single sign-on (SSO). For a modern authentication experience with support across a wider range of platforms, use the Microsoft Authentication Library (MSAL) with nested app authentication (NAA). For more information, see [Enable single sign-on in an Office Add-in with nested app authentication](enable-nested-app-authentication-in-your-add-in.md).

This article explains how to register an Office Add-in with the Microsoft identity platform so that you can use legacy Office SSO. Register the add-in when you begin developing it so that when you progress to testing or production, you can change the existing registration or create separate registrations for development, testing, and production versions of the add-in.

The following table itemizes the information that you need to carry out this procedure and the corresponding placeholders that appear in the instructions.

|Information  |Examples  |Placeholder  |
|---------|---------|---------|
|A human readable name for the add-in. (Uniqueness recommended, but not required.)|`Contoso Marketing Excel Add-in (Prod)`| `<add-in-name>` |
|An application ID which Azure generates for you as part of the registration process.|`c6c1f32b-5e55-4997-881a-753cc1d563b7`|`<app-id>`|
|The fully qualified domain name (except for protocol) of the add-in. *You must use a domain that you own.* For this reason, you cannot use certain well-known domains such as `azurewebsites.net` or `cloudapp.net`. The domain must be the same, including any subdomains, as is used in the URLs in the `<Resources>` section of the add-in's manifest.|`localhost:6789`, `addins.contoso.com`|`<fully-qualified-domain-name>`|
|The permissions to the Microsoft identity platform and Microsoft Graph that your add-in needs. (`profile` is always required.)|`profile`, `Files.Read.All`|N/A|

> [!CAUTION]
> **Sensitive information**: The application ID URI (`<fully-qualified-domain-name>`) is logged as part of the authentication process when an add-in using SSO is activated in Office running inside of Microsoft Teams. The URI mustn't contain sensitive information.

[!INCLUDE [register-sso-add-in-aad-v2-include](../includes/register-sso-add-in-aad-v2-include.md)]
