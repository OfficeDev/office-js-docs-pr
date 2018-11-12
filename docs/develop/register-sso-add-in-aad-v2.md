---
title: Register an Office Add-in that uses SSO with the Azure AD v2.0 endpoint
description: ''
ms.date: 04/10/2018 
---

# Register an Office Add-in that uses SSO with the Azure AD v2.0 endpoint

This article explains how to register an Office Add-in with the Azure AD v2.0 endpoint. You need to register the add-in when you begin developing it. When you progress to testing or production, you can change the existing registration or create separate registrations for development, testing, and production versions of the add-in.

The following table itemizes the information that you need to carry out this procedure and the corresponding placeholders that appear in the instructions. 

|Information  |Examples  |Placeholder  |
|---------|---------|---------|
|A human readable name for the add-in. (Uniqueness recommended, but not required.)    |`Contoso Marketing Excel Add-in (Prod)`        |**$ADD-IN-NAME$**         |
|The fully qualified domain name (except for protocol) of the add-in. *You must use a domain that you own.* For this reason, you cannot use certain well-known domains such as `azurewebsites.net` or `cloudapp.net`. The domain must be the same, including any subdomains, as is used in the URLs in the `<Resources>` section of the add-in's manifest.  |`localhost:6789`, `addins.contoso.com`         |**$FQDN-WITHOUT-PROTOCOL$**         |
|The permissions to AAD and Microsoft Graph that your add-in needs. (`profile` is always required.)    |`profile`, `Files.Read.All`         |N/A         |

[!INCLUDE[](../includes/register-sso-add-in-aad-v2-include.md)]