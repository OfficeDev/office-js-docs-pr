---
title: Wildcard trusted domains
description: Learn how to use wildcards to specify trusted domains that aren't listed in the manifest.
ms.date: 11/14/2023
ms.localizationpriority: medium
---

# Wildcard trusted domains

Besides its own domain, an add-in can access resources in certain other domains such as authentication points for major identity providers and in any domain listed in the manifest. The latter domains are specified in the [AppDomains](/javascript/api/manifest/appdomains) element of the add-in only manifest or the [`"validDomains`"](/microsoft-365/extensibility/schema/root#validdomains) property of the unified manifest. Wildcards aren't allowed in the add-in only manifest. They are allowed in the unified manifest because some Teams apps and other Microsoft 365 apps honor them; but Office Add-ins don't honor `"validDomains"` that contain wildcards.

Windows administrators can make Office Add-ins, *running on Windows only*, honor domains that include a wildcard by setting the **HKEY_CURRENT_USER\SOFTWARE\Microsoft\Office\16.0\WEF\AllowedAppDomains** registry key with the domain. The following is an example.

```
[HKEY_CURRENT_USER\SOFTWARE\Microsoft\Office\16.0\WEF\AllowedAppDomains]
"AppDomain1"="https://*.contoso.com" 
```

Administrators can use a *.reg file to do automate the process. The following is an example of such a file.

```
Windows Registry Editor Version 5.00

[HKEY_CURRENT_USER\SOFTWARE\Microsoft\Office\16.0\WEF\AllowedAppDomains]
"AppDomain1"="https://*.europe.contoso.com" 

[HKEY_CURRENT_USER\SOFTWARE\Microsoft\Office\16.0\WEF\AllowedAppDomains]
"AppDomain2"="https://*.africa.contoso.com" 
```

> [!NOTE]
>
> - The domains are honored only in add-ins running on Windows desktop versions of Office. They aren't honored when an add-in is running in Office on the web even on computers where the registry change has been made. 
> - The registry setting affects *all* add-ins running on the computer: they all trust the domains in the registry key.
