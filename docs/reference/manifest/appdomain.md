---
title: AppDomain element in the manifest file
description: Specifies additional domains that are used by your add-in and should be trusted by Office.
ms.date: 06/12/2020
ms.localizationpriority: medium
---

# AppDomain element

Specifies an additional domain that Office should trust, in addition to the one specified in the [SourceLocation element](sourcelocation.md). Specifying a domain has these effects:

- It enables pages, routes, or other resources in the domain to be opened directly in the root task pane of the add-in on desktop Office platforms. (Specifying a domain in an **AppDomain** isn't necessary for Office on the web or to open a resource in an IFrame, nor it is necessary for opening a resource in a dialog opened with the [Dialog API](../../develop/dialog-api-in-office-add-ins.md).)
- It enables pages in the domain to make Office.js API calls from IFrames within the add-in.

**Add-in type:** Content, Task pane, Mail

## Syntax

```XML
<AppDomain>string</AppDomain>
```

> [!IMPORTANT]
> 1. The value of the **AppDomain** element must include the protocol (e.g., `<AppDomain>https://myappdomain.com</AppDomain>`).
> 2. If there is an explicit port for the domain, include it (e.g.,`<AppDomain>https://myappdomain.com:9999</AppDomain>`).
> 3. If a subdomain needs to be trusted, include it (e.g.,`<AppDomain>https://mysubdomain.myappdomain.com</AppDomain>`). The subdomain `mysubdomain.mydomain.com` and `mydomain.com` are different domains. If both need to be trusted, then both need to be in separate **AppDomain** elements.
> 4. Listing the same domain as the one specified in the [SourceLocation element](sourcelocation.md) has no effect and may be misleading. In particular, when you are developing on `localhost`, you don't need to create an **AppDomain** element for `localhost`.
> 5. Don't include any segments of a URL past the domain. For example, don't include the full URL of a page.
> 6. Do *not* put a closing slash, "/", on the value.

## Contained in

[AppDomains](appdomains.md)

## Remarks

For more information, see [Office Add-ins XML manifest](../../develop/add-in-manifests.md).
