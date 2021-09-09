---
title: AppDomains element in the manifest file
description: Lists any domains in addition to the domain specified in the `SourceLocation` element that your Office Add-in will use and should be trusted by Office.
ms.date: 06/12/2020
ms.localizationpriority: medium
---

# AppDomains element

Lists any domains, in addition to the domain specified in the `SourceLocation` element, that your Office Add-in will use and that should be trusted by Office. This enables pages in the domains to make calls to Office.js APIs from IFrames within the add-in and has other effects. For each additional domain, specify an **AppDomain** element.

 **Add-in type:** Content, Task pane, Mail

## Syntax

```XML
<AppDomains>
    <AppDomain>AppDomain1</AppDomain>
    <AppDomain>AppDomain2</AppDomain>
</AppDomains>
```

> [!IMPORTANT]
> There are restrictions on what can be the value of a **AppDomain** element. For more information, see [AppDomain](appdomain.md).

## Contained in

[OfficeApp](officeapp.md)

## Can contain

[AppDomain](appdomain.md)

## Remarks

By default, your add-in can load any page that is in the same domain as the location specified in the [SourceLocation element](sourcelocation.md). This element can't be empty.
