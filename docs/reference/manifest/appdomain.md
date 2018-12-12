# AppDomain element

Specifies an additional domain that will be used to load pages in the add-in window.

**Add-in type:** Content, Task pane, Mail

## Syntax

```XML
<AppDomain>string</AppDomain>
```

> [!IMPORTANT]
> The **AppDomain** must include the protocol e.g., `<AppDomain>https://myappdomain"<AppDomain>`.

## Contained in

[AppDomains](appdomains.md)

## Remarks

The  **AppDomains** and **AppDomain** elements are used to specify any additional domains other than the one specified in the [SourceLocation element](sourcelocation.md). For more information, see [Office Add-ins XML manifest](/office/dev/add-ins/develop/add-in-manifests).
