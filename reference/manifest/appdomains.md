
# AppDomains element
Lists any domains in addition to the domain specified in the [SourceLocation](../../reference/manifest/sourcelocation.md) element that your Office Add-in will use to load pages. For each additional domain, specify an [AppDomain](../../reference/manifest/appdomain.md) element.

 **Add-in type:** Content, Task pane, Mail


## Syntax:


```XML
<AppDomains>
    <AppDomain>AppDomain1</AppDomain>
    <AppDomain>AppDomain2</AppDomain>
</AppDomains>
```


## Contained in:

[OfficeApp](../../reference/manifest/officeapp.md)


## Can contain:

[AppDomain](../../reference/manifest/appdomain.md)


## Remarks

By default, your add-in can load any page that is in the same domain as the location specified in the [SourceLocation](../../reference/manifest/sourcelocation.md) element. To load pages that are not in the same domain as the add-in, specify the domains by using the **AppDomains** and **AppDomain** elements. This element can't be empty. 

For more information, see [Office Add-ins XML manifest](../../docs/overview/add-in-manifests.md).

