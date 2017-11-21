
# Requirements element
Specifies the minimum set of JavaScript API for Office requirements ([requirement sets](../../docs/overview/specify-office-hosts-and-api-requirements.md#SpecifyRequirementSets_sets) and/or methods) that your Office Add-in needs to activate.

 **Add-in type:** Content, Task pane, Mail


## Syntax:


```XML
<Requirements>
   ...
</Requirements>
```


## Contained in:

[OfficeApp](https://dev.office.com/reference/add-ins/manifest/officeapp)


## Can contain:



|**Element**|**Content**|**Mail**|**TaskPane**|
|:-----|:-----|:-----|:-----|
|[Sets](https://dev.office.com/reference/add-ins/manifest/sets)|x|x|x|
|[Methods](https://dev.office.com/reference/add-ins/manifest/methods)|x||x|

## Remarks

For more information about requirement sets, see [Specify Office hosts and API requirements](../../docs/overview/specify-office-hosts-and-api-requirements.md).

