
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

[OfficeApp](../../reference/manifest/officeapp.md)


## Can contain:



|**Element**|**Content**|**Mail**|**TaskPane**|
|:-----|:-----|:-----|:-----|
|[Sets](../../reference/manifest/sets.md)|x|x|x|
|[Methods](../../reference/manifest/methods.md)|x||x|

## Remarks

For more information about requirement sets, see [Specify Office hosts and API requirements](../../docs/overview/specify-office-hosts-and-api-requirements.md).

