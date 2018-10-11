# Requirements element

Specifies the minimum set of JavaScript API for Office requirements ([requirement sets](https://docs.microsoft.com/office/dev/add-ins/develop/office-versions-and-requirement-sets#specify-office-hosts-and-requirement-sets) and/or methods) that your Office Add-in needs to activate.

**Add-in type:** Content, Task pane, Mail

## Syntax

```XML
<Requirements>
   ...
</Requirements>
```

## Contained in

[OfficeApp](officeapp.md)

## Can contain

|**Element**|**Content**|**Mail**|**TaskPane**|
|:-----|:-----|:-----|:-----|
|[Sets](sets.md)|x|x|x|
|[Methods](methods.md)|x||x|

## Remarks

For more information about requirement sets, see [Office versions and requirement sets](https://docs.microsoft.com/office/dev/add-ins/develop/office-versions-and-requirement-sets).

