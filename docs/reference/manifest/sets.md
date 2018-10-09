# Sets element

Specifies the minimum subset of the JavaScript API for Office that your Office Add-in requires in order to activate.

**Add-in type:** Content, Task pane, Mail

## Syntax

```XML
<Sets DefaultMinVersion="n .n ">
   ...
</Sets>
```

## Contained in

[Requirements](requirements.md)

## Can contain

[Set](set.md)

## Attributes

|**Attribute**|**Type**|**Required**|**Description**|
|:-----|:-----|:-----|:-----|
|DefaultMinVersion|string|optional|Specifies the default  **MinVersion** attribute value for all child [Set](set.md) elements. The default value is "1.1".|

## Remarks

For more information about requirement sets, see [Office versions and requirement sets](https://docs.microsoft.com/office/dev/add-ins/develop/office-versions-and-requirement-sets).

For more information about the  **MinVersion** attribute of the **Set** element and the **DefaultMinVersion** attribute of the **Sets** element, see [Set the Requirements element in the manifest](https://docs.microsoft.com/office/dev/add-ins/develop/specify-office-hosts-and-api-requirements#set-the-requirements-element-in-the-manifest).

