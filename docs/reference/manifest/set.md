# Set element

Specifies a requirement set from the JavaScript API for Office that your Office Add-in requires to activate.

**Add-in type:** Content, Task pane, Mail

## Syntax

```XML
<Set Name="string" MinVersion="n .n">
```

## Contained in

[Sets](sets.md)

## Attributes

|**Attribute**|**Type**|**Required**|**Description**|
|:-----|:-----|:-----|:-----|
|Name|string|required|The name of a [requirement set](https://docs.microsoft.com/office/dev/add-ins/develop/office-versions-and-requirement-sets).|
|MinVersion|string|optional|Specifies the minimum version of the API set required by your add-in. Overrides the value of  **DefaultMinVersion**, if it is specified in the parent [Sets](sets.md) element.|

## Remarks

For more information about requirement sets, see [Office versions and requirement sets](https://docs.microsoft.com/office/dev/add-ins/develop/office-versions-and-requirement-sets).

For more information about the  **MinVersion** attribute of the **Set** element and the **DefaultMinVersion** attribute of the **Sets** element, see [Set the Requirements element in the manifest](https://docs.microsoft.com/office/dev/add-ins/develop/specify-office-hosts-and-api-requirements#set-the-requirements-element-in-the-manifest).

> [!IMPORTANT] 
> For mail add-ins, there is only one  `"Mailbox"` requirement set available. This requirement set contains the entire subset of API supported in mail add-ins for Outlook, and you must specify the `"Mailbox"` requirement set in your mail add-in's manifest (it's not optional as is the case for content and task pane add-ins). Also, you can't declare support for specific methods in mail add-ins.
