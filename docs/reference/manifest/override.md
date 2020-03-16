---
title: Override element in the manifest file
description: 'the Override element enables you to specify the value of a setting for an additional locale.'
ms.date: 03/19/2019
localization_priority: Normal
---

# Override element

Provides a way to specify the value of a setting for an additional locale.

**Add-in type:** Content, Task pane, Mail

## Syntax

```XML
<Override Locale="string" Value="string" />
```

## Contained in

|**Element**|
|:-----|
|[CitationText](citationtext.md)|
|[Description](description.md)|
|[DictionaryName](dictionaryname.md)|
|[DictionaryHomePage](dictionaryhomepage.md)|
|[DisplayName](displayname.md)|
|[HighResolutionIconUrl](highresolutioniconurl.md)|
|[IconUrl](iconurl.md)|
|[QueryUri](queryuri.md)|
|[SourceLocation](sourcelocation.md)|
|[SupportUrl](supporturl.md)|

## Attributes

|**Attribute**|**Type**|**Required**|**Description**|
|:-----|:-----|:-----|:-----|
|Locale|string|required|Specifies the culture name of the locale for this override in the BCP 47 language tag format, such as  `"en-US"`.|
|Value|string|required|Specifies value of the setting expressed for the specified locale.|

## See also

- [Localization for Office Add-ins](../../develop/localization.md)
