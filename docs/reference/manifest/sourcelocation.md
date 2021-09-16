---
title: SourceLocation element in the manifest file
description: The SourceLocation element specifies the source file locations for your Office Add-in.
ms.date: 05/12/2021
ms.localizationpriority: medium
---

# SourceLocation element

Specifies the source file locations for your Office Add-in as a URL between 1 and 2018 characters long. The source location must be an HTTPS address, not a file path.

**Add-in type:** Content, Task pane, Mail

## Syntax

```XML
<SourceLocation DefaultValue="string" />
```

## Contained in

- [DefaultSettings](defaultsettings.md) (Content and task pane add-ins)
- [FormSettings](formsettings.md) (Mail add-ins)
- [ExtensionPoint](extensionpoint.md) (Contextual and LaunchEvent mail add-ins)

## Can contain

[Override](override.md)

## Attributes

|Attribute|Type|Required|Description|
|:-----|:-----|:-----|:-----|
|DefaultValue|URL|required|Specifies the default value for this setting for the locale specified in the [DefaultLocale](defaultlocale.md) element.|
