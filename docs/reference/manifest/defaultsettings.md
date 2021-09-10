---
title: DefaultSettings element in the manifest file
description: Specifies the default source location and other default settings for your content or task pane add-in.
ms.date: 10/09/2018
ms.localizationpriority: medium
---

# DefaultSettings element

Specifies the default source location and other default settings for your content or task pane add-in.

**Add-in type:** Content, Task pane

## Syntax

```XML
<DefaultSettings>
  ...
</DefaultSettings>
```

## Contained in

[OfficeApp](officeapp.md)

## Can contain

|Element|Content|Mail|TaskPane|
|:-----|:-----|:-----|:-----|
|[SourceLocation](sourcelocation.md)|x||x|
|[RequestedWidth](requestedwidth.md)|x|||
|[RequestedHeight](requestedheight.md)|x|||

## Remarks

The source location and other settings in the **DefaultSettings** element apply only to content and task pane add-ins. For mail add-ins, you specify the default locations for source files and other default settings in the [FormSettings](formsettings.md) element.
