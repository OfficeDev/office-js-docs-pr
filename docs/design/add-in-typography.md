---
title: Typography guidelines for Office Add-ins
description: Learn what typefaces and font sizes to use in Office Add-ins.
ms.date: 01/06/2026
ms.topic: best-practice
ms.localizationpriority: medium
---

# Typography

Segoe is the standard typeface for Office. Use it in your add-in to align with Office task panes, dialog boxes, and content objects. [Fluent UI React](../quickstarts/fluent-react-quickstart.md) and [Fabric Core](fabric-core.md) give access to the Segoe typeface and its variations on font weights and sizes.

## Typography in Fluent UI React

For guidance on font weights, sizes, and customization provided by Fluent UI React, see [Fluent UI React Typography](https://developer.microsoft.com/fluentui#/styles/web/typography).

## Typography in Fabric Core

The following table lists Fabric Core's base classes that we recommend for use in Office Add-ins.

> [!NOTE]
> Text color isn't included in these base classes. Use Fabric Core's "neutral primary" for most text on white backgrounds.

|Type|Class|Size|Weight|Recommended Usage|
|----|-----|----|------|-------------|
|Hero|.ms-font-xxl|28 px|Segoe Light|<ul><li>This class is larger than all other typographic elements in Office. Use it sparingly to avoid unseating visual hierarchy.</li><li>Avoid use on long strings in constrained spaces.</li><li>Provide ample whitespace around text using this class.</li><li>Commonly used for first-run messages, hero elements, or other calls to action.</li></ul>|
|Title|.ms-font-xl|21 px|Segoe Light|<ul><li>This class matches the task pane title of Office applications.</li><li>Use it sparingly to avoid a flat typographic hierarchy.</li><li>Commonly used as the top-level element such as dialog box, page, or content titles.</li></ul>|
|Subtitle|.ms-font-l|17 px|Segoe Semilight|<ul><li>This class is the first stop below titles.</li><li>Commonly used as a subtitle, navigation element, or group header.</li></ul>|
|Body|.ms-font-m|14 px|Segoe Regular|<ul><li>Commonly used as body text within add-ins.</li></ul>|
|Caption|.ms-font-xs|11 px|Segoe Regular|<ul><li>Commonly used for secondary or tertiary text such as timestamps, by lines, captions, or field labels.</li></ul>|
|Annotation|.ms-font-mi|10 px|Segoe Semibold|<ul><li>The smallest step in the type ramp should be used rarely. It's available for circumstances where legibility isn't required.</li></ul>|

## See also

- [Color guidelines for Office Add-ins](add-in-color.md)
- [Branding patterns](branding-patterns.md)
- [Layout](add-in-layout.md)
