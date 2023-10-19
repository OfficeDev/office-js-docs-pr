---
title: Typography guidelines for Office Add-ins
description: Learn what typefaces and font sizes to use in Office Add-ins.
ms.date: 08/18/2023
ms.topic: best-practice
ms.localizationpriority: medium
---

# Typography

Segoe is the standard typeface for Office. Use it in your add-in to align with Office task panes, dialog boxes, and content objects. [Fabric Core](fabric-core.md) gives you access to Segoe. It provides a full type ramp of Segoe with many variations - across font weight and size - in convenient CSS classes. Not all Fabric Core sizes and weights will look great in an Office Add-in. To fit harmoniously or avoid conflicts, consider using a subset of the Fabric Core type ramp. The following table lists Fabric Core's base classes that we recommend for use in Office Add-ins.

> [!NOTE]
> Text color isn't included in these base classes. Use Fabric Core's "neutral primary" for most text on white backgrounds.
>
> To learn more about available typography, see [Web Typography](https://developer.microsoft.com/fluentui#/styles/web/typography).

|Type |Class |Size |Weight |Recommended Usage |
|------ |----- |---- |------ |----------------- |
|Hero|.ms-font-xxl |28 px | Segoe Light |<ul><li>This class is larger than all other typographic elements in Office. Use it sparingly to avoid unseating visual hierarchy.</li><li>Avoid use on long strings in constrained spaces.</li><li>Provide ample whitespace around text using this class.</li><li>Commonly used for first-run messages, hero elements, or other calls to action.</li></ul> |
|Title|.ms-font-xl |21 px |Segoe Light | <ul><li>This class matches the task pane title of Office applications.</li><li>Use it sparingly to avoid a flat typographic hierarchy.</li><li>Commonly used as the top-level element such as dialog box, page, or content titles.</li></ul> |
|Subtitle|.ms-font-l |17 px |Segoe Semilight | <ul><li>This class is the first stop below titles.</li><li>Commonly used as a subtitle, navigation element, or group header.</li><ul> |
|Body|.ms-font-m |14 px |Segoe Regular |<ul><li>Commonly used as body text within add-ins.</li><ul>|
|Caption|.ms-font-xs |11 px | Segoe Regular |<ul><li>Commonly used for secondary or tertiary text such as timestamps, by lines, captions, or field labels.</li><ul>|
|Annotation|.ms-font-mi |10 px |Segoe Semibold |<ul><li>The smallest step in the type ramp should be used rarely. It's available for circumstances where legibility isn't required.</li><ul>|
