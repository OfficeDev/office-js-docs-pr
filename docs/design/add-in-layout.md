---
title: Layout guidelines for Office Add-ins
description: Learn how to design consistent, responsive layouts for task panes, content areas, and dialogs in Office Add-ins, including margin, grid, and alignment best practices.
ms.date: 04/06/2026
ms.topic: best-practice
ms.localizationpriority: medium
ai-usage: ai-assisted
---

# Layout guidelines for Office Add-ins

A well-designed layout helps users navigate your add-in quickly and complete tasks with confidence. Every HTML container embedded in Office — whether a task pane, content area, or dialog — represents a screen in your add-in, and consistent layouts across those screens create a seamless experience.

If you already have a website your customers know, consider reusing its layouts. Adapt them to fit within Office HTML containers so the add-in feels like a natural part of the Office environment.

For detailed guidance on specific container types, see [Task pane add-ins](task-pane-add-ins.md) and [Content add-ins](content-add-ins.md). To learn how to assemble [Fluent UI React](../quickstarts/fluent-react-quickstart.md) or [Fluent UI (Fabric Core)](fabric-core.md) components into common layouts and user experience flows, see [UX design pattern templates](ux-design-pattern-templates.md).

## General layout guidelines

Follow these guidelines when you design your add-in layouts.

### Margins and spacing

- Use 20 pixels as the default margin for HTML containers. Avoid margins that are too narrow or too wide.
- Keep padding between elements at multiples of 4 pixels. Office interfaces use a 4 px grid.

### Alignment and visual hierarchy

- Align elements intentionally. Extra indents and new alignment points should reinforce the visual hierarchy, not create clutter.
- Follow common layout patterns. Conventions help users understand how to use an interface without extra effort.

### Responsiveness and consistency

- Keep layouts consistent across screens. Unexpected changes look like visual bugs and reduce user trust.
- Create responsive experiences that adapt to the width and height of the HTML container.

### Simplicity

- Avoid overcrowding the interface. Dense layouts cause confusion and make touch interactions harder.
- Avoid redundant elements like duplicate branding or commands.
- Consolidate controls and views to minimize unnecessary mouse movement.

## See also

- [Task pane add-ins](task-pane-add-ins.md)
- [Content add-ins](content-add-ins.md)
- [UX design pattern templates](ux-design-pattern-templates.md)
- [Office Add-in design language](add-in-design-language.md)
