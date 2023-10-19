---
title: Layout guidelines for Office Add-ins
description: Get guidelines for how to layout a task pane or dialog in an Office Add-in.
ms.date: 08/18/2023
ms.topic: best-practice
ms.localizationpriority: medium
---

# Layout

Each HTML container embedded in Office will have a layout. These layouts are the main screens of your add-in. In them you'll create experiences that enable customers to initiate actions, modify settings, view, scroll, or navigate content. Design your add-in with a consistent layouts across screens to guarantee continuity of experience. If you have an existing website that your customers are familiar with using, consider reusing layouts from your existing web pages. Adapt them to fit harmoniously within Office HTML containers.

For guidelines on layout, see [Task pane](task-pane-add-ins.md), [Content](content-add-ins.md). For more information about how to assemble [Fluent UI React](../quickstarts/fluent-react-quickstart.md), or [Office UI Fabric JS](fabric-core.md), components into common layouts and user experience flows, see [UX design patterns templates](ux-design-pattern-templates.md).

Apply the following general guidelines for layouts.

- Avoid narrow or wide margins on your HTML containers. 20 pixels is a great default.
- Align elements intentionally. Extra indents and new points of alignment should aid visual hierarchy.
- Office interfaces are on a 4px grid. Aim to keep your padding between elements at multiples of 4.
- Overcrowding your interface can lead to confusion and inhibit ease of use with touch interactions.
- Keep layouts consistent across screens. Unexpected layout changes look like visual bugs that contribute to a lack of confidence and trust with your solution.
- Follow common layout patterns. Conventions help users understand how to use an interface.
- Avoid redundant elements like branding or commands.
- Consolidate controls and views to avoid requiring too much mouse movement.
- Create responsive experiences that adapt to HTML container widths and heights.
