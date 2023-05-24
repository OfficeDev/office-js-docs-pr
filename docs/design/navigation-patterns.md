---
title: Navigation patterns for Office Add-ins
description: Learn best practices for using command bars, tab bars, and back buttons to design the navigation of an Office Add-in.
ms.date: 05/18/2023
ms.topic: best-practice
ms.localizationpriority: medium
---

# Navigation patterns

The main features of an add-in are accessed through specific command types and limited screen area. It's important that navigation is intuitive, provides context, and allows the user to move easily throughout the add-in.

## Best practices

| Do    | Don't |
| :---- | :---- |
| Ensure the user has a clearly visible navigation option. | Don't complicate the navigation process by using non-standard UI.
| Utilize the following components as applicable to allow users to navigate through your add-in. | Don't make it difficult for the user to understand their current place or context within the add-in

## Command Bar

The CommandBar is a surface within the task pane that houses commands that operate on the content of the window, panel, or parent region it resides above. Optional features include a hamburger menu access point, search, and side commands.

![Illustration showing a command bar within an Office desktop application task pane. This example shows a command bar immediately below the add-in name that includes a hamburger menu and search.](../images/add-in-command-bar.png)

## Tab Bar

The tab bar shows navigation using buttons with vertically stacked text and icons. Use the tab bar to provide navigation using tabs with short and descriptive titles.

![Illustration showing a tab bar within an Office desktop application task pane. This example shows a tab bar immediately below the add-in name with "Home", "Settings", "Favorites", and "Account" tabs.](../images/add-in-tab-bar.png)

## Back Button

The back button allows users to recover from a drill-down navigational action. This pattern helps ensure users follow an ordered series of steps.

![Illustration showing a back button within an Office desktop application task pane. This example shows a back button immediately below the add-in name, in the top left.](../images/add-in-back-button.png)
