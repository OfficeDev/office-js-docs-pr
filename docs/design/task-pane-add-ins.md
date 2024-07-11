---
title: Task panes in Office Add-ins
description: Task panes give users access to interface controls that run code to modify documents or emails, or display data from a data source.
ms.date: 08/18/2023
ms.topic: overview
ms.localizationpriority: medium
---


# Task panes in Office Add-ins

Task panes are interface surfaces that typically appear on the right side of the window within Word, PowerPoint, Excel, and Outlook. Task panes give users access to interface controls that run code to modify documents or emails, or display data from a data source. Use task panes when you don't need to embed functionality directly into the document.

*Figure 1. Typical task pane layout*

![A typical task pane layout with section tabs at the top, company logo and company name in the bottom left, and a settings icon in the bottom right.](../images/overview-with-app-task-pane.png)

## Best practices

|Do|Don't|
|:-----|:--------|
|Include the name of your add-in in the title.|Don't append your company name to the title.|
|Use short descriptive names in the title.|Don't append strings such as "add-in," "for Word," or "for Office" to the title of your add-in.|
|Include some navigational or commanding element such as the CommandBar or Pivot at the top of your add-in.|*None*|
|Include a branding element such as the BrandBar at the bottom of your add-in unless your add-in is to be used within Outlook.|*None*|

## Variants

The following images show the various task pane sizes with the Office app ribbon at a 1366x768 resolution. For Excel, additional vertical space is required to accommodate the formula bar.  

*Figure 2. Office 2016 desktop task pane sizes*

![Desktop task pane sizes at 1366x768 resolution.](../images/office-2016-taskpane-sizes.png)

- Excel - 320x455 pixels
- PowerPoint - 320x531 pixels
- Word - 320x531 pixels
- Outlook - 348x535 pixels

<br/>

*Figure 3. Office task pane sizes*

![Task pane sizes at 1366x768 resolution.](../images/office-365-taskpane-sizes.png)

- Excel - 350x378 pixels
- PowerPoint - 348x391 pixels
- Word - 329x445 pixels
- Outlook (on the web) - 320x570 pixels

## Personality menu

Personality menus can obstruct navigational and commanding elements located near the top right of the add-in. The following are the current dimensions of the personality menu on Windows and Mac. (The personality menu isn't supported in Outlook.)

For Windows, the personality menu measures 12x32 pixels, as shown.

*Figure 4. Personality menu on Windows*

![The personality menu on Windows desktop.](../images/personality-menu-win.png)

For Mac, the personality menu measures 26x26 pixels, but floats 8 pixels in from the right and 6 pixels from the top, which increases the space to 34x32 pixels, as shown.

*Figure 5. Personality menu on Mac*

![The personality menu on Mac desktop.](../images/personality-menu-mac.png)

## Implementation

For a sample that implements a task pane, see [Excel Add-in JS WoodGrove Expense Trends](https://github.com/OfficeDev/Excel-Add-in-WoodGrove-Expense-Trends) on GitHub.

## See also

- [Fabric Core in Office Add-ins](fabric-core.md)
- [UX design patterns for Office Add-ins](../design/ux-design-pattern-templates.md)
