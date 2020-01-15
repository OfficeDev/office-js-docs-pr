---
title: Task panes in Office Add-ins
description: Task panes give users access to interface controls that run code to modify documents or emails, or display data from a data source.
ms.date: 01/14/2020
localization_priority: Priority
---


# Task panes in Office Add-ins
 
Task panes are interface surfaces that typically appear on the right side of the window within Word, PowerPoint, Excel, and Outlook. Task panes give users access to interface controls that run code to modify documents or emails, or display data from a data source. Use task panes when you don't need to embed functionality directly into the document.

*Figure 1. Typical task pane layout*

![Image displaying a typical task pane layout](../images/overview-with-app-task-pane.png)

## Best practices

|**Do**|**Don't**|
|:-----|:--------|
|<ul><li>Include the name of your add-in in the title.</li></ul>|<ul><li>Don't append your company name to the title.</li></ul>|
|<ul><li>Use short descriptive names in the title.</li></ul>|<ul><li>Don't append strings such as “add-in,” “for Word,” or “for Office” to the title of your add-in.</li></ul>|
|<ul><li>Include some navigational or commanding element such as the CommandBar or Pivot at the top of your add-in.</li></ul>||
|<ul><li>Include a branding element such as the BrandBar at the bottom of your add-in unless your add-in is to be used within Outlook.</li></ul>||


## Variants

The following images show the various task pane sizes with the Office ribbon at a 1366x768 resolution. For Excel, additional vertical space is required to accommodate the formula bar.  

*Figure 2. Office 2016 desktop task pane sizes*

![Image displaying the desktop task pane sizes at 1366x768](../images/office-2016-taskpane-sizes.png)

- Excel - 320x455
- PowerPoint - 320x531
- Word - 320x531
- Outlook - 348x535

<br/>

*Figure 3. Office 365 task pane sizes*

![Image displaying the desktop task pane sizes at 1366x768](../images/office-365-taskpane-sizes.png)

- Excel - 350x378
- PowerPoint - 348x391
- Word - 329x445
- Outlook (on the web) - 320x570

## Personality menu

Personality menus can obstruct navigational and commanding elements located near the top right of the add-in. The following are the current dimensions of the personality menu on Windows and Mac.

For Windows, the personality menu measures 12x32 pixels, as shown.

*Figure 4. Personality menu on Windows*

![Image showing the personality menu on Windows desktop](../images/personality-menu-win.png)

For Mac, the personality menu measures 26x26 pixels, but floats 8 pixels in from the right and 6 pixels from the top, which increases the space to 34x32 pixels, as shown.

*Figure 5. Personality menu on Mac*

![Image showing the personality menu on Mac desktop](../images/personality-menu-mac.png)

## Implementation

For a sample that implements a task pane, see [Excel Add-in JS WoodGrove Expense Trends](https://github.com/OfficeDev/Excel-Add-in-WoodGrove-Expense-Trends) on GitHub. 


## See also

- [Office UI Fabric in Office Add-ins](office-ui-fabric.md) 
- [UX design patterns for Office Add-ins](../design/ux-design-pattern-templates.md)

