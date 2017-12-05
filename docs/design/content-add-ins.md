---
title: Content Office Add-ins
description: ''
ms.date: 12/04/2017
---



# Content Office Add-ins

Content add-ins are surfaces that can be embedded directly into Word, Excel, or PowerPoint documents. Content add-ins give users access to interface controls that run code to modify documents or display data from a data source. Use content add-ins when you want to embed functionality directly into the document.  

*Figure 1. Typical layout for content add-ins*

![An example image displaying a typical layout for content add-ins.](../images/overview-with-app-content.png)

## Best practices

- Include some navigational or commanding element such as the CommandBar or Pivot at the top of your add-in.
- Include a branding element such as the BrandBar at the bottom of your add-in (applies to Word, Excel, and PowerPoint add-ins only).

## Variants

Content add-in sizes for Word, Excel, and PowerPoint in Office 2016 desktop and Office 365 are user specified.

## Personality menu

Personality menus can obstruct navigational and commanding elements located near the top right of the add-in. The following are the current dimensions of the personality menu on Windows and Mac.

For Windows, the personality menu measures 12x32 pixels, as shown.

*Figure 2. Personality menu on Windows* 

![Image showing the personality menu on Windows desktop](../images/personality-menu-win.png)


For Mac, the personality menu measures 26x26 pixels, but floats 8 pixels in from the right and 6 pixels from the top, which increases the occupied space to 34x32 pixels, as shown.

*Figure 3. Personality menu on Mac*

![Image showing the personality menu on Mac desktop](../images/personality-menu-mac.png)

## Implementation

For a sample that implements a content add-in, see [Excel Content Add-in Humongous Insurance](https://github.com/OfficeDev/Excel-Content-Add-in-Humongous-Insurance) in GitHub.

## See also

- [Office UI Fabric in Office Add-ins](office-ui-fabric.md) 
- [UX design patterns for Office Add-ins](ux-design-patterns.md)
