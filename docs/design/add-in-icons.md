---
title: Icon guidelines for Office Add-ins
description: ''
ms.date: 03/02/2019
localization_priority: Priority
---

# Icons
Icons are the visual representation of a behavior or concept. They are often used to add meaning to controls and commands. Visuals, either realistic or symbolic, enable the user to navigate the UI the same way signs help users navigate their environment. They should be simple, clear, and contain only the necessary details to enable customers to quickly parse what action will occur when they choose a control.

Office ribbon interfaces have a standard visual style. This ensures consistency and familiarity across Office apps. The guidelines will help you design a set of PNG assets for your solution that fit in as a natural part of Office.

Many HTML containers contain controls with iconography. Use Office UI Fabric’s custom font to render Office styled icons in your add-in. Fabric’s icon font contains many glyphs for common Office metaphors that you can scale, color, and style to suit your needs. If you have an existing visual language with your own set of icons, feel free to use it in your HTML canvases. Building continuity with your own brand with a standard set of icons is an important part of any design language. Be careful to avoid creating confusion for customers by conflicting with Office metaphors.


## Design icons for add-in commands

[Add-in commands](add-in-commands.md) add buttons, text, and icons to the Office UI. Your add-in command buttons should provide meaningful icons and labels that clearly identify the action the user is taking when they use a command. This article provides stylistic and production guidelines that help you design icons that integrate seamlessly with Office. 

## Office icon design principles

The Office 2013 release of the Office desktop clients includes refreshed iconography. The overriding stylistic change is reduction. The new icons include only essential communicative elements. Non-essential elements including perspective, gradients, and light source are removed. The simplified icons support faster parsing of commands and controls. Follow this style to best fit with Office.

Office icons are based on the following design principles: 

- Modern interpretation of Office icon collection 
- Fresh yet familiar  
- Simple, clear, and direct 

The following image shows icons that apply the modern design principles.

![Image showing old Office icons and the refreshed modern interpretation of the icons](../images/icons-images.png)

## Best practices

Follow these guidelines when you create your icons: 

|Do|Don't|
|:---|:---|
|Keep visuals simple and clear, focusing on the key element(s) of the communication.| Don't use artifacts that make your icon look messy.|
|Use the Office icon language to represent behaviors or concepts.|Don’t repurpose Office UI Fabric glyphs for add-in commands in the Office ribbon or contextual menus. Fabric icons are stylistically different and will not match.|
|Reuse common Office visual metaphors such as paintbrush for format or magnifying glass for find.|Don't reuse visual metaphors for different commands. Using the same icon for different behaviors and concepts can cause confusion. |
|Redraw your icons to make them small or larger. Take the time to redraw cutouts, corners, and rounded edges to maximize line clarity. |Don't resize your icons by shrinking or enlarging in size. This can lead to poor visual quality and unclear actions. Complex icons created at a larger size may lose clarity if resized to be smaller without redraw. |
|Use a white fill for accessibility. Most objects in your icons will require a white background to be legible across Office UI themes and in high-contrast modes.  ||
|Use the PNG format with a transparent background. ||
|Avoid localizable content in your icons, including typographic characters, indications of paragraph rags, and question marks. ||



## Icon size recommendations and requirements

Office desktop icons are bitmap images. Different sizes will render depending on the user's DPI setting and touch mode. Include all eight supported sizes to create the best experience in all supported resolutions and contexts. The following are the supported sizes - three are required:

- 16 px (Required)
- 20 px
- 24 px
- 32 px (Required)
- 40 px
- 48 px
- 64 px (Recommended, best for Mac)
- 80 px (Required)  

Make sure to redraw your icons for each size rather than shrink them to fit.

![Illustration that shows the recommendation to resize icons rather than shrink icons](../images/icon-resizing.png)

<!--
The following table shows the icon sizes that render for different modes at different DPI settings.

|DPI |**Small**||**Medium**||**Large**||**Extra large**|
|:---|:---|:---|:---|:---|:---|:---|:---|
|    |**Mouse**|**Touch**|**Mouse**|**Touch**|**Mouse**|**Touch**|-|
|100%|16px|20px|24px||32px|40px|48px|
|125%|20px|24px|||40px|48px|60px|
|150%|24px|24px|36px||48px|48px|72px|
|200%|32px|40px|48px||64px|80px|96px|
|250%|40px||||80px||120px|
|300%|48px||||96px||144px

> [!NOTE]
> At DPI settings of 150% or greater, the icon does not get swapped out for a larger size when Touch mode is engaged. At DPI settings greater than 250%, Touch mode is turned off by default.

The following table lists the locations for certain icon sizes.

|Location|100% DPI|200% DPI|250% DPI|
|:-------|:-------|:-------|:-------|
|Small ribbon button|16px|32px|40px|
|Contextual menu|16px|32px|40px|
|Quick access toolbar (QAT)|16px|32px|40px|
|Large ribbon icon|32px|64px|80px|

-->

## Icon anatomy and layout

Office icons are typically comprised of a base element with action and conceptual modifiers overlayed. Action modifiers represent concepts such as add, open, new, or close. Conceptual modifiers represent status, alteration, or a description of the icon. 

To create commands that align with the Office UI, follow layout guidelines for the base element and modifiers. This ensures that your commands look professional and that your customers will trust your add-in. If you make exceptions to these guidelines, do so intentionally.

The following image shows the layout of base elements and modifiers in an Office icon.

![Image showing an icon base element in the center with a modifier on the lower right and an action modifier on the upper left](../images/icon-layouts.png)

- Center base elements in the pixel frame with empty padding all around.
- Place action modifiers on the top left. 
- Place conceptual modifiers on the bottom right.
- Limit the number of elements in your icons. At 32px, limit the number of modifiers to a maximum of two. At 16px, limit the number of modifiers to one.

###Base element padding
Place base elements consistently across sizes. If base elements can't be centered in the frame, align them to the top left, leaving the extra pixels on the bottom right. For best results, apply the padding guidelines listed in the following table.

###Modifiers
All modifiers should have a 1px transparent cutout between each element, including the background. Elements should not directly overlap. Create whitespace between rules and edges. Modifiers can vary slightly in size, but use these dimensions as a starting point.


|**Icon size**|**Padding around base element**|**Modifier size**|
|:---|:---|:---|
|16px|0|9px|
|20px|1px|10px|
|24px|1px|12px|
|32px|2px|14px|
|40px|2px|20px|
|48px|3px|22px|
|64px|5px|29px|
|80px|5px|38px|


## Icon colors

> [!NOTE]
> These color guidelines are for ribbon icons used in [Add-in commands](add-in-commands.md). These icons are not rendered with Microsoft UI Fabric and the color palate is different from the palate described at [Microsoft UI Fabric | Colors | Shared](https://fluentfabric.azurewebsites.net/#/color/shared).

Office icons have a limited color palette. Use the colors listed in the following table to guarantee seamless integration with the Office UI. Apply the following guidelines to the use of color: 

- Use color to communicate meaning rather than for embellishment. It should highlight or emphasize an action, status, or an element that explicitly differentiates the mark.  
- If possible, use only one additional color beyond gray. Limit additional colors to two at the most.
- Colors should have a consistent appearance in all icon sizes. Office icons have slightly different color palettes for different icon sizes. 16px and smaller icons are slightly darker and more vibrant than 32px and larger icons. Without these subtle adjustments, colors appear to vary across sizes.   

|**Color name**|**RGB**|**Hex**|**Color**|**Category**|
|:---|:---|:---|:---|:---|
|Text Gray (80)|80, 80, 80|#505050| ![Text gray 80 color image](../images/color-text-gray-80.png) |Text|
|Text Gray (95)|95, 95, 95|#5F5F5F| ![Text gray 95 color image](../images/color-text-gray-95.png) |Text|
|Text Gray (105)|105, 105, 105|#696969| ![Text gray 105 color image](../images/color-text-gray-105.png) |Text|
|Dark Gray 32|128, 128, 128|#808080| ![Dark gray 32 color image](../images/color-dark-gray-32.png) |32 and above|
|Medium Gray 32|158, 158, 158|#9E9E9E| ![Medium gray 32 color image](../images/color-medium-gray-32.png) |32 and above|
|Light Gray ALL|179, 179, 179|#B3B3B3| ![Light gray all color image](../images/color-light-gray-all.png) |All sizes|
|Dark Gray 16|114, 114, 114|#727272| ![Dark gray 16 color image](../images/color-dark-gray-16.png) |16 and below|
|Medium Gray 16|144, 144, 144|#909090| ![Medium gray 16 color image](../images/color-medium-gray-16.png) |16 and below|
|Blue 32|77, 130, 184|#4d82B8| ![Blue 32 color image](../images/color-blue-32.png) |32 and above|
|Blue 16|74, 125, 177|#4A7DB1| ![Blue 16 color image](../images/color-blue-16.png) |16 and below|
|Yellow ALL|234, 194, 130|#EAC282| ![Yellow all color image](../images/color-yellow-all.png) |All sizes|
|Orange 32|231, 142, 70|#E78E46| ![Orange 32 color image](../images/color-orange-32.png) |32 and above|
|Orange 16|227, 142, 70|#E3751C| ![Orange 16 color image](../images/color-orange-16.png) |16 and below|
|Pink ALL|230, 132, 151|#E68497| ![Pink all color image](../images/color-pink-all.png) |All sizes|
|Green 32|118, 167, 151|#76A797| ![Green 32 color image](../images/color-green-32.png) |32 and above|
|Green 16|104, 164, 144|#68A490| ![Green 16 color image](../images/color-green-16.png) |16 and below|
|Red 32|216, 99, 68|#D86344| ![Red 32 color image](../images/color-red-32.png) |32 and above|
|Red 16|214, 85, 50|#D65532| ![Red 16 color image](../images/color-red-16.png) |16 and below|
|Purple 32|152, 104, 185|#9868B9| ![Purple 32 color image](../images/color-purple-32.png) |32 and above|
|Purple 16|137, 89, 171|#8959AB| ![Purple 16 color image](../images/color-purple-16.png) |16 and below|


## Icons in high contrast modes

Office icons are designed to render well in high contrast modes. Foreground elements are well differentiated from backgrounds to maximize legibility and enable recoloring. In high contrast modes, Office will recolor any pixel of your icon with a red, green, or blue value less than 190 to full black. All other pixels will be white. In other words, each RGB channel is assessed where 0-189 values are black and 190-255 values are white. Other high-contrast themes recolor using the same 190 value threshold but with different rules. For example, the high-contrast white theme will recolor all pixels greater than 190 opaque but all other pixels as transparent. Apply the following guidelines to maximize legibility in high-contrast settings:

- Aim to differentiate foreground and background elements along the 190 value threshold.
- Follow Office icon visual styles.
- Use colors from our icon palette.
- Avoid the use of gradients.
- Avoid large blocks of color with similar values.

## See also

- [Add-in development best practices](../concepts/add-in-development-best-practices.md)
- [Add-in commands for Excel, Word, and PowerPoint](../design/add-in-commands.md)




- Avoid relying on your logo or brand to communicate what an add-in command does. Brand marks aren't always recognizable at smaller icon sizes and when modifiers are applied. Brand marks often conflict with Office ribbon icon styles, and can compete for user attention in a saturated environment.


