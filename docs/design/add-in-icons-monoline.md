---
title: Monoline style icon guidelines for Office Add-ins
description: ''
ms.date: 12/09/2019
localization_priority: Normal
---

# Monoline style icon guidelines for Office Add-ins

Monoline style iconography are used in Office 365. If you would prefer that your icons match the Fresh style of non-subscription Office 2013+, see [Fresh style icon guidelines for Office Add-ins](add-in-icons-fresh.md).

## Office Monoline visual style

The goal of the Monoline style to have consistent, clear, and accessible iconography to communicate action and features with simple visuals, ensure the icons are accessible to all users, and have a style that is consistent with those used elsewhere in Windows.

The following guidelines are for 3rd party developers who want to create icons for features that will be consistent with the icons already present Office products.

### Design principles

-   Simple, clean, clear.
-   Contain only necessary elements.
-   Inspired by Windows icon style.
-   Accessible to all users.

#### Conveying meaning

-   Use descriptive elements such as a page to represent a document or an envelope to represent mail.
-   Use the same element to represent the same concept, i.e., mail is always represented by an envelope, not a stamp.
-   Use a core metaphor during concept development.

#### Reduction of Elements

-   Reduce the icon to its core meaning, using only elements that are essential to the metaphor.
-   Limit the number of elements in an icon to two, regardless of icon size.

#### Consistency

Sizes, arrangement, and color of icons should be consistent.

#### Styling

##### Perspective

Monoline icons are forward-facing by default. Certain elements that require perspective and/or rotation, such as a cube, are allowed, but exceptions should be kept to a minimum.

##### Embellishment

Monoline is a clean minimal style. Everything uses flat color, which means there are no gradients, textures, or light sources.

## Designing

### Sizes

We recommend that you produce each icon in all these sizes to support high DPI devices. The absolutely *required* sizes are 16px, 20px, and 32px, as those are the 100% sizes.

**16px, 20px, 24px, 32px, 40px, 48px, 64px, 80px, 96px**

### Layout

The following is an example of icon layout with a modifier.

![Example of icon with modifier](../images/monolineicon1.png)  ![The same example with a grid background callouts for the base, modifier, padding, and cutout.](../images/monolineicon2.png)

#### Elements

- **Base**: The main concept that the icon represents. This is usually the only visual needed for the icon, but sometimes the main concept can be enhanced with a secondary element, a modifier.

- **Modifier** Any element that overlays the base; that is, a modifier that typically represents an action or a status. It modifies the base element by acting as an addition, alteration, or a descriptor.

![Grid with the base area and modifier areas on it.](../images/monolineicon3.png)

### Construction

#### Element placement

Base elements are placed in the center of the icon within the padding. If it can't be placed perfectly centered, then the base should err to the top right. In the following example, the icon is perfectly centered:

![Image showing perfectly centered icon](../images/monolineicon4.png)

In the following example, the icon is erring to the left.

![Image showing icon that errs to the left](../images/monolineicon5.png)

Modifiers are almost always placed in the bottom right corner of the icon canvas. In some rare cases, modifiers are placed in a different corner. For example, if the base element would be unrecognizable with the modifier in the bottom right corner, then consider placing it in the upper left corner.

![Image showing a few icons with the modifier in the lower right, but one with the modifier in the upper left](../images/monolineicon6.png)

#### Padding

Each size icon has a specified amount of padding around the icon. The base element stays within the padding, but the modifier should butt up to the edge of the canvas, extending outside of the padding---to the edge of the icon border. The following images show the recommended padding to use for each of the icon sizes.

|**16px**|**20px**|**24px**|**32px**|**40px**|**48px**|**64px**|**80px**|**96px**|
|:---|:---|:---|:---|:---|:---|:---|:---|:---|
|![16 px icon](../images/monolineicon7.png)|![20 px icon](../images/monolineicon8.png)|![24 px icon](../images/monolineicon9.png)|![32 px icon](../images/monolineicon10.png)|![40 px icon](../images/monolineicon11.png)|![48 px icon](../images/monolineicon12.png)|![64 px icon](../images/monolineicon13.png)|![80 px icon](../images/monolineicon14.png)|![96 px icon](../images/monolineicon15.png)|

#### Line weights

Monoline is a style dominated by line and outlined shapes. Depending on what size you are producing the icon should use the following line weights.

|**Icon Size:**|**16px**|**20px**|**24px**|**32px**|**40px**|**48px**|**64px**|**80px**|**96px**|
|:---|:---|:---|:---|:---|:---|:---|:---|:---|:---|
|**Line Weight:**|1px|1px|1px|1px|2px|2px|2px|2px|3px|
||![16 px icon](../images/monolineicon16.png)|![20 px icon](../images/monolineicon17.png)|![24 px icon](../images/monolineicon18.png)|![32 px icon](../images/monolineicon19.png)|![40 px icon](../images/monolineicon20.png)|![48 px icon](../images/monolineicon21.png)|![64 px icon](../images/monolineicon22.png)|![80 px icon](../images/monolineicon23.png)|![96 px icon](../images/monolineicon24.png)|

#### Cutouts

When an icon element is placed on top of another element, a cutout (of the bottom element) is used to provide space between the two elements, mainly for readability purposes. This usually happens when a modifier is placed on top of a base element, but there are also cases where neither of the elements is a modifier. These cutouts between the two elements is
sometimes referred to as a "gap".

The size of the gap should be the same width as the line weight used on that size. If making a 16px icon, the gap width would be 1px and if it is a 48px icon then the gap should be 2px. The following example shows a 32px icon with a gap of 1px between the modifier and the underlying base.

![32px icon with a gap of 1px between the modifier and the underlying base](../images/monolineicon25.png)

In some cases, the gap can be increase by a 1/2px if the modifier has a diagonal or curved edge and the standard gap doesn't provide enough separation. This will likely only affect the icons with 1px line weight; 16px, 20px, 24px, and 32px.

#### Background fills

Most icons in the Monoline icon set require background fills. However, there are cases where the object would not naturally have a fill, so no fill should be applied. The following icons have a white fill:

![Five icons have a white fill](../images/monolineicon26.png)

The following icons have no fill. (The gear icon is included to show that the center hole is not filled.)
![Five icons with no fill](../images/monolineicon27.png)

##### Best practices for fills

###### Dos:

- Fill any element that has a defined boundary, and would naturally have a fill.
- Use a separate shape to create the background fill.
- Use **Background Fill** from the [color palette](#color).
- Maintain the pixel separation between overlapping elements.
- Fill between multiple objects.

###### Don'ts:

- Don't fill objects that would not naturally be filled; for example, a paperclip.
- Don't fill brackets.
- Don't fill behind numbers or alpha characters.

### Color

The color palette has been designed for simplicity and accessibility. It contains 4 neutral colors and two variations for blue, green, yellow, red, and purple. Orange is intentionally not included in the Monoline icon color palette. Each color is intended to be used in specific ways as outlined in this section.

#### Palette

![The four shades of gray in monoline](../images/monoline-grayshades.png)

![The color palette in monoline](../images/monoline-colors.png)

#### How to use color

In the Monoline color palette, all colors have Standalone, Outline, and Fill variations. Generally, elements are constructed with a fill and a border. The colors are applied in one of the following patterns:

- The Standalone color alone for objects that have no fill.
- The border uses the Outline color and the fill uses the Fill color.
- The border uses the Standalone color and the fill uses the Background Fill color.

The following are examples of using color.

![Three icons with color in a border or fill or both](../images/monolineicon28.png)

The most common situation will be to have an element use Dark Gray Standalone with Background Fill.

When using a colored Fill, it should always be with its corresponding Outline color. For example, Blue Fill should only be used with Blue Outline. But there are two exceptions to this general rule:

- Background Fill can be used with any color Standalone.
- Light Gray Fill can be used with two different Outline colors: Dark Gray or Medium Gray.

#### When to use color

Color should be used to convey the meaning of the icon rather than for embellishment. It should **highlight the action** to the user. When a modifier is added to a base element that has color, the base element is typically turned into Dark Gray and Background Fill so that the modifier can be the element of color, such as the case below with the "X" modifier being added to the picture base in the leftmost icon of the following set.

![Five icons that use color](../images/monolineicon29.png)

You should limit your icons to **one** additional color, other than the Outline and Fill mentioned above. However, more colors can be used if it is vital for its metaphor, with a limit of two additional colors other than gray. In rare cases, there are exceptions when more colors are needed. The following are good examples of icons that use just one color.

  ![An image of five icons with one color each](../images/monolineicon30.png)

But the following icons use too many colors.

  ![An image of five icons with multiple colors](../images/monolineicon31.png)


Use **Medium Gray** for interior "content", such as grid lines in an icon of a spreadsheet. Additional interior colors are used when the content needs to show the behavior of the control.

![Five icons with medium gray interior elements](../images/monolineicon32.png)

#### Text lines

When text lines are in a "container" (for example, text on a document), use medium gray. Text lines not in a container should be **Dark Gray**.

### Text

Avoid using text characters in icons. Since Office products are used around the world, we want to keep icons as language neutral as possible.

## Production

### Icon file format

The final icons should be saved as .png image files. Use PNG format with a transparent background and have 32-bit depth.
