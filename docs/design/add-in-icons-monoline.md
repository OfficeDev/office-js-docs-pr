---
title: Monoline style icon guidelines for Office Add-ins
description: ''
ms.date: 12/09/2019
localization_priority: Priority
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

-   Use the same element to represent the same concept, i.e. mail is always represented by an envelope, not a stamp.

-   Use a core metaphor during concept development.

#### Reduction of Elements

-   Reduction of the icon to its core meaning, using only elements that are essential to the metaphor.

-   Limiting the number of elements in an icon to two, regardless of icon size.

#### Consistency

Sizes, arrangement, and color of icons should be consistent.

#### Styling

##### Perspective

Monoline icons are forward-facing by default. Certain elements that require perspective and/or rotation, such as a cube are allowed, but exceptiosn should be kept to a minimum.

##### Embellishment

Monoline is a clean minimalistic style. Everything uses flat color, which means there are no gradients, textures, or light sources.

## Designing

### Sizes

Multiple sizes will be needed. It is recommended to produces the following sizes:

**16x, 20x, 24x, 32x, 40x, 48x, 64x, 80x, 96x**

We recommend that you produce each icon in all these sizes to support high DPI devices. The absolutely *required* sizes are 16x, 20x, and 32x, as those are the 100% sizes.

### Layout

The following is an example of icon layout with a modifier.

![Example of icon with modifier](../images/monolineicon1.png)!  [The same example with a grid background callouts for the base, modifier, padding and modifier cutout](../images/monolineicon2.png)

#### Elements

- **Base**: The main concept that the icon represents. This is usually the only visual needed for the icon, but sometimes the main concept can be enhanced with a secondary element, a modifier.

- **Modifier** Any element that overlays the base; that is, a modifier that typically represents an action or a status. It modifies the base element by acting as an addition, alteration, or a descriptor.

![Grid with the base area and modifier areas on it.](../images/monolineicon3.png)

### Construction

#### Element placement

Base elements are placed in the center of the icon with in the padding. If it cannot place perfectly centered, then the base should err to the top right. In the following examples, the icon on the left is perfectly centered, and the one on the right is erring to the left.

![Image showing perfectly centered icon](../images/monolineicon4.png)   ![Image showing icon that errs to the left](../images/monolineicon5.png)

Modifiers are almost always placed in the bottom right corner of the icon canvas. In some rare cases, modifiers are placed in a different corner. For example, if the base element would be unrecognizable with the modifier in the bottom right corner, then consider placing it in the upper left corner.

![Image showing a few icons with the modifier in the lower right, but one with the modifier in the upper left](../images/monolineicon6.png)

#### Padding

Each size icon has a specified amount of padding around the icon. The base element stays within the padding, but the modifier should but up to the edge of the canvas, extending outside of the padding---to the edge of the icon border. The following images show the recommended padding to use for each of the icon sizes.

|**16x**|**20x**|**24x**|**32x**|**40x**|**48x**|**64x**|**80x**|**96x**|
|:---|:---|:---|:---|:---|:---|:---|:---|:---|
|![16 px icon](../images/monolineicon7.png)|![20 px icon](../images/monolineicon8.png)|![24 px icon](../images/monolineicon9.png)|![32 px icon](../images/monolineicon10.png)|![40 px icon](../images/monolineicon11.png)|![48 px icon](../images/monolineicon12.png)|![64 px icon](../images/monolineicon13.png)|![80 px icon](../images/monolineicon14.png)|![96 px icon](../images/monolineicon15.png)|

#### Line weights

Monoline is a style dominated by line and outlined shapes. Depending on what size you are producing the icon should use the following line weights.

|**Icon Size:**|**16x**|**20x**|**24x**|**32x**|**40x**|**48x**|**64x**|**80x**|**96x**|
|:---|:---|:---|:---|:---|:---|:---|:---|:---|:---|
|**Line Weight:**|1px|1px|1px|1px|2px|2px|2px|2px|3px|
||![16 px icon](../images/monolineicon16.png)|![20 px icon](../images/monolineicon17.png)|![24 px icon](../images/monolineicon18.png)|![32 px icon](../images/monolineicon19.png)|![40 px icon](../images/monolineicon20.png)|![48 px icon](../images/monolineicon21.png)|![64 px icon](../images/monolineicon22.png)|![80 px icon](../images/monolineicon23.png)|![96 px icon](../images/monolineicon24.png)|

#### Cutouts

When an icon element is placed on top of another element, a cutout (of the bottom element) is used to provide space between the two elements, mainly for readability purposes. This usually happens when a modifier is placed on top of a base element, but there are also cases where neither of the elements is a modifier. These cutouts between the two elements is
sometimes referred to as a "gap".

The size of the gap should be the same width as the line weight used on that size. If making a 16x icon, the gap width would be 1px and if it is a 48x icon then the gap should be 2px. The following example shows a 32x icon with a gap of 1px between the modifier and the underlying base.

![32x icon with a gap of 1px between the modifier and the underlying base](../images/monolineicon25.png)

In some cases, the gap can be increase by a 1/2px if the modifier has a diagonal or curved edge and the standard gap doesn't provide enough separation. This will likely only affect the icons with 1px line weight; 16x, 20x, 24x, and 32x.

#### Background fills

Most icons in the Monoline icon set require background fills. However, there are cases where the object would not naturally have a fill, so no fill should be applied. The following icons have a white fill:

![Five icons have a white fill](../images/monolineicon26.png)

The following icons have no fill. (The gear icon is included to show that the center hole is not filled.)
![Five icons with no fill](../images/monolineicon27.png)

##### Best practices for fills

Dos:

- Any element that has a defined boundary, and would naturally have a fill.
- Use a separate shape to create the background fill.
- Use **Background Fill** from the Universal color palette.
- Maintain the pixel separation between overlapping elements.
- Fill between multiple objects.

Don'ts:

- Don't fill objects that would not naturally be filled; for example, a paperclip.
- Don't fill brackets.
- Don't fill behind numbers or alpha characters.

### Color

The color palette has been designed for simplicity and accessibility. It contains 4 neutral colors and two variations for blue, green, yellow, red, and purple. Orange is intentionally not included in the Office icon color palette. Each color is intended to be used in specific ways as outlined in this section.

#### Palette


|**Sample**|**Description**|
|:---|:---|
|<span style="background-color:#3A3A38>x</span>|Dark Gray -- Standalone/Outline<br>58,58,56<br>#3A3A38|
|x|Medium Gray -- Outline/Content<br>121,119,116<br>#797774|
|x|Background Fill<br>250,250,250<br>#FAFAFA|
|x|Light Gray - Fill<br>200,198,196<br>#C8C6C4|




  -------------------------------------
     Dark Gray -- Standalone/Outline\
     58,58,56\
     \#3A3A38
  -- ----------------------------------
     Medium Gray -- Outline/Content\
     121,119,116\
     \#797774

     Background Fill\
     250,250,250\
     \#FAFAFA

     Light Gray - Fill\
     200,198,196\
     \#C8C6C4
  -------------------------------------

###

  -------------------------
     Blue -- Standalone\
     30,139,205\
     \# 1E8BCD
  -- ----------------------
     Green - Standalone\
     24,171,80\
     \#18AB50

     Yellow - Standalone\
     251,152,59\
     \#FB983B

     Red - Standalone\
     237,61,59\
     \#ED3D3B

     Purple - Standalone\
     168,70,178\
     \#A846B2

     Blue -- Outline\
     0,99,177\
     \#0063B1

     Green - Outline\
     48,144,72\
     \#309048

     Yellow - Outline\
     237, 135, 51\
     \#ED8733

     Red - Outline\
     212, 35, 20\
     \#D42314

     Purple - Outline\
     146, 46, 155\
     \#922E9B

     Blue - Fill\
     131, 190, 236\
     \#83BEEC

     Green - Fill\
     161, 221, 170\
     \#A1DDAA

     Yellow - Fill\
     248, 219, 143\
     \#F8DB8F

     Red - Fill\
     255, 145, 152\
     \#FF9198

     Purple - Fill\
     212, 146, 216\
     \#D492D8
  -------------------------

#### How To Use Color

In the Monoline color pallet, all colors have a Standalone, Outline and Fill variations. Generally, elements are constructed with a fill and a border. Color is applied either:

-   With the border using the Outline color and the fill using the Fill color

-   With the border using the Standalone color and the fill using the Background Fill color

-   If the element has no fill, then the Standalone color is used


![Image showing old Office icons and the refreshed modern interpretation of the icons](../images/monolineicon28.png)
  ---------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------
  ![C:\\Users\\Suzanne Alphin\\AppData\\Local\\Microsoft\\Windows\\INetCache\\Content.Word\\Guildlines\_ColorLable.png](media/image29.png){width="5.0in" height="1.3541666666666667in"}
  ---------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------

The most common situation will be to have an element use Dark Gray Standalone with Background Fill.

When using a colored Fill, it should always be with its corresponding Outline color. Example, Blue Fill should only be used with Blue Outline.

##### Exceptions:

- Background Fill can be used with any color Standalone.

- Light Gray Fill can be used with 2 different Outline colors: Dark Gray or Medium Gray

#### When To Use Color

Color should be used to convey the meaning of the icon rather than for embellishment. It should **highlight the action** to the user. When a modifier is added to a base element that has color, the base element is typically turned into Dark gray and Background Fill so that the modifier can be the element of color, such as the case below with the \"X\"
modifier being added to the picture base.

![Image showing old Office icons and the refreshed modern interpretation of the icons](../images/monolineicon29.png)
  -----------------------------------------------------------------------------------
  ![](media/image30.png){width="4.166666666666667in" height="0.6666666666666666in"}
  -----------------------------------------------------------------------------------

Color is generally **limited to 1 additional color** other than the Outline and Fill mentioned above. However, more colors can be used if it is vital for its metaphor, with a limit to 2 additional colors other than gray. In rare cases, there are exceptions when more colors are needed.

![Image showing old Office icons and the refreshed modern interpretation of the icons](../images/monolineicon30.png)

  ✔
  -----------------------------------------------------------------------------------
  ![](media/image31.png){width="4.166666666666667in" height="0.6666666666666666in"}

  ![Image showing old Office icons and the refreshed modern interpretation of the icons](../images/monolineicon31.png)

  ❌
  -----------------------------------------------------------------------------------
  ![](media/image32.png){width="4.166666666666667in" height="0.6666666666666666in"}

Use **Medium Gray** for interior \"content\", grid lines and dashes. Additional interior colors are used when the content needs to show the behavior of the control

![Image showing old Office icons and the refreshed modern interpretation of the icons](../images/monolineicon32.png)

  -----------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------
  ![C:\\Users\\Suzanne Alphin\\AppData\\Local\\Microsoft\\Windows\\INetCache\\Content.Word\\Guildlines\_MediumGray.png](media/image33.png){width="4.166666666666667in" height="0.6666666666666666in"}
  -----------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------

#### Text Lines

When text lines are in a \"container\" (text on a document), use medium grey. Text lines not in a container should be dark grey.

### Text

Avoid using text characters in icons. Since Office products are used around the world, we want to keep icons as language neutral as possible.

## Production

### Icon file format

The final icons should be saved out as .png image files. Use PNG format with a transparent background and have 32-bit depth.

### Sizes

As stated earlier the final recommended sizes are:

**16x, 20x, 24x, 32x, 40x, 48x, 64x, 80x, 96x**

However, the absolutely required sizes are 16x, 20x, and 32x.
