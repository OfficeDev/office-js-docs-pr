# Content Office Add-ins

Content add-ins are surfaces that can be embedded directly into Word, Excel, or PowerPoint documents. Content add-ins give users access to interface controls that run code to modify documents or display data from a data source. Use content add-ins when you want to embed functionality directly into the document.  

**Example: Content add-in**

![An example image displaying a typical layout for content add-ins.](../../images/overview_withApp_content.png)

## Best practices

- Include some navigational or commanding element such as the CommandBar or Pivot at the top of your add-in.
- Include a branding element such as the BrandBar at the bottom of your add-in (applies to Word, Excel, and PowerPoint add-ins only).

## Variants

Content add-in sizes for Word, Excel, and PowerPoint in Office 2016 desktop and Office 365 are user specified.

## Personality menu

Personality menus can obstruct navigational and commanding elements located near the top right of the add-in. The following are the current dimensions of the personality menu on Windows and Mac.

**Personality menu on Windows** 

For Windows, the personality menu measures 12x32 pixels, as shown.

![Image showing the personality menu on Windows desktop](../../images/personalityMenu_Win.png)

**Personality menu on Mac**

For Mac, the personality menu measures 26x26 pixels, but floats 8 pixels in from the right and 6 pixels from the top, which increases the occupied space to 34x32 pixels, as shown.

![Image showing the personality menu on Mac desktop](../../images/personalityMenu_Mac.png)

## Implementation

For a sample that implements a content add-in, see [Excel Content Add-in Humongous Insurance](https://github.com/OfficeDev/Excel-Content-Add-in-Humongous-Insurance) in GitHub.

## Additional resources

- [Office UI Fabric in Office Add-ins](office-ui-fabric.md) 
- [UX design patterns for Office Add-ins](ux-design-patterns.md)