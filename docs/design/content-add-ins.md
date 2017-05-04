# Content Office Add-ins
 
Content add-ins are surfaces that can be embedded directly into Word, Excel, or PowerPoint documents. Content add-ins allow users to utilize interface controls that run code to modify documents or display data from a data source for example. Content add-ins should be utilized when embedding functionality directly into the document is needed and/or wanted.  

**Example: Content add-in**

![An example image displaying a typical layout for content add-ins.](../images/overview_withApp_content.png)

### Best Practices

|**Do**|**Don't**|
|:-----|:--------|
|Include some navigational or commanding element such as the CommandBar or Pivot at the top of your add-in.| |
|Include a branding element such as the BrandBar at the bottom of your add-in unless your add-in is to be used within Outlook.| |

### Variants

Office 2016 Desktop & Office 365 Online Content Add-in Sizes:
* Excel: User specified
* PowerPoint: User specified
* Word: User specified

### Personality Menu

> Note: Personality menus can obstruct navigational and commanding elements located near the top right of the add-in. Listed below are the current dimensions of the personality menu on Windows and Mac.

**Windows:** The personality menu measures 12x32 pixels.

![Image showing the personality meny on Windows Desktop](../images/personalityMenu_Win.png)

**Mac:** The personality menu measures 26x26 pixels but floats 8 pixels in from the right and 6 pixels from the top increasing the occupied space to 34x32 pixels.

![Image showing the personality meny on Mac Desktop](../images/personalityMenu_Mac.png)

## Implementation

For details, see [Office Add-ins platform overview](https://dev.office.com/docs/add-ins/overview/office-add-ins) on the Microsoft Dev Center website.

## Additional resources

* [UX Pattern Sample](https://office.visualstudio.com/DefaultCollection/OC/_git/GettingStarted-FabricReact)
* [GitHub Development Resources](https://github.com/OfficeDev/Office-Add-in-UX-Design-Patterns-Code)

