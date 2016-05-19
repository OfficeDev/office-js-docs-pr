# Design Office Add-ins

Desgin your Office Add-in UX to integrate seamlessly with Office to provide an efficient, natural experience for your users. Take advantage of [add-in commands](add-in-commands.md) (Office UI extensions) to provide access to your add-in and apply the [best practices](add-in-development-best-practices.md) that we recommend when you create custom HTML-based UI. 
 
 
Apply the following principles as you design your add-in: 

- **Design explicitly for Office**. The functionality and look and feel of an add-in must harmoniously complement the Office experience, including applying the the Office or document theme.
 
- **Make users more efficient**. Help users get one job done without getting in the way of other jobs. Allow for seamless interaction between Office documents and your add-in. 

- **Favor content over chrome**. Emphasize the add-in's content and functionality over any accessory chrome. Maximize the use of space by avoiding superfluous UI elements that don't add value to the user experience.  

- **Keep users in control**. Allow users to control their experience, understand any important decisions, and easily reverse actions the add-in performs. 

- **Design for all platforms and input methods**. Add-ins are designed to work on all the platforms that Office supports, and your add-in UX should be optimized to work across platforms and form factors. Support mouse/keyboard and [touch](add-in-development-best-practices.md#bk_Touch) input devices, and ensure that your custom HTML UI is responsive to adapt to different form factors. 


##UX design patterns

To help you create a first-class user experience for your add-in, we provide templates that illustrate common UX design patterns. These templates reflect [best practices](add-in-development-best-practices.md) for creating compelling, world-class add-ins, and include patterns for first-run experiences, branding elements, and user notifications. They use [Office UI Fabric](https://dev.office.com/fabric) components and styles and include elements that naturally extend the Office UI.

To access the templates, see [Office Add-in UX Design Patterns](https://github.com/OfficeDev/Office-Add-in-UX-Design-Patterns) on GitHub. The Adobe Illustrator files are also available; you can download and update them to reflect your own designs. You can also copy the code files from the [Office Add-in UX design patterns code](https://github.com/OfficeDev/Office-Add-in-UX-Design-Patterns-Code) repo to your add-in project and customize them as needed. 

##Color
If you decide to use your own color palette: 
 
- Use color to help communicate your brand value to users, and to add emotion and delight to your add-in user experience.
- Use color meaningfully and consistently in your add-in. For example, choose one color as an accent to give your add-in a consistent visual theme.
- Avoid using the same color for both interactive and non-interactive elements. If you use color to indicate items users can interact with, such as navigation, links, and buttons, don't use the same color for static items.
- If you use color for text or white text on a colored background, be sure that your colors have enough contrast to meet accessibility guidelines (4.5:1 contrast ratio).
- Be aware of color blindness —- use more than just colors to indicate interactivity.

##Theming
Whether you decide to adopt the Office color scheme or to use your own, we encourage you to use our Theming APIs. Add-ins that are part of the Office theming experience will feel much more integrated with Office.

- For Outlook, Word, and Excel add-ins, use the [Context.officeTheme](../../reference/shared/office.context.officetheme.md) property to match the theme of the Office applications. This API is currently only available in Office 2016.  
- For PowerPoint add-ins that create new objects, see [Use Office themes in your PowerPoint add-ins](../../use-document-themes-in-your-powerpoint-add-ins.md).

##Additional resources

- [Office voice](https://msdn.microsoft.com/en-us/library/mt484351.aspx)
- [Create accessible add-ins](https://msdn.microsoft.com/en-us/library/mt598623.aspx)