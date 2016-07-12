# Design guidelines for Office Add-ins

Office Add-ins extend the Office experience by providing contextual functionality that users can access within Office clients. Add-ins empower users to get more done by enabling them to access third-party functionality within Office, without costly context switches. 

 Your add-in UX design must integrate seamlessly with Office to provide an efficient, natural interaction for your users. Take advantage of add-in commands (Office UI extensions) to provide access to your add-in and use the [UI elements](ui-elements/ui-elements.md) and [best practices](https://dev.office.com/docs/add-ins/overview/add-in-development-best-practices) that we recommend when you create custom HTML-based UI. 
 
 
## Core Office Add-in design principles
Regardless of the underlying framework you use to create your custom UI, apply the following principles as you design your add-in: 

- **Design explicitly for Office**. The functionality and look and feel of an add-in must harmoniously complement the Office experience, including applying the the Office or document theme.
 
- **Make users more efficient**. Help users get one job done without getting in the way of other jobs. Allow for seamless interaction between Office documents and your add-in. 

- **Favor content over chrome**. Emphasize the add-in's content and functionality over any accessory chrome. Maximize the use of space by avoiding superfluous UI elements that don't add value to the user experience.  

- **Keep users in control**. Allow users to control their experience, understand any important decisions, and easily reverse actions the add-in performs. 

- **Design for all platforms and input methods**. Add-ins are designed to work on all the platforms that Office supports, and your add-in UX should be optimized to work across platforms and form factors. Support mouse/keyboard and touch input devices, and ensure that your custom HTML UI is responsive to adapt to different form factors. For more information, see [Touch](https://msdn.microsoft.com/EN-US/library/mt590883.aspx#bk_Touch). 


## Design language
We recommend that you adopt the Office design language and use [Office UI Fabric](https://dev.office.com/fabric) to create custom HTML-based experiences in your add-ins. If your organization already has a design language, you're welcome to use it, as long as the end result is a harmonious experience for Office users. 


## Add-in building blocks
You can use two types of UI elements to create your add-ins: 

- [Add-in commands](ui-elements/ui-elements.md#add-in-commands) enable you to add native UX hooks into Office applications
- [Custom HTML-based UI](ui-elements/ui-elements.md#custom-html-based-ui) allows you to take advantage of the power of HTML within Office clients. 

For details about how to use these building blocks, see [UI elements](ui-elements/ui-elements.md).  

## UX design patterns

To help you create a first-class user experience for your add-in, we provide templates that illustrate common UX design patterns. These templates reflect [best practices](https://dev.office.com/docs/add-ins/overview/add-in-development-best-practices) for creating compelling, world-class add-ins, and include patterns for first-run experiences, branding elements, and user notifications. They use [Office UI Fabric](https://dev.office.com/fabric) components and styles and include elements that naturally extend the Office UI.

To access the templates, see the [Office Add-in UX Design Patterns](https://github.com/OfficeDev/Office-Add-in-UX-Design-Patterns) repo. The Adobe Illustrator files are also available; you can download and update them to reflect your own designs. You can also copy the code files from the [Office Add-in UX design patterns code](https://github.com/OfficeDev/Office-Add-in-UX-Design-Patterns-Code) repo to your add-in project and customize them as needed. 

## Recommended layouts and interaction patterns
We provide recommended layouts for each add-in type, along with **end-to-end** examples to help you put everything together. To learn more about how to lay out your add-in, see the following:

- [Layout for task pane containers](ui-elements/layout-for-task-pane-add-ins.md)
- [Layout for content add-ins](ui-elements/layout-for-content-add-ins.md) 
- [Layouts for mail add-ins](ui-elements/layouts-for-outlook-add-ins.md)

See also Interaction patterns for examples of common scenarios for add-ins and their corresponding interaction patterns.

## Additional resources

- [Office UI Fabric](https://dev.office.com/fabric) 

