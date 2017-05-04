# Office UI Fabric in Office Add-ins 

Office UI Fabric is a JavaScript front-end framework for building user experiences for Office and Office 365. Fabric provides visuals-focused components that you can extend, rework, and use in your Office Add-in. Because Fabric uses the Office Design Language, Fabric's UX components look like a natural extension of Office. 

If you are building an add-in, we encourage you to use Office UI Fabric to create your user experience. Using Office UI Fabric is optional.

The following sections explain how to get started using Fabric to meet your requirements. 

## Use Fabric Core: icons, fonts, colors
Fabric Core contains basic elements of the design language such as icons, colors, type, and grid. Fabric core is framework independent. Both Fabric React and Fabric JS use Fabric Core.

To get started using Fabric Core:

1. Add the CDN reference to the HTML on your page.  

	`<link rel="stylesheet" href="https://static2.sharepointonline.com/files/fabric/office-ui-fabric-js/1.4.0/css/fabric.min.css">`   
    
2. Use Fabric icons and fonts. 

To use a Fabric icon, include the "i" element on your page, and then reference the appropriate classes. You can control the size of the icon by changing the font size. For example, the following code shows how to make an extra-large table icon that uses the themePrimary (#0078d7) color. 
   
`<i class="ms-Icon ms-font-xl ms-Icon--Table ms-fontColor-themePrimary"></i>`

To find more icons that are available in Office UI Fabric, use the search feature on the [Icons](https://dev.office.com/fabric#/styles/icons) page. When you find an icon to use in your add-in, be sure to prefix the icon name with `ms-Icon--`. 

For information about font sizes and colors that are available in Office UI Fabric, see [Typography](https://dev.office.com/fabric#/styles/typography) and [Colors](https://dev.office.com/fabric#/styles/colors).
 
## Using Fabric Components 
Fabric provides a variety of UX components that you can use to build your add-in. There are various types of components including:

- Input components, for example Button, Checkbox, and Toggle.
- Navigation components, for example Pivot, and Breadcrumb.
- Notification components, for example MessageBar, and Callout.  

Not all Fabric components are recommended for use in add-ins. We provide guidance on using Fabric components in add-ins. For example, see [Button](buttons.md) for guidance on using a Fabric Button in your add-in. 

You may use different JavaScript frameworks, such as Angular or React, to build your add-in. To get started using Fabric components with your framework, see the following resources.

|**Framework**|**Example**|
|:------------|:----------|
|**JavaScript only** (no framework)|[Using Office UI Fabric JS in Office Add-ins](using-office-ui-fabric-js.md).|
|**React**|[Using Office UI Fabric React in Office Add-ins](using-office-ui-fabric-react.md )|
|**Angular**| See [ngOfficeUIFabric](http://ngofficeuifabric.com/), which is a community project with Angular 1.5 directives, and [Consider wrapping Fabric components with Angular 2 components](https://dev.office.com/docs/add-ins/develop/add-ins-with-angular2#consider-wrapping-fabric-components-with-angular-2-components)|