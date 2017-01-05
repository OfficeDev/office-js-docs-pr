-
#Use Office UI Fabric in Office Add-ins

If you are building an Office Add-in, we encourage you to use [Office UI Fabric](https://dev.office.com/fabric) to create your user experience. 

Office UI Fabric is a JavaScript front-end framework for building user experiences for Office and Office 365. Fabric provides visuals-focused components that you can extend, rework, and use in your Office Add-in. Because Fabric uses the Office Design Language, Fabric's UX components look like a natural extension of Office.

Fabric consists of several projects:

- **Fabric JS (recommended)** - Implements UX components using JavaScript only. We recommend using this version of Fabric if you don't want to take a dependency on the React framework.  
- **Fabric React** - Implements the UX components using the React framework.
- **Fabric Core** - Contains the core elements of the design language such as icons, colors, type, and grid. Both Fabric JS and Fabric React use Fabric Core. 

The following steps walk you through the basics of using Fabric JS.  

##1. Add the Fabric CDN references
To reference Fabric from the CDN, add the following HTML code to your page.

    <link rel="stylesheet" href="https://static2.sharepointonline.com/files/fabric/office-ui-fabric-js/1.2.0/css/fabric.min.css">
    <link rel="stylesheet" href="https://static2.sharepointonline.com/files/fabric/office-ui-fabric-js/1.2.0/css/fabric.components.min.css">
    <script src="https://static2.sharepointonline.com/files/fabric/office-ui-fabric-js/1.2.0/js/fabric.min.js"></script>

That's it. Now you're ready to start using Fabric in your add-in. 

##2. Use Fabric icons and fonts
Using icons is simple. All you have to do is use an "i" element and reference the appropriate classes. You can control the size of the icon by changing the font size. For example, the following code shows how to make an extra-large table icon that uses the themePrimary (#0078d7) color. 
   
    <i class="ms-Icon ms-font-xl ms-Icon--Table ms-fontColor-themePrimary"></i>

To find more icons that are available in Office UI Fabric, use the search feature on the [Icons](https://dev.office.com/fabric#/styles/icons) page. When you find an icon to use in your add-in, be sure to prefix the icon name with `ms-Icon--`. 

For information about font sizes and colors that are available in Office UI Fabric, see [Typography](https://dev.office.com/fabric#/styles/typography) and [Colors](https://dev.office.com/fabric#/styles/colors).

##3. Use Fabric JS UX components

Fabric provides several UX components, like buttons or checkboxes, that you can use in your add-in. The following is a list of the Fabric JS UX components that we recommend for use in an add-in. To use one of the Fabric components in your add-in, follow the link to the Fabric documentation, and then follow the instructions in **Using this component**.

> Note: We may add additional components over time. 

- [Breadcrumb](https://github.com/OfficeDev/office-ui-fabric-js/blob/master/ghdocs/components/Breadcrumb.md)
- [Button](https://github.com/OfficeDev/office-ui-fabric-js/blob/master/ghdocs/components/Button.md)
- [Checkbox](https://github.com/OfficeDev/office-ui-fabric-js/blob/master/ghdocs/components/CheckBox.md)
- [ChoiceFieldGroup](https://github.com/OfficeDev/office-ui-fabric-js/blob/master/ghdocs/components/ChoiceFieldGroup.md)
- [Date Picker](https://github.com/OfficeDev/office-ui-fabric-js/blob/master/ghdocs/components/DatePicker.md) (For an example that shows how to implement the Date Picker in an add-in, see the [Excel Sales Tracker](https://github.com/OfficeDev/Excel-Add-in-JavaScript-SalesTracker) code sample.)
- [Dropdown](https://github.com/OfficeDev/office-ui-fabric-js/blob/master/ghdocs/components/Dropdown.md)
- [Label](https://github.com/OfficeDev/office-ui-fabric-js/blob/master/ghdocs/components/Label.md)
- [Link](https://github.com/OfficeDev/office-ui-fabric-js/blob/master/ghdocs/components/Link.md)
- [List](https://github.com/OfficeDev/office-ui-fabric-js/blob/master/ghdocs/components/List.md) (Consider changing the component's default styles in the CSS.)
- [MessageBanner](https://github.com/OfficeDev/office-ui-fabric-js/blob/master/ghdocs/components/MessageBanner.md)
- [MessageBar](https://github.com/OfficeDev/office-ui-fabric-js/blob/master/ghdocs/components/MessageBar.md)
- [Overlay](https://github.com/OfficeDev/office-ui-fabric-js/blob/master/ghdocs/components/Overlay.md)
- [Panel](https://github.com/OfficeDev/office-ui-fabric-js/blob/master/ghdocs/components/Panel.md)
- [Pivot](https://github.com/OfficeDev/office-ui-fabric-js/blob/master/ghdocs/components/Pivot.md)
- [ProgressIndicator](https://github.com/OfficeDev/office-ui-fabric-js/blob/master/ghdocs/components/ProgressIndicator.md)
- [Searchbox](https://github.com/OfficeDev/office-ui-fabric-js/blob/master/ghdocs/components/SearchBox.md)
- [Spinner](https://github.com/OfficeDev/office-ui-fabric-js/blob/master/ghdocs/components/Spinner.md)
- [Table](https://github.com/OfficeDev/office-ui-fabric-js/blob/master/ghdocs/components/Table.md)
- [TextField](https://github.com/OfficeDev/office-ui-fabric-js/blob/master/ghdocs/components/TextField.md)
- [Toggle](https://github.com/OfficeDev/office-ui-fabric-js/blob/master/ghdocs/components/Toggle.md)
   
## Updating your add-in to use Fabric JS
If you've been using a previous version of Office UI Fabric and you'd like to move to Fabric JS, make sure that you learn about, incorporate, and test the new components in your add-in. Keep the following points in mind to help you plan for your updates:

- Component initialization is simpler using Fabric JS. For previous versions of Fabric, you include the Fabric component's JavaScript file in your add-in project, included a `<Script>` reference to that file, and then initialize the component. In Fabric JS, you no longer need to include the Fabric component's JavaScript file and the associated `<Script>` reference. All you need to do is initialize the Fabric component.   
- Several components now provide functions that control the behavior of the UX component. For example, the checkbox control has a `toggle` function that switches between the checked and unchecked states. 
- Some icon class names and styles have been updated.
- The most notable change is the use of the `<label>` element in many components. The `<label>` element controls the style of the component. You might have to update your UX code to use the `<label>` element. For example, changing the value of the `<input>` element's checked attribute on a Fabric JS checkbox has no effect on the checkbox. Instead, you  use the `check`, `unCheck`, or `toggle` functions.   

##Next steps
If you're looking for an end-to-end code sample that shows you how to use Fabric JS, we've got you covered. See the following resource:

- [Excel Sales Tracker](https://github.com/OfficeDev/Excel-Add-in-JavaScript-SalesTracker) 

##Related resources
If you're looking for code samples or documentation on a previous release of Fabric, see the following:

- [UX design patterns (uses Fabric 2.6.1)](https://github.com/OfficeDev/Office-Add-in-UX-Design-Patterns-Code) 
- [Office Add-in Fabric UI sample (uses Fabric 1.0)](https://github.com/OfficeDev/Office-Add-in-Fabric-UI-Sample) 
- [Using Fabric 2.6.1 in an Office Add-in](https://dev.office.com/docs/add-ins/design/ui-elements/using-office-ui-fabric)
 

