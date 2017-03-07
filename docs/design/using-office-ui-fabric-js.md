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

    <link rel="stylesheet" href="https://static2.sharepointonline.com/files/fabric/office-ui-fabric-js/1.4.0/css/fabric.min.css">
    <link rel="stylesheet" href="https://static2.sharepointonline.com/files/fabric/office-ui-fabric-js/1.4.0/css/fabric.components.min.css">
    <script src="https://static2.sharepointonline.com/files/fabric/office-ui-fabric-js/1.4.0/js/fabric.min.js"></script>

That's it. Now you're ready to start using Fabric in your add-in. 

##2. Use Fabric icons and fonts
Using icons is simple. All you have to do is use an "i" element and reference the appropriate classes. You can control the size of the icon by changing the font size. For example, the following code shows how to make an extra-large table icon that uses the themePrimary (#0078d7) color. 
   
    <i class="ms-Icon ms-font-xl ms-Icon--Table ms-fontColor-themePrimary"></i>

To find more icons that are available in Office UI Fabric, use the search feature on the [Icons](https://dev.office.com/fabric#/styles/icons) page. When you find an icon to use in your add-in, be sure to prefix the icon name with `ms-Icon--`. 

For information about font sizes and colors that are available in Office UI Fabric, see [Typography](https://dev.office.com/fabric#/styles/typography) and [Colors](https://dev.office.com/fabric#/styles/colors).

##3. Use Fabric JS UX components

Fabric provides several UX components, like buttons or checkboxes, that you can use in your add-in. The following is a list of the Fabric JS UX components that we recommend for use in an add-in. To use one of the Fabric components in your add-in, follow the link to the Fabric documentation, and then follow the instructions in **Using this component**.

> **Note:** We will add additional components over time. 

- [Breadcrumb](https://dev.office.com/fabric-js/Components/Breadcrumb/Breadcrumb.html)
- [Button](https://dev.office.com/fabric-js/Components/Button/Button.html) (Consider using the small button variant in your add-in. Add 16px of padding to small buttons to ensure a 40px minimum touch target on touch devices.)
- [Checkbox](https://dev.office.com/fabric-js/Components/CheckBox/CheckBox.html)
- [ChoiceFieldGroup](https://dev.office.com/fabric-js/Components/ChoiceFieldGroup/ChoiceFieldGroup.html)
- [Date Picker](https://dev.office.com/fabric-js/Components/DatePicker/DatePicker.html) (For an example that shows how to implement the Date Picker in an add-in, see the [Excel Sales Tracker](https://github.com/OfficeDev/Excel-Add-in-JavaScript-SalesTracker) code sample.)
- [Dropdown](https://dev.office.com/fabric-js/Components/Dropdown/Dropdown.html)
- [Label](https://dev.office.com/fabric-js/Components/Label/Label.html)
- [Link](https://dev.office.com/fabric-js/Components/Link/Link.html)
- [List](https://dev.office.com/fabric-js/Components/List/List.html) (Consider changing the component's default styles in the CSS.)
- [MessageBanner](https://dev.office.com/fabric-js/Components/MessageBanner/MessageBanner.html)
- [MessageBar](https://dev.office.com/fabric-js/Components/MessageBar/MessageBar.html)
- [Overlay](https://dev.office.com/fabric-js/Components/Overlay/Overlay.html)
- [Panel](https://dev.office.com/fabric-js/Components/Panel/Panel.html)
- [Pivot](https://dev.office.com/fabric-js/Components/Pivot/Pivot.html)
- [ProgressIndicator](https://dev.office.com/fabric-js/Components/ProgressIndicator/ProgressIndicator.html)
- [Searchbox](https://dev.office.com/fabric-js/Components/SearchBox/SearchBox.html)
- [Spinner](https://dev.office.com/fabric-js/Components/Spinner/Spinner.html)
- [Table](https://dev.office.com/fabric-js/Components/Table/Table.html)
- [TextField](https://dev.office.com/fabric-js/Components/TextField/TextField.html)
- [Toggle](https://dev.office.com/fabric-js/Components/Toggle/Toggle.html)
   
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
 

