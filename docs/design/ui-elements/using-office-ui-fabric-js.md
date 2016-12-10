
#Use Office UI Fabric in Office Add-ins

If you are building an Office Add-in, we encourage you to use [Office UI Fabric](https://dev.office.com/fabric) to create your user experience. 

What is Office UI Fabric?
Office UI Fabric is a JavaScript front-end framework for building user  experiences for Office and Office 365. Fabric provides visuals-focused components to extend, re-work and use in your Office add-in. Because Fabric uses the Office Design Language, Fabric's UX components look like a natural extension of Office.

Fabric consists of several frameworks including:

- **Fabric JS (recommended)**, which implements UX components using JavaScript only. We recommend using this version of Fabric if you don't want to take a dependency on the React framework.  
- **Fabric React**, which implements the UX components using the React framework.
- **Fabric Core**, which contains the core elements of the design language such as icons, colors, type, and grid. Both Fabric JS and Fabric React use Fabric Core. 

The following steps walk you through the basics of using Fabric JS.  

##1. Add the Fabric CDN references.
To reference Fabric from the CDN, add the following HTML code to your page.

    <link rel="stylesheet" href="https://static2.sharepointonline.com/files/fabric/office-ui-fabric-js/1.2.0/css/fabric.min.css">
    <link rel="stylesheet" href="https://static2.sharepointonline.com/files/fabric/office-ui-fabric-js/1.2.0/css/fabric.components.min.css">
    <script src="https://static2.sharepointonline.com/files/fabric/office-ui-fabric-js/1.2.0/js/fabric.min.js"></script>

That's it. Now you're ready to start using Fabric in your add-in. 

##2. Use Fabric icons and fonts
Using icons are simple. All you have to do is use an "i" element and reference the appropriate classes. You can control the size of the icon by changing the font size. For example, the following code shows how to make an extra large table icon, that uses the themePrimary (#0078d7) color. 
   
    <i class="ms-Icon ms-font-xl ms-Icon--Table ms-fontColor-themePrimary"></i>

To find more icons available in Office UI Fabric, use the search feature on the [Icons](https://dev.office.com/fabric#/styles/icons) page. When you find an icon to use in your add-in, ensure you prefix the icon name with `ms-Icon--`. For more information on font sizes and colors available in Office UI Fabric, see [Typography](https://dev.office.com/fabric#/styles/typography) and [Colors](https://dev.office.com/fabric#/styles/colors).

##3. Using Fabric UX components
Fabric provides several UX components, like buttons or checkboxes, that you can use in your add-in. A number of improvements have been made to the UX components in Fabric JS including:

- You no longer need to include the UX component's JavaScript code files in your add-in project.  
- Several of the components now provide functions that control the behavior of the UX component. For example, the checkbox control has a `toggle` function that switches between the checked and unchecked states. 

> Important: Fabric JS's UX components have changed significantly. The most notable change is the use of the `<label>` element in all components. The `<label>` element controls the style of the component. If you are planning to use Fabric JS, ensure you include sufficient time to learn, incorporate, and test the new components in your add-in. For example, changing the value of the `<input>` element's checked attribute on a Fabric JS checkbox has no effect on the checkbox. Instead, you should use the `check`, `unCheck`, or `toggle` functions. 

The following table lists the Fabric JS UX components we recommend that you use in your add-in. To use one of the Fabric components in your add-in, follow the links to the Fabric documentation, and then follow the instructions in **Using this component**.

> Note: We'll be updating our recommendations over time.  

| Fabric UX component | Description	|
|:---------------|:--------|
|[Breadcrumb](https://github.com/OfficeDev/office-ui-fabric-js/blob/master/ghdocs/components/Breadcrumb.md)|Recommended|
|[Button](https://github.com/OfficeDev/office-ui-fabric-js/blob/master/ghdocs/components/Button.md)|Recommended|
|Callout|Not recommended - working on a fix.|
|[Checkbox](https://github.com/OfficeDev/office-ui-fabric-js/blob/master/ghdocs/components/CheckBox.md)|Recommended|
|[ChoiceFieldGroup](https://github.com/OfficeDev/office-ui-fabric-js/blob/master/ghdocs/components/ChoiceFieldGroup.md)|Recommended|
|CommandBar|Not recommended. Consider using [Fabric React’s CommandBar](https://dev.office.com/fabric#/components/commandbar)|
|ContextMenu|Not recommended - working on a fix.|
|[Date Picker](https://github.com/OfficeDev/office-ui-fabric-js/blob/master/ghdocs/components/DatePicker.md)|Recommended. For an example of how the Date Picker was implemented in an add-in, see the [Excel Sales Tracker](https://github.com/OfficeDev/Office-Add-in-Fabric-UI-Sample) code sample.|
|Dialog|Not recommended  - working on a fix.|
|[Dropdown](https://github.com/OfficeDev/office-ui-fabric-js/blob/master/ghdocs/components/Dropdown.md)|Recommended|
|Facepile|Not recommended - working on a fix.|
|[Label](https://github.com/OfficeDev/office-ui-fabric-js/blob/master/ghdocs/components/Label.md)|Recommended|
|[Link](https://github.com/OfficeDev/office-ui-fabric-js/blob/master/ghdocs/components/Link.md)|Recommended|
|[List](https://github.com/OfficeDev/office-ui-fabric-js/blob/master/ghdocs/components/List.md)|Recommended with changes to the styles in the CSS.|
|[MessageBanner](https://github.com/OfficeDev/office-ui-fabric-js/blob/master/ghdocs/components/MessageBanner.md)| Recommended.|
|[MessageBar](https://github.com/OfficeDev/office-ui-fabric-js/blob/master/ghdocs/components/MessageBar.md)| Recommended.|
|OrgChart|Not recommended.|
|[Overlay](https://github.com/OfficeDev/office-ui-fabric-js/blob/master/ghdocs/components/Overlay.md)|Recommended|
|[Panel](https://github.com/OfficeDev/office-ui-fabric-js/blob/master/ghdocs/components/Panel.md)|Recommended|
|People Picker|Not recommended. Consider using [Fabric React’s PeoplePicker](https://dev.office.com/fabric#/components/peoplepicker)|
|Persona Card|Not recommended|
|[Pivot](https://github.com/OfficeDev/office-ui-fabric-js/blob/master/ghdocs/components/Pivot.md)|Recommended|
|[ProgressIndicator](https://github.com/OfficeDev/office-ui-fabric-js/blob/master/ghdocs/components/ProgressIndicator.md)|Recommended|
|[Searchbox](https://github.com/OfficeDev/office-ui-fabric-js/blob/master/ghdocs/components/SearchBox.md)|Recommended|
|[Spinner](https://github.com/OfficeDev/office-ui-fabric-js/blob/master/ghdocs/components/Spinner.md)|Recommended|
|[Table](https://github.com/OfficeDev/office-ui-fabric-js/blob/master/ghdocs/components/Table.md)|Recommended|
|[TextField](https://github.com/OfficeDev/office-ui-fabric-js/blob/master/ghdocs/components/TextField.md)|Recommended|
|[Toggle](https://github.com/OfficeDev/office-ui-fabric-js/blob/master/ghdocs/components/Toggle.md)|Recommended|
   
##Next steps
If you're looking for end-to-end code samples that show you how to use Fabric, or Fabric 2.6.1 documentation, we've got you covered. See the following resources:

- [Excel Sales Tracker (uses Fabric JS)](https://github.com/OfficeDev/Excel-Add-in-JavaScript-SalesTracker). 
- [UX design patterns (uses Fabric 2.6.1)](https://github.com/OfficeDev/Office-Add-in-UX-Design-Patterns-Code). 
- [Office Add-in Fabric UI sample (uses Fabric 1.0)](https://github.com/OfficeDev/Office-Add-in-Fabric-UI-Sample). 
- [Using Fabric 2.6.1 in an Office Add-in](https://dev.office.com/docs/add-ins/design/ui-elements/using-office-ui-fabric). 

 

