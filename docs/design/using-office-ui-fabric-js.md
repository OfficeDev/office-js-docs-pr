
# Use Office UI Fabric JS in Office Add-ins

Office UI Fabric is a JavaScript front-end framework for building user experiences for Office and Office 365. If you build an add-in using JavaScript only, without using a framework like Angular or React, consider using Fabric JS to create your user experience. For more information, see [Office UI Fabric JS](https://dev.office.com/fabric-js).

The following steps walk you through the basics of using Fabric JS.  

## 1. Add the Fabric CDN references
To reference Fabric from the CDN, add the following HTML code to your page.

    <link rel="stylesheet" href="https://static2.sharepointonline.com/files/fabric/office-ui-fabric-js/1.4.0/css/fabric.min.css">
    <link rel="stylesheet" href="https://static2.sharepointonline.com/files/fabric/office-ui-fabric-js/1.4.0/css/fabric.components.min.css">
    <script src="https://static2.sharepointonline.com/files/fabric/office-ui-fabric-js/1.4.0/js/fabric.min.js"></script>

## 2. Use Fabric JS UX components

Fabric JS provides several UX components, like buttons or checkboxes, that you can use in your add-in. The following is a list of the Fabric JS UX components that we recommend for use in an add-in. To use one of the Fabric components in your add-in, follow the link to the Fabric documentation, and then follow the instructions in **Using this component**. 

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

## Next steps
If you're looking for an end-to-end code sample that shows you how to use Fabric JS, we've got you covered. See the following resource:

- [Excel Sales Tracker](https://github.com/OfficeDev/Excel-Add-in-JavaScript-SalesTracker) 

## Related resources
If you're looking for code samples or documentation on a previous release of Fabric, see the following:

- [UX design patterns (uses Fabric 2.6.1)](https://github.com/OfficeDev/Office-Add-in-UX-Design-Patterns-Code) 
- [Office Add-in Fabric UI sample (uses Fabric 1.0)](https://github.com/OfficeDev/Office-Add-in-Fabric-UI-Sample) 
- [Using Fabric 2.6.1 in an Office Add-in](https://dev.office.com/docs/add-ins/design/ui-elements/using-office-ui-fabric)
 

