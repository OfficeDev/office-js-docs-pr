---
title: Office UI Fabric in Office Add-ins 
description: 'Get an overview of how to use the Office UI Fabric components in Office Add-ins.'
ms.date: 05/03/2021
localization_priority: Normal
---

# Office UI Fabric in Office Add-ins

Office UI Fabric is a JavaScript front-end framework for building user experiences for Office. Fabric provides visuals-focused components that you can extend, rework, and use in your Office Add-in. Because Fabric uses the Office Design Language, Fabric's UX components look like a natural extension of Office.

If you are building an add-in, we encourage you to use Office UI Fabric to create your user experience. Using Office UI Fabric is optional.

The following sections explain how to get started using Fabric to meet your requirements.

## Use Fabric Core: icons, fonts, colors

Fabric Core contains basic elements of the design language such as icons, colors, type, and grid. Fabric core is framework independent. Fabric Core is used by and included with Fabric React.

To get started using Fabric Core:

1. Add the CDN reference to the HTML on your page.  

    ```html
    <link rel="stylesheet" href="https://static2.sharepointonline.com/files/fabric/office-ui-fabric-core/9.6.1/css/fabric.min.css">
    ```

2. Use Fabric icons and fonts.

    To use a Fabric icon, include the "i" element on your page, and then reference the appropriate classes. You can control the size of the icon by changing the font size. For example, the following code shows how to make an extra-large table icon that uses the themePrimary (#0078d7) color.

    ```html
    <i class="ms-Icon ms-font-xl ms-Icon--Table ms-fontColor-themePrimary"></i>
    ```

    To find more icons that are available in Office UI Fabric, use the search feature on the [Icons](https://developer.microsoft.com/fabric#/styles/icons) page. When you find an icon to use in your add-in, be sure to prefix the icon name with `ms-Icon--`.

    For information about font sizes and colors that are available in Office UI Fabric, see [Typography](https://developer.microsoft.com/fabric#/styles/typography) and [Colors](https://developer.microsoft.com/fabric#/styles/colors).

## Use Fabric Components

Fabric provides a variety of UX components that you can use to build your add-in. We do not expect that all fabric components will be used by a single add-in. Determine the best components for your scenario and user experience (for example, it may be hard to properly display a [Breadcrumb](https://developer.microsoft.com/fabric#/components/breadcrumb) in the task pane).

The following is a list of common [Fabric React UX components](https://developer.microsoft.com/fluentui#/controls/web) that we recommend for use in an add-in:

- [Button](https://developer.microsoft.com/fabric#/components/button)
- [Checkbox](https://developer.microsoft.com/fabric#/components/checkbox)
- [ChoiceGroup](https://developer.microsoft.com/fabric#/components/choicegroup)
- [Dropdown](https://developer.microsoft.com/fabric#/components/dropdown)
- [Label](https://developer.microsoft.com/fabric#/components/label)
- [List](https://developer.microsoft.com/fabric#/components/list)
- [Pivot](https://developer.microsoft.com/fabric#/components/pivot)
- [TextField](https://developer.microsoft.com/fabric#/components/textfield)
- [Toggle](https://developer.microsoft.com/fabric#/components/toggle)

You can use different JavaScript frameworks, such as Angular or React, to build your add-in. To get started using Fabric components with your framework, see the following resources.

|**Framework**|**Example**|
|:------------|:----------|
|**React**|[Using Office UI Fabric React in Office Add-ins](using-office-ui-fabric-react.md )|
