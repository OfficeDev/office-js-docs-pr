---
title: Fabric Core in Office Add-ins 
description: Get an overview of how to use Fabric Core, and Fabric UI components in Office Add-ins.
ms.date: 10/18/2022
ms.topic: overview
ms.localizationpriority: medium
---

# Fabric Core in Office Add-ins

Fabric Core is an open-source collection of CSS classes and Sass mixins that's *intended for use in non-React* Office Add-ins. Fabric Core contains basic elements of the Fluent UI design language such as icons, colors, typefaces, and grids. Fabric Core is framework independent, so it can be used with any single-page application or any server-side web UI framework. (It's called "Fabric Core" instead of "Fluent Core" for historical reasons.)

If your add-in's UI isn't React-based, you can also make use of a set of non-React components. See [Use Office UI Fabric JS components](#use-office-ui-fabric-js-components).

> [!NOTE]
> This article describes the use of Fabric Core in the context of Office Add-ins, but it's also used in a wide range of Microsoft 365 apps and extensions. For more information, see [Fabric Core](https://developer.microsoft.com/fluentui#/get-started/web#fabric-core) and the open source repo [Office UI Fabric Core](https://github.com/OfficeDev/office-ui-fabric-core).

[!INCLUDE [alert-fluent-ui-web-components](../includes/alert-fluent-ui-web-components.md)]

## Use Fabric Core: icons, fonts, colors

1. Add the content delivery network (CDN) reference to the HTML on your page.

    ```html
    <link rel="stylesheet" href="https://res-1.cdn.office.net/files/fabric-cdn-prod_20230815.002/office-ui-fabric-core/11.0.0/css/fabric.min.css">
    ```

2. Use Fabric Core icons and fonts.

    To use a Fabric Core icon, include the "i" element on your page, and then reference the appropriate classes. You can control the size of the icon by changing the font size. For example, the following code shows how to make an extra-large table icon that uses the themePrimary (#0078d7) color.

    ```html
    <i class="ms-Icon ms-font-xl ms-Icon--Table ms-fontColor-themePrimary"></i>
    ```

    For more detailed instructions, see [Fluent UI Icons](https://developer.microsoft.com/fluentui#/styles/web/icons). To find more icons that are available in Fabric Core, use the search feature on that page. When you find an icon to use in your add-in, be sure to prefix the icon name with `ms-Icon--`.

    For information about font sizes and colors that are available in Fabric Core, see [Typography](https://developer.microsoft.com/fluentui#/styles/web/typography) and the **Colors** table of contents at [Colors](https://developer.microsoft.com/fluentui#/styles/web/colors).

Examples are included in the [Samples](#samples) later in this article.

## Use Office UI Fabric JS components

Add-ins with non-React UIs can also use any of the many components from [Office UI Fabric JS](https://github.com/OfficeDev/office-ui-fabric-js), including buttons, dialogs, pickers, and more. See the readme of the repo for instructions.

Examples are included in the [Samples](#samples) later in this article.

## Samples

The following sample add-ins use Fabric Core and/or Office UI Fabric JS components. Some of these repos are archived, meaning that they are no longer being updated with bug or security fixes, but you can still use them to learn how to use Fabric Core and Fabric UI components.

- [Excel Add-in JavaScript SalesTracker](https://github.com/OfficeDev/Excel-Add-in-JavaScript-SalesTracker)
- [Excel Add-in SalesLeads](https://github.com/OfficeDev/Excel-Add-in-SalesLeads)
- [Excel Add-in WoodGrove Expense Trends](https://github.com/OfficeDev/Excel-Add-in-WoodGrove-Expense-Trends)
- [Excel Content Add-in Humongous Insurance](https://github.com/OfficeDev/Excel-Content-Add-in-Humongous-Insurance)
- [Office Add-in Fabric UI Sample](https://github.com/OfficeDev/Office-Add-in-Fabric-UI-Sample)
- [Office-Add-in-UX-Design-Patterns-Code](https://github.com/OfficeDev/Office-Add-in-UX-Design-Patterns-Code)
- [Outlook Add-in GifMe](https://github.com/OfficeDev/Outlook-Add-in-GifMe)
- [PowerPoint Add-in Microsoft Graph ASPNET InsertChart](https://github.com/OfficeDev/PowerPoint-Add-in-Microsoft-Graph-ASPNET-InsertChart)
- [Word Add-in Angular2 StyleChecker](https://github.com/OfficeDev/Word-Add-in-Angular2-StyleChecker)
- [Word Add-in JS Redact](https://github.com/OfficeDev/Word-Add-in-JS-Redact)
- [Word Add-in MarkdownConversion](https://github.com/OfficeDev/Word-Add-in-MarkdownConversion)
