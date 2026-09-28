---
title: Fabric Core in Office Add-ins
description: Learn how to use Fabric Core CSS classes for icons, colors, typography, and responsive grids in non-React Office Add-ins.
ms.date: 09/28/2026
ms.topic: overview
ms.localizationpriority: medium
ai-usage: ai-assisted
---

# Fabric Core in Office Add-ins

If you're building a non-React Office Add-in, Fabric Core gives you ready-made CSS classes and Sass mixins for icons, colors, typography, and responsive grids that are all aligned with the Fluent UI design language. Because it's framework-independent, you can use Fabric Core with any single-page application or server-side web UI framework. (It's called "Fabric Core" instead of "Fluent Core" for historical reasons.)

> [!NOTE]
> This article describes the use of Fabric Core in the context of Office Add-ins, but it's also used in a wide range of Microsoft 365 apps and extensions. For more information, see [Fabric Core](https://developer.microsoft.com/fluentui#/get-started/web#fabric-core) and the open source repo [Office UI Fabric Core](https://github.com/OfficeDev/office-ui-fabric-core).

[!INCLUDE [alert-fluent-ui-web-components](../includes/alert-fluent-ui-web-components.md)]

## Use Fabric Core: icons, fonts, colors

To get started, add the Fabric Core stylesheet to your add-in, and then apply CSS classes for icons, fonts, and colors.

### Add the stylesheet

Add the Fabric Core content delivery network (CDN) reference to the `<head>` of your HTML page.

```html
<link rel="stylesheet" href="https://res-1.cdn.office.net/files/fabric-cdn-prod_20230815.002/office-ui-fabric-core/11.0.0/css/fabric.min.css">
```

### Use icons

To use a Fabric Core icon, add an `<i>` element and apply the appropriate CSS classes. You control the icon's size with font-size classes and its color with color classes. The following example renders an extra-large table icon in the theme primary color (#0078d7).

```html
<i class="ms-Icon ms-font-xl ms-Icon--Table ms-fontColor-themePrimary"></i>
```

The class pattern is `ms-Icon--<IconName>`. To browse available icons and find their names, see [Fluent UI Icons](https://developer.microsoft.com/fluentui#/styles/web/icons). When you find an icon to use, prefix the icon name with `ms-Icon--`.

### Use fonts and colors

Fabric Core provides CSS classes for Segoe UI font sizes (`ms-font-s`, `ms-font-m`, `ms-font-l`, `ms-font-xl`, and others) and theme-aware colors (`ms-fontColor-themePrimary`, `ms-fontColor-neutralSecondary`, and others). The following example applies a large font size and the primary theme color to a heading.

```html
<h2 class="ms-font-l ms-fontColor-themePrimary">Sales by region</h2>
<p class="ms-font-m ms-fontColor-neutralSecondary">Updated quarterly.</p>
```

For the full list of available sizes, see [Typography](https://developer.microsoft.com/fluentui#/styles/web/typography). For theme-aware and neutral colors, see [Colors](https://developer.microsoft.com/fluentui#/styles/web/colors).

## Use Office UI Fabric JS components

If your non-React add-in needs interactive UI components&mdash;buttons, dialogs, dropdowns, pickers, and more&mdash;you can use [Office UI Fabric JS](https://github.com/OfficeDev/office-ui-fabric-js). This library provides pre-built, styled components that match the Office design language without requiring React.

To use a component, reference the Fabric JS script and stylesheet in your HTML, then initialize the component in JavaScript. See the [repository's README](https://github.com/OfficeDev/office-ui-fabric-js#office-ui-fabric-js) for setup instructions and the full component list.

## See also

- [Design the UI of Office Add-ins](add-in-design.md)
- [Use Fluent UI React in Office Add-ins](../quickstarts/fluent-react-quickstart.md)
- [Color in Office Add-ins](add-in-color.md)
- [Typography in Office Add-ins](add-in-typography.md)
- [Icons in Office Add-ins](add-in-icons.md)
- [Layout for Office Add-ins](add-in-layout.md)
