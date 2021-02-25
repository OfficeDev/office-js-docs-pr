---
title: Position a custom tab on the ribbon
description: 'Learn how to control where a custom tab appears on the Office ribbon and whether it has focus by default.'
ms.date: 02/25/2021
localization_priority: Normal
---


# Position a custom tab on the ribbon

You can specify where you want your add-in's custom tab to appear on the Office application's ribbon by using markup in the add-in's manifest.

> [!NOTE]
> This article assumes that you are familiar with the article [Basic concepts for add-in commands](add-in-commands.md). Please review it if you have not done so recently.

> [!IMPORTANT]
>
> - The add-in feature and markup described in this article is *only available in PowerPoint on the web*.
> - The markup described in this article only works on platforms that support requirement set **AddinCommands 1.3**. See [Behavior on unsupported platforms](#behavior-on-unsupported-platforms) below.

Specify where you want a custom tab to appear by identifying which built-in Office tab you want it to be next to and specifying whether it should be on the left or right side of the built-in tab. Make these specifications by including either an [InsertBefore](../reference/manifest/customtab.md#insertbefore) (left) or an [InsertAfter](../reference/manifest/customtab.md#insertafter) (right) element in the [CustomTab](../reference/manifest/customtab.md) element of your add-in's manifest. (You cannot have both elements.)

In the following example, the custom tab is configured to appear *just after* the **Review** tab. Note that the value of the `<InsertAfter>` element is the ID of the built-in Office tab. 

```xml
<ExtensionPoint xsi:type="ContosoRibbonTab">
  <CustomTab id="TabCustom1">
    <Group id="myCustomTab.group1">
       <!-- additional markup omitted -->
    </Group>
    <Label resid="customTabLabel1" />
    <InsertAfter>TabReview</InsertAfter>
  </CustomTab>
</ExtensionPoint>
```

Keep the following points in mind.

- The  `<InsertBefore>` and  `<InsertAfter>` elements are optional. If you use neither, then your custom tab will appear as the rightmost tab on the ribbon.
- The  `<InsertBefore>` and  `<InsertAfter>` elements are mutually exclusive. You cannot use both.
- If the user installs more than one add-in whose custom tab is configured for the same place, say after the **Review** tab, then the tab for the most recently installed add-in will be located in that place. The tabs of the previously installed add-ins will be moved over one place. For example, the user installs add-ins A, B, and C in that order and all are configured to insert a tab after the **Review** tab, then the tabs will appear in this order: **Review**, **AddinCTab**, **AddinBTab**, **AddinATab**.
- Users can customize the ribbon in the Office application. For example, a user can move or hide your add-in's tab. You cannot prevent this or detect that it has happened.
- If a user moves one of the built-in tabs, then Office interprets the `<InsertBefore>` and  `<InsertAfter>` elements in terms of *the default location of the built-in tab*. For example, if the user moves the **Review** tab to the right end of the ribbon, Office will interpret the markup in the example above as meaning "put the custom tab just to the right of *where the **Review** tab would be by default*."

## Specifying which tab has focus when the document opens

Office always gives default focus to the tab that is immediately to the right of the **File** tab. By default this is the **Home** tab. If you configure your custom tab to be before the **Home** tab, with `<InsertBefore>TabHome</InsertBefore>`, then your custom tab will have focus when the document opens.

> [!IMPORTANT]
> Giving excessive prominence to your add-in inconveniences and annoys users and administrators. Do not position a custom tab before the **Home** tab unless your add-in is the primary way users will interact with the document.

## Behavior on unsupported platforms

If your add-in is installed on a platform that doesn't support [requirement set AddinCommands 1.3](../reference/requirement-sets/add-in-commands-requirement-sets.md), then the markup described in this article is ignored and your custom tab will appear as the rightmost tab on the ribbon. To prevent your add-in from being installed on platforms that don't support the markup, add a reference to the requirement set in the `<Requirements>` section of the manifest. For instructions, see [Set the Requirements element in the manifest](../develop/specify-office-hosts-and-api-requirements.md#set-the-requirements-element-in-the-manifest). Alternatively, you can design your add-in to have an alternate experience when **AddinCommands 1.3** is not supported, as described in [Use runtime checks in your JavaScript code](../develop/specify-office-hosts-and-api-requirements.md#use-runtime-checks-in-your-javascript-code). For example, if your add-in contains instructions that assume the custom tab is where you want it, you could have an alternate version that assumes the tab is the rightmost.
