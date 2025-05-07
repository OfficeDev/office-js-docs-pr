---
title: Position a custom tab on the ribbon
description: Learn how to control where a custom tab appears on the Office ribbon and whether it has focus by default.
ms.date: 02/12/2025
ms-topic: how-to
ms.localizationpriority: medium
---


# Position a custom tab on the ribbon

You can specify where you want your add-in's custom tab to appear on the Office application's ribbon by using markup in the add-in's manifest.

> [!NOTE]
> This article assumes that you're familiar with the article [Basic concepts for add-in commands](add-in-commands.md). Please review it if you haven't done so recently.

> [!IMPORTANT]
>
> - The add-in feature and markup described in this article is *only available in PowerPoint on the web*.
> - The markup described in this article only works on platforms that support requirement set **AddinCommands 1.3**. See [Behavior on unsupported platforms](#behavior-on-unsupported-platforms) below.

By default, a custom tab is added to the end of the ribbon. However, you can specify where you want a custom tab to appear by identifying which built-in Office tab you want it to be next to and specifying whether it should be on the left or right side of the built-in tab. Open the tab for the type of manifest your add-in uses for the details of the manifest markup.

# [Unified manifest for Microsoft 365](#tab/jsonmanifest)

[!include[Unified manifest host application support note](../includes/unified-manifest-support-note.md)]

To position your custom tab, include a [`"position"`](/microsoft-365/extensibility/schema/extension-ribbons-array-tabs-item-position) property in the [`"extensions.ribbons.tabs"`](/microsoft-365/extensibility/schema/extension-ribbons-array#tabs) object. Set the `"position.builtInTabId"` property to the ID of the built-in tab that you want your custom tab to be next to. (See [Find the IDs of built-in Office ribbon tabs](../develop/built-in-ui-ids.md).) Set the `"position.align"` property to either `"before"` (left) or `"after"` (right).

In the following example, the custom tab is configured to appear *just after* the **Review** tab.

```json
"extensions": [
    ...
    {
        ...
        "ribbons": [
            ...
            {
                ...
                "tabs": [
                    {
                        "id": "MyTab",
                        ...
                        "position": {
                            "builtInTabId": "TabReview",
                            "align": "after"
                        }
                    }
                ]
            }
        ]
    }
]
```

# [Add-in only manifest](#tab/xmlmanifest)

To position your custom tab, include either an [InsertBefore](/javascript/api/manifest/customtab#insertbefore) (left) or an [InsertAfter](/javascript/api/manifest/customtab#insertafter) (right) element in the [CustomTab](/javascript/api/manifest/customtab) element of your add-in's manifest. (You cannot have both elements.)

In the following example, the custom tab is configured to appear *just after* the **Review** tab. Note that the value of the **\<InsertAfter\>** element is the ID of the built-in Office tab. See [Find the IDs of built-in Office ribbon tabs](../develop/built-in-ui-ids.md).

```xml
<ExtensionPoint xsi:type="ContosoRibbonTab">
  <CustomTab id="Contoso.TabCustom2">
    <Group id="Contoso.TabCustom2.group2">
       <!-- additional markup omitted -->
    </Group>
    <Label resid="customTabLabel1" />
    <InsertAfter>TabReview</InsertAfter>
  </CustomTab>
</ExtensionPoint>
```

---

## How user actions can affect custom tab positioning

- If the user installs more than one add-in whose custom tab is configured for the same place, say after the **Review** tab, then the tab for the most recently installed add-in will be located in that place. The tabs of the previously installed add-ins will be moved over one place. For example, the user installs add-ins A, B, and C in that order and all are configured to insert a tab after the **Review** tab, then the tabs will appear in this order: **Review**, **AddinCTab**, **AddinBTab**, **AddinATab**.
- Users can customize the ribbon in the Office application. For example, a user can move or hide your add-in's tab. You cannot prevent this or detect that it has happened.
- If a user moves one of the built-in tabs, then Office interprets the positioning markup in the manifest in terms of *the default location of the built-in tab*. For example, if the user moves the **Review** tab to the right end of the ribbon, Office will interpret the markup in the previous example as meaning "put the custom tab just to the right of *where the **Review** tab would be by default*."

## Specify which tab has focus when the document opens

Office always gives default focus to the tab that's immediately to the right of the **File** tab. By default this is the **Home** tab. If you configure your custom tab to be before the **Home** tab, then your custom tab will have focus when the document opens.

> [!IMPORTANT]
> Giving excessive prominence to your add-in inconveniences and annoys users and administrators. Don't position a custom tab before the **Home** tab unless your add-in is the primary way users will interact with the document.

## Behavior on unsupported platforms

If your add-in is installed on a platform that doesn't support [requirement set AddinCommands 1.3](/javascript/api/requirement-sets/common/add-in-commands-requirement-sets), then the markup described in this article is ignored and your custom tab will appear as the rightmost tab on the ribbon. To prevent your add-in from being installed on platforms that don't support the markup, you must specify **AddinCommands 1.3** in the manifest as a requirement for installation. For instructions, see [Specify which Office versions and platforms can host your add-in](../develop/specify-office-hosts-and-api-requirements.md#specify-which-office-versions-and-platforms-can-host-your-add-in). Alternatively, design your add-in to have an alternate experience when **AddinCommands 1.3** isn't supported, as described in [Design for alternate experiences](../develop/specify-office-hosts-and-api-requirements.md#design-for-alternate-experiences). For example, if your add-in contains instructions that assume the custom tab is where you want it, you could have an alternate version that assumes the tab is the rightmost.
