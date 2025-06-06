---
title: Integrate built-in Office buttons into custom control groups and tabs
description: Learn how to include built-in Office buttons in your custom command groups and tabs on the Office ribbon.
ms.date: 06/10/2025
ms.topic: how-to
ms.localizationpriority: medium
---


# Integrate built-in Office buttons into custom control groups and tabs

You can insert built-in Office buttons into your custom control groups on the Office ribbon by using markup in the add-in's manifest. (You can't insert your custom add-in commands into a built-in Office group.) You can also insert entire built-in Office control groups into your custom ribbon tabs.

> [!NOTE]
> This article assumes that you're familiar with the article [Basic concepts for add-in commands](add-in-commands.md). Please review it if you haven't done so recently.

> [!IMPORTANT]
> The add-in feature described in this article is only available in **PowerPoint** on the web, on Windows, and on Mac.

Open the tab for the type of manifest your add-in uses for the details of the manifest markup.

# [Unified manifest for Microsoft 365](#tab/jsonmanifest)

[!include[Unified manifest host application support note](../includes/unified-manifest-support-note.md)]

## Insert a built-in control group into a custom tab

To insert a built-in Office control group into a custom tab, add a group object with a [`"builtInGroupId"`](/microsoft-365/extensibility/schema/extension-ribbons-array-tabs-item#builtintabid) property *instead of an [`"id"`](/microsoft-365/extensibility/schema/extension-ribbons-array-tabs-item#id) property* to the [`"groups"`](/microsoft-365/extensibility/schema/extension-ribbons-array-tabs-item#groups) array of your custom tab object. Set to the ID of the built-in group. See [Find the IDs of controls and control groups](#find-the-ids-of-controls-and-control-groups). *The built-in group object should have no other properties.*

The following example adds the Office Paragraph control group to a custom tab.

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
                        "groups": [
                            ... // Optionally, other groups in the tab.
                            {
                                "builtInGroupId": "Paragraph"
                            },
                            ... // Optionally, other groups in the tab.
                        ]
                    }
                ]
            }
        ]
    }
]
```

## Insert a built-in control into a custom group

To insert a built-in Office control into a custom group, add a control object with a `"builtInControlId"` property *instead of an `"id"` property* to the [`"controls"`](/microsoft-365/extensibility/schema/extension-ribbons-custom-tab-groups-item#controls) array of your custom group object. Set to the ID of the built-in control. See [Find the IDs of controls and control groups](#find-the-ids-of-controls-and-control-groups). *The built-in control object should have no other properties.*

The following example adds the Office Superscript control to a custom group.

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
                        ...
                        "groups": [
                            {
                                "id": "MyGroup",
                                ...
                                "controls": [
                                    ... // Optionally, other controls in the group.
                                    {
                                        "builtInControlId": "Superscript"
                                    },
                                    ... // Optionally, other controls in the group.
                                ]
                            }
                        ]
                    }
                ]
            }
        ]
    }
]
```

# [Add-in only manifest](#tab/xmlmanifest)

## Insert a built-in control group into a custom tab

To insert a built-in Office control group into a tab, add an [OfficeGroup](/javascript/api/manifest/customtab#officegroup) element as a child element in the parent **\<CustomTab\>** element. The `id` attribute of the of the **\<OfficeGroup\>** element is set to the ID of the built-in group. See [Find the IDs of controls and control groups](#find-the-ids-of-controls-and-control-groups).

The following markup example adds the Office Paragraph control group to a custom tab and positions it to appear just after a custom group.

```xml
<ExtensionPoint xsi:type="PrimaryCommandSurface">
  <CustomTab id="Contoso.TabCustom1">
    <Group id="Contoso.myCustomTab.group1">
       <!-- Additional markup omitted. -->
    </Group>
    <OfficeGroup id="Paragraph" />
    <Label resid="customTabLabel1" />
  </CustomTab>
</ExtensionPoint>
```

## Insert a built-in control into a custom group

To insert a built-in Office control into a custom group, add an [OfficeControl](/javascript/api/manifest/group#officecontrol) element as a child element in the parent **\<Group\>** element. The `id` attribute of the **\<OfficeControl\>** element is set to the ID of the built-in control. See [Find the IDs of controls and control groups](#find-the-ids-of-controls-and-control-groups).

The following markup example adds the Office Superscript control to a custom group and positions it to appear just after a custom button.

```xml
<ExtensionPoint xsi:type="PrimaryCommandSurface">
  <CustomTab id="Contoso.TabCustom2">
    <Group id="Contoso.TabCustom2.group1">
        <Label resid="residCustomTabGroupLabel"/>
        <Icon>
            <bt:Image size="16" resid="blue-icon-16" />
            <bt:Image size="32" resid="blue-icon-32" />
            <bt:Image size="80" resid="blue-icon-80" />
        </Icon>
        <Control xsi:type="Button" id="Contoso.Button1">
            <!-- Information on the control omitted. -->
        </Control>
        <OfficeControl id="Superscript" />
        <!-- Other controls, as needed. -->
    </Group>
    <Label resid="customTabLabel1" />
  </CustomTab>
</ExtensionPoint>
```

---

> [!NOTE]
> Users can customize the ribbon in the Office application. Any user customizations will override your manifest settings. For example, a user can remove a button from any group and remove any group from a tab.

## Find the IDs of controls and control groups

The IDs for supported controls and control groups are in files in the repo [Office Control IDs](https://github.com/OfficeDev/office-control-ids). Follow the instructions in the ReadMe file of that repo.

## Behavior on unsupported platforms

If your add-in is installed on a platform that doesn't support [requirement set AddinCommands 1.3](/javascript/api/requirement-sets/common/add-in-commands-requirement-sets), then the markup described in this article is ignored and the built-in Office controls/groups won't appear in your custom groups/tabs. To prevent your add-in from being installed on platforms that don't support the markup, you must specify **AddinCommands 1.3** in the manifest as a requirement for installation. For instructions, see [Specify which Office versions and platforms can host your add-in](../develop/specify-office-hosts-and-api-requirements.md#specify-which-office-versions-and-platforms-can-host-your-add-in). Alternatively, design your add-in to have an experience when **AddinCommands 1.3** isn't supported, as described in [Design for alternate experiences](../develop/specify-office-hosts-and-api-requirements.md#design-for-alternate-experiences). For example, if your add-in contains instructions that assume the built-in buttons are in your custom groups, you could design a version that assumes that the built-in buttons are only in their usual places.
