---
title: Show or hide controls on a custom tab
description: Learn how to programmatically show or hide buttons, menus, and groups on a custom ribbon tab in your Office Add-in.
ms.date: 08/19/2026
ms.topic: how-to
ms.localizationpriority: medium
---

# Show or hide add-in commands on a custom tab

Dynamic ribbon visibility helps declutter the ribbon. Showing only the add-in commands that matter to each user makes those commands easier to find. For example, after a user signs in, your add-in can show only the functionality that applies to that user.

> [!NOTE]
> This article shows how to manage the visibility of add-in commands on a custom tab. To learn how to programmatically enable or disable commands without hiding them from the ribbon, see [Change the availability of add-in commands](disable-add-in-commands.md).

## Supported Office applications and requirement set

Setting the visibility of add-in commands on custom tabs requires [RibbonApi 1.3](/javascript/api/requirement-sets/common/ribbon-api-requirement-sets) support. This feature is supported in **Excel**, **PowerPoint**, and **Word**.

## Supported controls

The following table lists the ribbon controls whose visibility can be changed on a custom tab.

| Ribbon controls | Support |
| ---- | ---- |
| **Buttons** | Supported |
| **Groups** | Supported |
| **Menus** | Supported |
| **Menu items** | Not supported |

> [!NOTE]
> To configure a custom tab to only show in certain contexts, see [Create custom contextual tabs in Office Add-ins](contextual-tabs.md).

## Try out a completed add-in sample

To test the visibility of controls on a ribbon, try out the [Show or hide controls on a custom ribbon tab sample](https://github.com/OfficeDev/Office-Add-in-samples/tree/main/Samples/office-visibility).

## Configure a shared runtime

To programmatically change the visibility of a control or group, your add-in must use a [shared runtime](../testing/runtimes.md#shared-runtime). For guidance, see [Configure your Office Add-in to use a shared runtime](../develop/configure-your-add-in-to-use-a-shared-runtime.md).

## Set the initial visibility in the manifest

> [!NOTE]
> Initial visibility configuration is only supported in an add-in that uses the unified manifest for Microsoft 365.

By default, buttons, groups, and menus on a custom tab are visible when the Office application starts. To initially hide a control, set its `"visible"` property to `false`. The location of the `"visible"` property in the manifest depends on the control.

- **Button** or **menu**: Specified in the applicable button or menu object in the ["extensions.ribbons.tabs.groups.controls"](/microsoft-365/extensibility/schema/extension-common-custom-group-controls-item#visible) array.
- **Group**: Specified in the applicable group object in the ["extensions.ribbons.tabs.groups"](/microsoft-365/extensibility/schema/extension-ribbons-custom-tab-groups-item#visible) array.

The following example configures a sample **Reporting** group and its **View report** button to be visible when the add-in starts. The **Export report** button in the same group is initially hidden.

```json
"extensions": [
    {
        "ribbons": [
            {
                "tabs": [
                    {
                        "id": "Contoso.UserToolsTab",
                        "label": "User tools",
                        "groups": [
                            {
                                "id": "Contoso.ReportingGroup",
                                "label": "Reporting",
                                "controls": [
                                    {
                                        "id": "Contoso.ViewReportButton",
                                        "type": "button",
                                        "label": "View report",
                                        "icons": [
                                          {
                                            "size": 16,
                                            "url": "icon_16.png"
                                          },
                                          {
                                            "size": 32,
                                            "url": "icon_32.png"
                                          },
                                          {
                                            "size": 80,
                                            "url": "icon_80.png"
                                          }
                                        ],
                                        "supertip": {
                                            "title": "View report",
                                            "description": "View report"
                                        },
                                        "actionId": "viewReport",
                                        "visible": true
                                    },
                                    {
                                        "id": "Contoso.ExportReportButton",
                                        "type": "button",
                                        "label": "Export report",
                                        "icons": [
                                          {
                                            "size": 16,
                                            "url": "icon_16.png"
                                          },
                                          {
                                            "size": 32,
                                            "url": "icon_32.png"
                                          },
                                          {
                                            "size": 80,
                                            "url": "icon_80.png"
                                          }
                                        ],
                                        "supertip": {
                                            "title": "Export report",
                                            "description": "Export report"
                                        },
                                        "actionId": "exportReport",
                                        "visible": false
                                    }
                                ],
                                "visible": true
                            }
                        ]
                    }
                ]
            }
        ]
    }
]
```

## Programmatically change visibility

To change the visibility of a button, menu, or group at runtime, create a [RibbonUpdaterData](/javascript/api/office/office.ribbonupdaterdata) object that specifies the following.

- The IDs of the control and its parent group and tab, as applicable. The IDs must match those declared in the manifest.
- The visibility of the control.

Then, pass the **RibbonUpdaterData** object to the [Office.ribbon.requestUpdate](/javascript/api/office/office.ribbon#office-office-ribbon-requestupdate-member(1)) method.

> [!TIP]
> At startup, avoid calling **requestUpdate** if the manifest already specifies the visibility you need. To change the visibility of multiple controls, include them in one **RibbonUpdaterData** object and call **requestUpdate** once. This practice helps prevent the ribbon from flickering.

### Show or hide a button or menu

To set the visibility of a button or menu, configure the [`Office.Control.visible`](/javascript/api/office/office.control#office-office-control-visible-member) property. The following example shows a button. The same pattern applies to a menu control.

```typescript
async function setButtonVisibility(visible: boolean) {
    const button: Office.Control = {
        id: "Contoso.ExportReportButton",
        visible: visible
    };
    const group: Office.Group = {
        id: "Contoso.ReportingGroup",
        controls: [button]
    };
    const tab: Office.Tab = {
        id: "Contoso.UserToolsTab",
        groups: [group]
    };
    const ribbonUpdater: Office.RibbonUpdaterData = { tabs: [tab] };

    await Office.ribbon.requestUpdate(ribbonUpdater);
}
```

### Show or hide a group

Use the [`Office.Group.visible`](/javascript/api/office/office.group#office-office-group-visible-member) property to set the visibility of a group. The following example shows or hides a group and all the controls it contains.

```typescript
async function setGroupVisibility(visible: boolean) {
    const group: Office.Group = {
        id: "Contoso.ReportingGroup",
        visible: visible
    };
    const tab: Office.Tab = { id: "Contoso.UserToolsTab", groups: [group] };
    const ribbonUpdater: Office.RibbonUpdaterData = { tabs: [tab] };

    await Office.ribbon.requestUpdate(ribbonUpdater);
}
```

> [!TIP]
> The Microsoft 365 application controls when the ribbon is updated. The **requestUpdate** method queues an update request and resolves its `Promise` as soon as the request is queued, not when the ribbon is updated.

## See also

- [Add-in commands](add-in-commands.md)
- [Change the availability of add-in commands](disable-add-in-commands.md)
- [Create custom contextual tabs in Office Add-ins](contextual-tabs.md)
