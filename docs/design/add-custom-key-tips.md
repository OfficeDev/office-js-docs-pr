---
title: Custom KeyTips for Office Add-ins
description: Learn how to add custom KeyTips, also known as sequential key shortcuts or access keys, to your Office Add-in.
ms.date: 05/27/2026
ms.topic: how-to
ms.localizationpriority: medium
---

# Add custom KeyTips to your Office Add-ins

KeyTips, also known as sequential key shortcuts or access keys, provide an efficient keyboard navigation method for your add-in's users. Unlike simultaneous keyboard shortcuts (such as <kbd>Ctrl</kbd>+<kbd>S</kbd>) that are used to run operations within an add-in, KeyTips are pressed in sequence to quickly access options from the Office ribbon. KeyTips help users:

- Quickly navigate the ribbon and access actions.
- Support keyboard-only navigation for accessibility.

> [!TIP]
> The keyboard command to access KeyTips varies depending on the platform. To familiarize yourself with KeyTips, see [Use the keyboard to work with the ribbon](https://support.microsoft.com/office/954cd3f7-2f77-4983-978d-c09b20e31f0e).

## Supported Office applications and platforms

Custom KeyTips are supported in the following Office applications and platforms.

| Office application | Office on the web | Office on Windows | Office on Mac |
| ----- | ----- | ----- | ----- |
| Excel | Supported | Supported | Supported |
| Outlook | Not supported | Not supported | Not supported |
| PowerPoint | Supported | Supported | Supported |
| Word | Supported | Supported | Supported |

> [!NOTE]
>
> Custom KeyTips are supported in Office on Windows and on Mac in specific host versions.
>
> - **Windows**: Version 2603 (Build 19822.20000) or later
> - **Mac**: Version 16.107 (Build 26030819) or later
>
> To verify the version of the host application, call [Office.context.diagnostics.version](/javascript/api/office/office.contextinformation#office-office-contextinformation-version-member).

## Supported surfaces and controls

KeyTips can be defined for the add-in's controls and the ribbon tab in which the add-in appears. The following table outlines which ribbon tabs and controls allow KeyTips for each supported Office application.

| Tab type | Excel | PowerPoint | Word |
| ----- | ----- | ----- | ----- |
| Built-in Office tabs | Supported | Supported | Supported |
| Buttons | Supported | Supported | Supported |
| Contextual tab | Not available | Not available | Not available |
| Custom tab | Supported | Supported | Supported |
| Menus | Supported | Supported | Supported |
| Menu items | Not available | Not available | Not available |

## Configure the manifest

Custom KeyTips are defined in your add-in's manifest. The following example customizes KeyTips to access the add-in and its actions from the built-in Home tab and a custom tab. Note the following about this markup.

- The `"keytip"` property defines the key sequence users press to activate a custom KeyTip. It supports up to three uppercase alphanumeric characters ("A-Z", "0-9"). The `"keytip"` property is specified for the following tabs and controls.
  - The Home tab on the ribbon ([`"extensions.ribbons.tabs"`](/microsoft-365/extensibility/schema/extension-ribbons-array-tabs-item) object whose `"builtInTabId"` property is set to `"TabHome"`). For guidance on built-in Office ribbon tabs and their IDs, see [Find the IDs of built-in Office ribbon tabs](../develop/built-in-ui-ids.md).
  - The custom tab (`"extensions.ribbons.tabs"` object whose `"id"` property is set to `"CustomTab"`).
  - The add-in's button on the ribbon ([`"extensions.ribbons.tabs.groups.controls"`](/microsoft-365/extensibility/schema/extension-common-custom-group-controls-item) object whose `"type"` property is set to `"button"`).
  - The add-in's menu control on the ribbon ([`"extensions.ribbons.tabs.groups.controls"`](/microsoft-365/extensibility/schema/extension-common-custom-group-controls-item) object whose `"type"` property is set to `"menu"`).

```json
{
    "extensions": [
        {
            ...
            "ribbons": [
                {
                    ...
                    "tabs": [
                        {
                            "builtInTabId": "TabHome",
                            "groups": [
                                {
                                    "id": "ContosoGroup",
                                    ...
                                    "controls": [
                                        {
                                            "id": "AnalyzeButton",
                                            "type": "button",
                                            "label": "Analyze Data",
                                            "icons": [
                                                {
                                                    "size": 16,
                                                    "url": "https://localhost:3000/assets/icon-16.png"
                                                },
                                                {
                                                    "size": 32,
                                                    "url": "https://localhost:3000/assets/icon-32.png"
                                                },
                                                {
                                                    "size": 80,
                                                    "url": "https://localhost:3000/assets/icon-80.png"
                                                }
                                            ],
                                            "supertip": {
                                                "title": "Analyze Data",
                                                "description": "Perform advanced data analysis."
                                            },
                                            "actionId": "analyzeData",
                                            "keytip": "AD"
                                        },
                                        {
                                            "id": "DataMenu",
                                            "type": "menu",
                                            "label": "Manage Data",
                                            "icons": [
                                                {
                                                    "size": 16,
                                                    "url": "https://localhost:3000/assets/icon-16.png"
                                                },
                                                {
                                                    "size": 32,
                                                    "url": "https://localhost:3000/assets/icon-32.png"
                                                },
                                                {
                                                    "size": 80,
                                                    "url": "https://localhost:3000/assets/icon-80.png"
                                                }
                                            ],
                                            "supertip": {
                                                "title": "Manage Data",
                                                "description": "Options to manage data."
                                            },
                                            "items": [
                                                ...
                                            ],
                                            "keytip": "MD"
                                        }
                                    ]
                                }
                            ],
                            "keytip": "CH"
                        },
                        {
                            "id": "CustomTab",
                            "label": "Contoso custom tab",
                            "groups": [
                                {
                                    "id": "CustomContosoGroup",
                                    "label": "Contoso Custom Group",
                                    "icons": [
                                        {
                                            "size": 16,
                                            "url": "https://localhost:3000/assets/icon-16.png"
                                        },
                                        {
                                            "size": 32,
                                            "url": "https://localhost:3000/assets/icon-32.png"
                                        },
                                        {
                                            "size": 80,
                                            "url": "https://localhost:3000/assets/icon-80.png"
                                        }
                                    ],
                                    "controls": [
                                        {
                                            "id": "CustomTabButton",
                                            ...
                                            "keytip": "CTB"
                                        }
                                    ]
                                }
                            ],
                            "overriddenByRibbonApi": true,
                            "keytip": "CT"
                        }
                    ]
                }
            ]
        }
    ]
}
```

## Handle KeyTip conflicts

The Microsoft 365 host application checks for KeyTip conflicts between add-ins and built-in commands. If the host application detects a conflict, it automatically assigns a fallback KeyTip using the <kbd>Y</kbd> prefix followed by a number, such as <kbd>Y1</kbd>, <kbd>Y2</kbd>, or <kbd>Y3</kbd>. These fallback KeyTips ensure that each command remains uniquely accessible by keyboard.

> [!TIP]
> Choose custom KeyTips that are unlikely to overlap with built-in ribbon commands. For a list of built-in KeyTips for each Microsoft 365 host application, see the following articles.
>
> - [Excel](https://support.microsoft.com/office/1798d9d5-842a-42b8-9c99-9b7213f0040f)
> - [PowerPoint](https://support.microsoft.com/office/ebb3d20e-dcd4-444f-a38e-bb5c5ed180f4)
> - [Word](https://support.microsoft.com/office/95ef89dd-7142-4b50-afb2-f762f663ceb2)

## Behavior and limitations

When implementing KeyTips for your add-in, be aware of the following behaviors and limitations.

- Custom KeyTips are only supported in an add-in that uses a unified manifest. If your add-in uses an add-in only manifest, you must convert it to a unified manifest to configure KeyTips. For guidance on converting your manifest, see [Convert an add-in to use the unified manifest for Microsoft 365](../develop/convert-xml-to-json-manifest.md).
- Custom KeyTips support up to three uppercase alphanumeric characters and must be unique across tabs and controls.
- Custom KeyTips won't work in earlier versions of Office applications that don't support KeyTips. In these earlier versions, users will see the default KeyTips assigned to the add-in instead.
- User can't modify the add-in's KeyTips.
- In Office on the web, the <kbd>X</kbd> key is reserved and can't be used as a custom KeyTip.
- In Office on Mac, KeyTips are turned off by default. Users must turn on KeyTips in their Office settings to use the KeyTips defined for your add-in. For more information, see [Use the keyboard to work with the ribbon](https://support.microsoft.com/office/954cd3f7-2f77-4983-978d-c09b20e31f0e#picktab=mac).

## See also

- [Add custom keyboard shortcuts to your Office Add-ins](keyboard-shortcuts.md)
- [Accessibility guidelines for Office Add-ins](accessibility-guidelines.md)
