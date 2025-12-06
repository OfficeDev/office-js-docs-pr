---
title: Custom KeyTips for Office Add-ins
description: Learn how to add custom KeyTips, also known as sequential key shortcuts or access keys, to your Office Add-in.
ms.date: 12/04/2025
ms.topic: how-to
ms.localizationpriority: medium
---

# Add custom KeyTips to your Office Add-ins

KeyTips, also known as sequential key shortcuts or access keys, provide an efficient keyboard navigation method for your add-in's users. Unlike simultaneous keyboard shortcuts (such as Ctrl+S) that are used to run operations within an add-in, KeyTips are pressed in sequence to quickly access options from the Office ribbon. KeyTips help users:

- Quickly navigate the ribbon and access actions more efficiently.
- Support keyboard-only navigation for accessibility.

> [!TIP]
> The keyboard command to access KeyTips varies depending on the platform. To familiarize yourself with KeyTips, see [Use the keyboard to work with the ribbon](https://support.microsoft.com/office/954cd3f7-2f77-4983-978d-c09b20e31f0e).

## Supported Office applications and platforms

Custom KeyTips are supported in the following Office applications and platforms.

| Office app | Office on the web | Office on Windows | Office on Mac |
| ----- | ----- | ----- | ----- |
| Excel | Supported | Supported | Supported |
| Outlook | Not supported | Not supported | Not supported |
| PowerPoint | Supported | Supported | Supported |
| Word | Supported | Supported | Supported |

//TODO - Minimum supported version, if applicable

> [!NOTE]
> In Office on Mac, KeyTips are turned off by default. Users must turn on KeyTips in their Office settings to use the KeyTips defined for your add-in. For more information, see [Use the keyboard to work with the ribbon](https://support.microsoft.com/office/954cd3f7-2f77-4983-978d-c09b20e31f0e#picktab=mac).

## Supported surfaces and controls

KeyTips can be defined for the add-in's controls and the ribbon tab in which the add-in appears. The following table outlines which ribbon tabs and controls allow KeyTips for each supported Office application.

| Tab type | Excel | PowerPoint | Word |
| ----- | ----- | ----- | ----- | ----- |
| Built-in Office tabs | Supported | Supported | Supported |
| Custom tab | Supported |  Supported | Supported |
| Contextual tab | Supported | Not available | Not available |
| Buttons | Supported | Supported | Supported |
| Menu items | Supported | Supported | Supported |

## Configure the manifest

> [!IMPORTANT]
> Custom KeyTips are only supported with the unified manifest for Microsoft 365. If your add-in uses the add-in only manifest, you must migrate to the unified manifest to use KeyTips. For more information, see [Office Add-ins with the unified app manifest for Microsoft 365](../develop/unified-manifest-overview.md).

Custom KeyTips are defined in your add-in's manifest. The following example customizes KeyTips to access the add-in and its actions from the built-in Home tab of Excel, PowerPoint, or Word. Note the following about this markup.

- The `"keytip"` property defines the custom KeyTip. It's specified for the following tabs and controls.
  - The Home tab on the ribbon (`"extensions.ribbons.tabs"` object whose `"builtInTabId"` property is set to `"TabHome"`). For guidance on built-in Office ribbon tabs and their IDs, see [Find the IDs of built-in Office ribbon tabs](../develop/built-in-ui-ids.md).
  - The custom contextual tab (`"extensions.ribbons.tabs"` object whose `"id"` property is set to `"CustomTab"`).
  - The add-in's button on the ribbon (`"extensions.ribbons.tabs.groups.controls"` object).
  - The add-in's menu items on the ribbon (`"extensions.ribbons.tabs.groups.controls.items"`).
- KeyTips support up to three uppercase alphanumeric characters ("A-Z", "0-9").
- KeyTips must be unique across tabs and controls.

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
                                            ...
                                            "items": [
                                                {
                                                    "id": "InsertData1",
                                                    "type": "menuItem",
                                                    "label": "Type 1",
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
                                                        "title": "Insert data type 1",
                                                        "description": "Insert data type 1."
                                                    },
                                                    "actionId": "insertDataType1",
                                                    "keytip": "D1"
                                                },
                                                {
                                                    "id": "InsertData2",
                                                    "type": "menuItem",
                                                    "label": "Type 2",
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
                                                        "title": "Insert data type 2",
                                                        "description": "Insert data type 2."
                                                    },
                                                    "actionId": "insertDataType2",
                                                    "keytip": "D2"
                                                }
                                            ]
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

//TODO

## Localize KeyTips

You may need to localize KeyTips to support multiple languages and keyboard layouts.

To localize your custom KeyTips, see [Localize strings in your app manifest](/microsoftteams/platform/concepts/build-and-test/apps-localization#localize-strings-in-your-app-manifest).

## Behavior and limitations

When implementing KeyTips for your add-in, be aware of the following behaviors and limitations.

- Custom KeyTips are only supported in an add-in that uses a unified manifest. If your add-in uses an add-in only manifest, you must convert it to a unified manifest to configure KeyTips. For guidance on converting your manifest, see [Convert an add-in to use the unified manifest for Microsoft 365](../develop/convert-xml-to-json-manifest.md).
- Custom KeyTips support up to three uppercase alphanumeric characters.
- Custom KeyTips won't work in earlier versions of Office applications that don't support KeyTips. In these earlier versions, users will see the default KeyTips assigned to the add-in instead.
- User can't modify the add-in's KeyTips.

## See also

- [Add custom keyboard shortcuts to your Office Add-ins](keyboard-shortcuts.md)
- [Accessibility guidelines for Office Add-ins](accessibility-guidelines.md)
