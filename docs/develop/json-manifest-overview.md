---
title: Teams manifest for Office Add-ins (preview)
description: Get an overview of the preview JSON manifest.
ms.date: 05/24/2022
ms.localizationpriority: high
---

# Teams manifest for Office Add-ins (preview)

Microsoft is making a number of improvements to the Microsoft 365 developer platform. These improvements will provide more consistency in the development, deployment, installation, and administration of all types of extensions of Microsoft 365, including Office Add-ins. These changes will not break existing add-ins. 

Two of the most important features we are aiming at, *but are not yet available*:

- It should be possible to surface a single web app as multiple types of Microsoft 365 extensions. For example a web app can be both an Office Add-in and a custom tab in Teams.
- All types of Microsoft 365 extensions should use the same manifest format (JSON) and schema. It will be based on the current Teams manifest schema. In support of the first bullet, it will be possible to specify multiple types of extensions in the manifest.  

We have taken an important first step toward these goals by making it possible for Outlook Add-ins, running on Windows only, to be created with a version of the Teams JSON manifest.

The new manifest is available for preview and we encourage experienced add-in developers to experiment with it. It should not be used in production add-ins. During the early preview period, the following limitations apply:

- The preview version of the Teams manifest only supports Outlook add-ins and only on subscription Office for Windows. We're working on extending support to Excel, PowerPoint, and Word.
- It is not yet possible to combine an add-in with Teams apps, such as a custom Teams tab, or any other Microsoft 365 extension types. So, you cannot yet add an add-in to an existing Teams app. We're working on this too.

> [!TIP]
> Anxious to get started with the preview Teams manifest? Begin with [Build an Outlook add-in with a Teams manifest (preview)](../quickstarts/outlook-quickstart-json-manifest.md).

## Overview of the JSON manifest

### Schemas and general points

There are a total of 7 [Schemas](/openspecs/office_file_formats/ms-owemxml/c6a06390-34b8-4b42-82eb-b28be12494a8) to define the current XML manifest. There is just one schema for the [preview JSON manifest](/microsoftteams/platform/resources/dev-preview/developer-preview-intro.md). 

### Conceptual mapping of the preview JSON and current XML manifests

This section describes the preview JSON manifest for readers that are familiar with the current XML manifest. Some points to keep in mind: 

- JSON does not have the XML distinction between attribute and element value. So typically the JSON that maps to an XML element makes both the element value and each of the attributes a child property. For example, the following shows some XML markup and its JSON equivalent.
  
  ```xml
  <MyThing color="blue">Some text</MyThing>
  ```

  ```json
  "myThing" : {
      "color": "blue",
      "text": "Some text"
  }
  ```
- There are many places in the current XML manifest where an element with a plural name has children with the singular version of the same name. For example, the markup to configure a custom menu includes an **Items** element which can have multiple **Item** element children. The JSON equivalent of these plural elements is a property with an array as its value. The members of the array are *anonymous* objects, not properties named "item" or "item1", "item2", etc. The following is an example.

  ```json
  "items": [
      {
          -- markup for a menu item is here --
      },
      {
          -- markup for another menu item is here --
      }
  ]
  ```

#### Top level structure

The root level of the preview JSON manifest, which roughly corresponds to the **OfficeApp** element in the current XML manifest, is an anonymous object. 

The children of **OfficeApp** are commonly divided into two notional categories. The **VersionOverrides** element is one category. The other consists of all the other children of **OfficeApp** which are collectively referred to as the base manifest. So too, the preview JSON manifest has a similar division. There is a top-level "extension" property that roughly corresponds in its purposes and child properties to the **VersionOverrides** element. The preview JSON manifest also has over 10 other top-level properties that collectively serve the same purposes as the base manifest of the XML manifest. So, these other properties can be thought of collectively as the base manifest of the JSON manifest. 

> [!NOTE]
> When it becomes possible to combine an add-in with other Microsoft 365 extension types in a single web app, then there will be other top-level properties that don't fit into the notion of the base manifest. Roughly speaking, there will be a top-level property for every kind of Microsoft 365 extension type, such as "configurableTabs", "bots" and "connectors". For examples, see the [Teams manifest documentation](/microsoftteams/platform/resources/schema/manifest-schema). This structure makes clear that the "extension" property represents an Office add-in as one type of Microsoft 365 extension.

#### Base manifest

The base manifest properties specify characteristics of the add-in that *any* type of extension of Microsoft 365, including Teams tabs or message extensions, would be expected to have, not just Office add-ins. These characteristics include a public name and a unique ID. The following table shows a mapping of some critical top level properties in the preview JSON manifest to the XML elements in the current manifest, where the mapping principle is the *purpose* of the markup.

|JSON property|Purpose|XML element(s)|Comments|
|:-----|:-----|:-----|:-----|
|"$schema"| Identifies the manifest schema. | attributes of **OfficeApp** and **VersionOverrides** | |
|"id"| GUID of the add-in. | **Id**| |
|"version"| Version of the add-in. | **Version** | |
|"manifestVersion"| Version of the manifest schema. |  attributes of **OfficeApp** | |
|"name"| Public name of the add-in. | **DisplayName** | |
|"description"| Public description of the add-in.  | **Description** | |
|"accentColor"||| This property has no equivalent in the current XML manifest and is not used in the preview of the JSON manifest. But it must be present. |
|"developer"| Identifies the developer of the add-in. | **ProviderName** | |
|"localizationInfo"| Configures the default locale and other supported locales. | **DefaultLocale** and **Override** | |
|"webApplicationInfo"| Identifies the add-in's web app as it is known in Azure Active Directory. | **WebApplicationInfo** | In the current XML manifest, the **WebApplicationInfo** element is inside **VersionOverrides**, not the base manifest. |
|"authorization"| Identifies any Microsoft Graph permissions that the add-in needs. | **WebApplicationInfo** | See comment in preceding row. |

The **Hosts**, **Requirements**, and **ExtendedOverrides** elements are part of the base manifest in the current XML manifest. But concepts and purposes associated with these elements are configured inside the "extension" property of the preview JSON manifest. 

#### "extension" property

The "extension" property in the preview JSON manifest primarily represents characteristics of the add-in that would not be relevant to other kinds of Microsoft 365 extensions. For example, the Office applications that the add-in extends (such as, Excel, PowerPoint, Word, and Outlook) are specified inside the "extension" property as are customizations of the Office application ribbon. The property's purposes closely match those of the **VersionOverrides** element in the current XML manifest.

> [!NOTE]
> The **VersionOverrides** section of the manifest has a "double jump" system for many string resources. Strings, including URLs, are specified and assigned an ID in the **Resources** child of **VersionOverrides**. Elements that require a string have a `resid` attribute that matches the ID of a string in the **Resources** element. The "extension" property of the preview JSON manifest simplifies things by defining strings directly as property values. There is nothing equivalent to the **Resources** element.

The following table shows a mapping of some high level child properties of the "extension" property in the preview JSON manifest to XML elements in the current manifest. Dot notation is used to reference child properties.

|JSON property|Purpose|XML element(s)|Comments|
|:-----|:-----|:-----|:-----|
| "requirements.capabilities" | Identifies the requirement sets that the add-in needs to be installable. | **Requirements** and **Sets** | |
| "requirements.scopes" | Identifies the Office applications in which the add-in can be installed. | **Hosts** |  |
| "ribbons" | The ribbons that the add-in customizes. | **Hosts**, **ExtensionPoints**, and various **\*FormFactor** elements | The "ribbons" property is an array of anonymous objects that each merge the purposes of the these three elements. See ["ribbons" table](#ribbons-table).|
| "alternatives" | Specifies backwards compatibility with an equivalent COM add-in, XLL, or both. | **EquivalentAddins** | See the [EquivalentAddins - See also](/javascript/api/manifest/equivalentaddins#see-also) for background information. |
| "runtimes"  | Configures various kinds of "UI-less" add-ins such as custom functions and functions run directly from custom ribbon buttons. | **Runtimes**. **FunctionFile**, and **ExtensionPoint** (of type CustomFunctions) |  |
| "autoRunEvents" | Configures an event handler for a specified event. | **Event** and **ExtensionPoint** (of type Events) |  |

##### "ribbons" table

The following table maps the child properties of the anonymous child objects in the "ribbons" array onto XML elements in the current manifest. 

|JSON property|Purpose|XML element(s)|Comments|
|:-----|:-----|:-----|:-----|
| "contexts" | Specifies the command surfaces that the add-in customizes. | various **\*CommandSurface** elements, such as **PrimaryCommandSurface** and **MessageReadCommandSurface** |  |
| "tabs" | Configures custom ribbon tabs. | **CustomTab** | The names and hierarchy of the descendant properties of "tabs" closely match the descendants of **CustomTab**.  |

## Sample preview JSON manifest

The following is an example of a preview JSON-manifest for an add-in.

```json
{
    "$schema": "https://raw.githubusercontent.com/OfficeDev/microsoft-teams-app-schema/op/extensions/MicrosoftTeams.schema.json",
    "id": "82a9d9c3-4702-4322-bbc4-6fe7f9b01483",
    "version": "1.0.0",
    "manifestVersion": "devPreview",
    "name": {
        "short": "Basic Office Example",
        "full": "Transform text to uppercase/lowercase."
    },
    "description": {
        "short": "Example MetaOS app that demonstrates various Office features.",
        "full": "Example MetaOS app that demonstrates various Office features like ribbon, menubar, context menu, keyboard shortcuts, custom functions."
    },
    "icons": {
        "outline": "small_icon.png",
        "color": "color_icon.png"
    },
    "accentColor": "#230201",
    "developer": {
        "name": "Microsoft Corp.",
        "websiteUrl": "https://aka.ms/opc_metaos_examples",
        "privacyUrl": "https://aka.ms/opc_metaos_privacy",
        "termsOfUseUrl": "https://aka.ms/opc_metaos_examples"
    },
    "localizationInfo": {
        "defaultLanguageTag": "en-us",
        "additionalLanguages": [
            {
                "languageTag": "es-es",
                "file": "locales/es-es.json"
            }
        ]
    },
    "webApplicationInfo": {
        "id": "c62f9f19-d901-48c8-a184-9a69d83305bc",
        "resource": "api://www.oepbuild2017.com/prodapp"
    },
    "authorization": {
        "permissions": {
            "resourceSpecific": [
                {
                    "name": "Mailbox.ReadWrite",
                    "type": "Delegated"
                },
                {
                    "name": "Document.ReadWrite",
                    "type": "Delegated"
                }
            ]
        }

    },
    "extension": {
        "requirements": {
            "scopes": [
                "document",
                "presentation",
                "workbook",
                "mail"
            ],
            "capabilities": [
               { "name": "AddinCommands", "minVersion": "1.1" }
            ]
        },
        "getStartedMessages": [
            {
                "requirements": {
                    "formFactors": ["desktop"]
                },
                "title": "Get Started",
                "description": "Your sample add-in loaded succesfully. Click the buttons to get started.",
                "learnMoreUrl": "https://aka.ms/opc_metaos_examples"
            }
        ],
        "runtimes": [
            {
                "requirements": {
                        "scopes": [
                            "workbook"
                        ],
                        "capabilities": [
                            { "name": "CustomFunctions", "minVersion": "1.1" }
                    ]
                },
                "id": "text",
                "type": "general",
                "code": {
                    "page": "https://aka.ms/opc_metaos_examples/alpha/elements/en-us/actions_text.html",
                    "script": "https://aka.ms/opc_metaos_examples/alpha/elements/en-us/actions_text.js"
                },
                "lifetime": "short",
                "actions": {
                    "items": [
                        {
                            "id": "text.toUppercase",
                            "type": "execution",
                            "name": "ToUppercase"
                        },
                        {
                            "id": "text.toLowercase",
                            "type": "execution",
                            "name": "ToLowercase"
                        },
                        {
                            "id": "text.showDashboard",
                            "type": "contextual-launch",
                            "view": "dashboard",
                            "name": "Dashboard"
                        }
                    ]
                },
                "functions": {
                    "namespace": {
                        "id": "Microsoft.Alpha",
                        "name": "Microsoft.Alpha"
                    },
                    "items": [
                        {
                            "id": "text.toUppercase",
                            "name": "ToUppercase",
                            "description": "Returns the input text as uppercase.",
                            "parameters": [
                                {
                                    "name": "InputText",
                                    "description": "Input text.",
                                    "type": "string"
                                }
                            ],
                            "result": {
                                "dimensionality": "scalar"
                            },
                            "stream": false,
                            "volatile": true,
                            "cancelable": false,
                            "requiresAddress": true,
                            "requiresParameterAddress": false
                        },
                        {
                            "id": "text.toLowercase",
                            "name": "ToLowercase",
                            "description": "Returns the input text as lowercase.",
                            "parameters": [
                                {
                                    "name": "InputText",
                                    "description": "Input text.",
                                    "type": "string"
                                }
                            ],
                            "result": {
                                "dimensionality": "scalar"
                            }
                        }
                    ],
                    "allowErrorForDataTypeAny": false,
                    "allowRichDataForDataTypeAny": true
                }
            },
    
            {
                "requirements": {
                    "capabilities": [
                        { "name": "MailBox", "minVersion": "1.10" }
                    ]
                },
                "id": "text",
                "type": "general",
                "code": {
                    "page": "https://aka.ms/opc_metaos_examples/alpha/elements/en-us/actions_text.html",
                    "script": "https://aka.ms/opc_metaos_examples/alpha/elements/en-us/actions_text.js"
                },
                "lifetime": "short",
                "actions": {
                    "items": [
                        {
                            "id": "text.onMessageSending",
                            "type": "execution",
                            "name": "OnMessageSending"
                        },
                        {
                            "id": "text.onNewMessageComposeCreated",
                            "type": "execution",
                            "name": "OnNewMessageComposeCreated"
                        }
                    ]
                }
            },
            {
                "requirements": {
                    "capabilities": [
                        { "name": "CustomFunctions", "maxVersion": "0.0" }
                    ]
                },
                "id": "text",
                "type": "general",
                "code": {
                    "page": "https://aka.ms/opc_metaos_examples/alpha/elements/en-us/actions_text.html",
                    "script": "https://aka.ms/opc_metaos_examples/alpha/elements/en-us/actions_text.js"
                },
                "lifetime": "short",
                "actions": {
                    "items": [
                        {
                            "id": "text.toLowercase",
                            "type": "execution",
                            "name": "ToLowercase"
                        },
                        {
                            "id": "text.toUppercase",
                            "type": "execution",
                            "name": "ToUppercase"
                        },
                        {
                            "id": "text.showDashboard",
                            "type": "contextual-launch",
                            "view": "dashboard"
                        }
                    ]
                }
            }
        ],
        "contextMenus": [
            {
                "menus": [
                    {
                        "type": "cell",
                        "controls": [
                            {
                                "id": "menuForCell",
                                "type": "menu",
                                "label": "Menu",
                                "icons": [
                                    { "size": 16, "file": "test_16.png" },
                                    { "size": 32, "file": "test_32.png" },
                                    { "size": 80, "file": "test_80.png" }
                                ],
                                "supertip": {
                                    "title": "Change text case",
                                    "description": "This allow you to change text to lowercase or uppercase."
                                },
                                "items": [
                                    {
                                        "id": "menu.uppercase",
                                        "type": "menuItem",
                                        "label": "To uppercase",
                                        "supertip": {
                                            "title": "Change text to uppercase",
                                            "description": "This will change the text to uppercase."
                                        },
                                        "actionId": "text.toUppercase"
                                    },
                                    {
                                        "id": "menu.lowercase",
                                        "type": "menuItem",
                                        "label": "To lowercase",
                                        "supertip": {
                                            "title": "Change text to lowercase",
                                            "description": "This will change the text to lowercase."
                                        },
                                        "actionId": "text.toLowercase"
                                    }
                                ]
                            },
                            {
                                "id": "showDashboard",
                                "type": "button",
                                "label": "Show dashboard",
                                "icons": [
                                    { "size": 16, "file": "test_16.png" },
                                    { "size": 32, "file": "test_32.png" },
                                    { "size": 80, "file": "test_80.png" }
                                ],
                                "supertip": {
                                    "title": "Show dashboard",
                                    "description": "click to open dashboard"
                                },
                                "actionId": "text.showDashboard"
                            }
                        ]
                    },
                    {
                        "type": "text",
                        "controls": [
                            {
                                "id": "menuForText",
                                "type": "menu",
                                "label": "Menu",
                                "icons": [
                                    { "size": 16, "file": "test_16.png" },
                                    { "size": 32, "file": "test_32.png" },
                                    { "size": 80, "file": "test_80.png" }
                                ],
                                "supertip": {
                                    "title": "Change text case",
                                    "description": "This allow you to change text to lowercase or uppercase."
                                },
                                "items": [
                                    {
                                        "id": "menu.uppercase",
                                        "type": "menuItem",
                                        "label": "To uppercase",
                                        "supertip": {
                                            "title": "Change text to uppercase",
                                            "description": "This will change the text to uppercase."
                                        },
                                        "actionId": "text.toUppercase"
                                    },
                                    {
                                        "id": "menu.lowercase",
                                        "type": "menuItem",
                                        "label": "To lowercase",
                                        "supertip": {
                                            "title": "Change text to lowercase",
                                            "description": "This will change the text to lowercase."
                                        },
                                        "actionId": "text.toLowercase"
                                    }
                                ]
                            },
                            {
                                "id": "showDashboard",
                                "type": "button",
                                "label": "Show dashboard",
                                "icons": [
                                    { "size": 16, "file": "test_16.png" },
                                    { "size": 32, "file": "test_32.png" },
                                    { "size": 80, "file": "test_80.png" }
                                ],
                                "supertip": {
                                    "title": "Show dashboard",
                                    "description": "click to open dashboard"
                                },
                                "actionId": "text.showDashboard"
                            }
                        ]
                    }
                ]
            }
        ],
        "ribbons": [
            {
                "requirements": {
                    "scopes": ["presentation"]
                },
                "tabs": [
                    {
                        "id": "dashboard",
                        "label": "Dashboard",
                        "position": {
                            "nativeId": "tabHome",
                            "align": "after"
                        },
                        "groups": [
                            {
                                "id": "dashboard",
                                "label": "Controls",
                                "icons": [
                                    { "size": 16, "file": "test_16.png" },
                                    { "size": 32, "file": "test_32.png" },
                                    { "size": 80, "file": "tst_80.png" }
                                ],
                                "controls": [
                                    {
                                        "id": "showDashboard",
                                        "type": "button",
                                        "label": "Show dashboard",
                                        "icons": [
                                            { "size": 16, "file": "test_16.png" },
                                            { "size": 32, "file": "test_32.png" },
                                            { "size": 80, "file": "test_80.png" }
                                        ],
                                        "supertip": {
                                            "title": "Show dashboard",
                                            "description": "click to open dashboard"
                                        },
                                        "actionId": "text.showDashboard"
                                    },
                                    {
                                        "nativeId": "undo"
                                    }
                                ]
                            },
                            {
                                "nativeId": "font"
                            }
                        ]
                    }
                ]
            },
            {
                "contexts": [
                    "default"
                ],
                "tabs": [
                    {
                        "nativeId": "tabDefault",
                        "groups": [
                            {
                                "id": "dashboard",
                                "label": "Controls",
                                "icons": [
                                    { "size": 16, "file": "test_16.png" },
                                    { "size": 32, "file": "test_32.png" },
                                    { "size": 80, "file": "test_80.png" }
                                ],
                                "controls": [
                                    {
                                        "id": "uppercase",
                                        "type": "button",
                                        "label": "To uppercase",
                                        "icons": [
                                            { "size": 16, "file": "test_16.png" },
                                            { "size": 32, "file": "test_32.png" },
                                            { "size": 80, "file": "test_80.png" }
                                        ],
                                        "supertip": {
                                            "title": "Change text to uppercase",
                                            "description": "This will change the text to uppercase."
                                        },
                                        "actionId": "text.toUppercase"
                                    },
                                    {
                                        "id": "lowercase",
                                        "type": "button",
                                        "label": "To lowercase",
                                        "enabled": false,
                                        "icons": [
                                            { "size": 16, "file": "test_16.png" },
                                            { "size": 32, "file": "test_32.png" },
                                            { "size": 80, "file": "test_80.png" }
                                        ],
                                        "supertip": {
                                            "title": "Change text to lowercase",
                                            "description": "This will change the text to lowercase."
                                        },
                                        "actionId": "text.toLowercase"
                                    },
                                    {
                                        "id": "menu",
                                        "type": "menu",
                                        "label": "Menu",
                                        "icons": [
                                            { "size": 16, "file": "test_16.png" },
                                            { "size": 32, "file": "test_32.png" },
                                            { "size": 80, "file": "test_80.png" }
                                        ],
                                        "supertip": {
                                            "title": "Change text case",
                                            "description": "This allow you to change text to lowercase or uppercase."
                                        },
                                        "items": [
                                            {
                                                "id": "menu.uppercase",
                                                "type": "menuItem",
                                                "label": "To uppercase",
                                                "enabled": false,
                                                "icons": [
                                                    { "size": 16, "file": "test_16.png" },
                                                    { "size": 32, "file": "test_32.png" },
                                                    { "size": 80, "file": "test_80.png" }
                                                ],
                                                "supertip": {
                                                    "title": "Change text to uppercase",
                                                    "description": "This will change the text to uppercase."
                                                },
                                                "actionId": "text.toUppercase"
                                            },
                                            {
                                                "id": "menu.lowercase",
                                                "type": "menuItem",
                                                "label": "To lowercase",
                                                "supertip": {
                                                    "title": "Change text to lowercase",
                                                    "description": "This will change the text to lowercase."
                                                },
                                                "actionId": "text.toLowercase",
                                                "overriddenByRibbonApi": true
                                            }
                                        ]
                                    },
                                    {
                                        "id": "showDashboard",
                                        "type": "button",
                                        "label": "Show dashboard",
                                        "icons": [
                                            { "size": 16, "file": "test_16.png" },
                                            { "size": 32, "file": "test_32.png" },
                                            { "size": 80, "file": "test_80.png" }
                                        ],
                                        "supertip": {
                                            "title": "Show dashboard",
                                            "description": "click to open dashboard"
                                        },
                                        "actionId": "text.showDashboard",
                                        "overriddenByRibbonApi": true
                                    }
                                ]
                            }
                        ]
                    }
                ]
            },
            {
                "contexts": [
                    "composeMail"
                ],
                "tabs": [
                    {
                        "nativeId": "tabDefault",
                        "groups": [
                            {
                                "id": "dashboard",
                                "label": "Controls",
                                "controls": [
                                    {
                                        "id": "uppercase",
                                        "type": "button",
                                        "label": "To uppercase",
                                        "icons": [
                                            { "size": 16, "file": "test_16.png" },
                                            { "size": 32, "file": "test_32.png" },
                                            { "size": 80, "file": "test_80.png" }
                                        ],
                                        "supertip": {
                                            "title": "Change text to uppercase",
                                            "description": "This will change the text to uppercase."
                                        },
                                        "actionId": "text.toUppercase"
                                    },
                                    {
                                        "id": "lowercase",
                                        "type": "button",
                                        "label": "To lowercase",
                                        "icons": [
                                            { "size": 16, "file": "test_16.png" },
                                            { "size": 32, "file": "test_32.png" },
                                            { "size": 80, "file": "test_80.png" }
                                        ],
                                        "supertip": {
                                            "title": "Change text to lowercase",
                                            "description": "This will change the text to lowercase."
                                        },
                                        "actionId": "text.toLowercase"
                                    },
                                    {
                                        "id": "menu",
                                        "type": "menu",
                                        "label": "Menu",
                                        "icons": [
                                            { "size": 16, "file": "test_16.png" },
                                            { "size": 32, "file": "test_32.png" },
                                            { "size": 80, "file": "test_80.png" }
                                        ],
                                        "supertip": {
                                            "title": "Change text case",
                                            "description": "This allow you to change text to lowercase or uppercase."
                                        },
                                        "items": [
                                            {
                                                "id": "menu.uppercase",
                                                "type": "menuItem",
                                                "label": "To uppercase",
                                                "supertip": {
                                                    "title": "Change text to uppercase",
                                                    "description": "This will change the text to uppercase."
                                                },
                                                "actionId": "text.toUppercase"
                                            },
                                            {
                                                "id": "menu.lowercase",
                                                "type": "menuItem",
                                                "label": "To lowercase",
                                                "icons": [
                                                    { "size": 16, "file": "test_16.png" },
                                                    { "size": 32, "file": "test_32.png" },
                                                    { "size": 80, "file": "test_80.png" }
                                                ],
                                                "supertip": {
                                                    "title": "Change text to lowercase",
                                                    "description": "This will change the text to lowercase."
                                                },
                                                "actionId": "text.toLowercase"
                                            }
                                        ]
                                    },
                                    {
                                        "id": "showDashboard",
                                        "type": "button",
                                        "label": "Show dashboard",
                                        "icons": [
                                            { "size": 16, "file": "test_16.png" },
                                            { "size": 32, "file": "test_32.png" },
                                            { "size": 80, "file": "test_80.png" }
                                        ],
                                        "supertip": {
                                            "title": "Show dashboard",
                                            "description": "click to open dashboard"
                                        },
                                        "actionId": "text.showDashboard"
                                    }
                                ]
                            }
                        ]
                    }
                ]
            },
            {
                "contexts": [
                    "readMail"
                ],
                "tabs": [
                    {
                        "nativeId": "tabDefault",
                        "groups": [
                            {
                                "id": "dashboard",
                                "label": "Controls",
                                "controls": [
                                    {
                                        "id": "showDashboard",
                                        "type": "button",
                                        "label": "Show dashboard",
                                        "icons": [
                                            { "size": 16, "file": "test_16.png" },
                                            { "size": 32, "file": "test_32.png" },
                                            { "size": 80, "file": "test_80.png" }
                                        ],
                                        "supertip": {
                                            "title": "Show dashboard",
                                            "description": "click to open dashboard"
                                        },
                                        "actionId": "text.showDashboard"
                                    }
                                ]
                            }
                        ]
                    }
                ]
            }
        ],
        "keyboards": [
            {
                "requirements": {
                    "capabilities": [
                        { "name": "SharedRuntime", "minVersion": "1.1" }
                    ],
                    "platforms": ["windows", "web"]
                },
                "shortcuts": [
                    {
                        "key":  "Ctrl+Shift+U",
                        "actionId": "text.toUppercase"
                    },
                    {
                        "key": "Ctrl+Shift+L",
                        "actionId": "text.toLowercase"
                    },
                    {
                        "key": "Ctrl+Shift+Up",
                        "actionId": "text.showDashboard"
                    }
                ]
            },
            {
                "requirements": {
                    "capabilities": [
                        { "name": "SharedRuntime", "minVersion": "1.1" }
                    ],
                    "platforms": ["mac"]
                },
                "shortcuts": [
                    {
                        "key": "Command+Shift+U",
                        "actionId": "text.toUppercase"
                    },
                    {
                        "key": "Command+Shift+L",
                        "actionId": "text.toLowercase"
                    },
                    {
                        "key": "Command+Shift+Up",
                        "actionId": "text.showDashboard"
                    }
                ]
            }
        ],
        "autoRunEvents": [
            {
                "requirements": {
                    "capabilities": [
                        { "name": "MailBox", "minVersion": "1.10" }
                    ]
                },
                "events": [
                    {
                        "id": "newMessageComposeCreated",
                        "actionId": "text.onNewMessageComposeCreated"
                    },
                    {
                        "id": "messageSending",
                        "actionId": "text.onMessageSending",
                        "options": {
                            "sendMode": "promptUser"
                        }
                    }
                ]
            }
        ],
        "alternatives": [
            {
                "requirements": {
                    "scopes": ["mail"]
                },
                "prefer": {
                    "comAddin": {
                        "progId": "ContosoExtension"
                    }
                },
                "hide": {
                    "storeOfficeAddin": {
                        "officeAddinId": "fca2794d-4aa5-4023-a84b-c60a3cbd33d4",
                        "assetId": "WA129485"
                    }
                }
            },
            {
                "requirements": {
                    "scopes": ["presentation", "document", "workbook"]
                },
                "prefer": {
                    "xllCustomfunctions": {
                        "fileName": "ContosoExtension.xll"
                    },
                    "comAddin": {
                        "progId": "ContosoExtension"
                    }
                },
                "hide": {
                    "customOfficeAddin": {
                        "officeAddinId": "b5a2794d-4aa5-4023-a84b-c60a3cbd33d4"
                    }
                }
            }
        ]
    }
}
```

## Next steps

- [Build an Outlook add-in with a Teams manifest (preview)](../quickstarts/outlook-quickstart-json-manifest.md).