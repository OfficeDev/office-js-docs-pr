---
title:  Unified Microsoft 365 manifest (preview)
description: Get an overview of the preview JSON-formatted unified Microsoft 365 manifest.
ms.topic: overview
ms.date: 06/15/2022
ms.localizationpriority: high
---

# Unified Microsoft 365 manifest (preview)

Microsoft is making a number of improvements to the Microsoft 365 developer platform. These improvements provide more consistency in the development, deployment, installation, and administration of all types of extensions of Microsoft 365, including Office Add-ins. These changes are compatible with existing add-ins.

One important improvement we are working on is the ability to create a single unit of distribution for all your Microsoft 365 extensions by using the same manifest format and schema, based on the current JSON-formatted unified Microsoft 365 manifest.

We've taken an important first step toward these goals by making it possible for you to create Outlook add-ins, running on Windows only, with a unified Microsoft 365 manifest.

> [!NOTE]
> The new manifest is available for preview and is subject to change based on feedback. We encourage experienced add-in developers to experiment with it. The preview manifest should not be used in production add-ins.

During the early preview period, the following limitations apply.

- The preview version of the unified manifest only supports Outlook add-ins and only in Office downloaded from a Microsoft 365 subscription then installed on Windows. We're working on extending support to Excel, PowerPoint, and Word.
- It isn't yet possible to combine, and sideload, an add-in with a Teams app, such as a Teams personal tab, or other Microsoft 365 extension types. In the coming months, we will continue to extend the preview to support these scenarios and provide additional tools to update manifests to the preview format.

> [!TIP]
> Ready to get started with the preview unified manifest? Begin with [Build an Outlook add-in with the unified Microsoft 365 manifest (preview)](../quickstarts/outlook-quickstart-json-manifest.md).

## Overview of the unified manifest

### Schemas and general points

There is just one schema for the [preview unified manifest](/microsoftteams/platform/resources/dev-preview/developer-preview-intro), in contrast to the current XML manifest which has a total of seven [Schemas](/openspecs/office_file_formats/ms-owemxml/c6a06390-34b8-4b42-82eb-b28be12494a8).  

### Conceptual mapping of the preview unified and current XML manifests

This section describes the preview unified manifest for readers who are familiar with the current XML manifest. Some points to keep in mind: 

- The unified manifest is JSON-formatted.
- JSON does not distinguish between attribute and element value like XML does. Typically, the JSON that maps to an XML element makes both the element value and each of the attributes a child property. The following example shows some XML markup and its JSON equivalent.
  
  ```xml
  <MyThing color="blue">Some text</MyThing>
  ```

  ```json
  "myThing" : {
      "color": "blue",
      "text": "Some text"
  }
  ```
- There are many places in the current XML manifest where an element with a plural name has children with the singular version of the same name. For example, the markup to configure a custom menu includes an **\<Items\>** element which can have multiple **\<Item\>** element children. The JSON equivalent of these plural elements is a property with an array as its value. The members of the array are *anonymous* objects, not properties named "item" or "item1", "item2", etc. The following is an example.

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

#### Top-level structure

The root level of the preview unified manifest, which roughly corresponds to the **\<OfficeApp\>** element in the current XML manifest, is an anonymous object. 

The children of **\<OfficeApp\>** are commonly divided into two notional categories. The **\<VersionOverrides\>** element is one category. The other consists of all the other children of **\<OfficeApp\>**, which are collectively referred to as the base manifest. So too, the preview unified manifest has a similar division. There is a top-level "extension" property that roughly corresponds in its purposes and child properties to the **\<VersionOverrides\>** element. The preview unified manifest also has over 10 other top-level properties that collectively serve the same purposes as the base manifest of the XML manifest. These other properties can be thought of collectively as the base manifest of the unified manifest. 

> [!NOTE]
> When it becomes possible to combine an add-in with other Microsoft 365 extension types in a single manifest, then there will be other top-level properties that don't fit into the notion of the base manifest. There will usually be a top-level property for every kind of Microsoft 365 extension type, such as "configurableTabs", "bots" and "connectors". For examples, see the [unified manifest documentation](/microsoftteams/platform/resources/schema/manifest-schema). This structure makes clear that the "extension" property represents an Office add-in as one type of Microsoft 365 extension.

#### Base manifest

The base manifest properties specify characteristics of the add-in that *any* type of extension of Microsoft 365 is expected to have. This includes Teams tabs and message extensions, not just Office add-ins. These characteristics include a public name and a unique ID. The following table shows a mapping of some critical top-level properties in the preview unified manifest to the XML elements in the current manifest, where the mapping principle is the *purpose* of the markup.

|JSON property|Purpose|XML elements|Comments|
|:-----|:-----|:-----|:-----|
|"$schema"| Identifies the manifest schema. | attributes of **\<OfficeApp\>** and **\<VersionOverrides\>** |*None.* |
|"id"| GUID of the add-in. | **\<Id\>**|*None.* |
|"version"| Version of the add-in. | **\<Version\>** |*None.* |
|"manifestVersion"| Version of the manifest schema. |  attributes of **\<OfficeApp\>** |*None.* |
|"name"| Public name of the add-in. | **\<DisplayName\>** |*None.* |
|"description"| Public description of the add-in.  | **\<Description\>** |*None.* |
|"accentColor"|*None.* |*None.* | This property has no equivalent in the current XML manifest and isn't used in the preview of the unified manifest. But it must be present. |
|"developer"| Identifies the developer of the add-in. | **\<ProviderName\>** |*None.* |
|"localizationInfo"| Configures the default locale and other supported locales. | **\<DefaultLocale\>** and **\<Override\>** |*None.* |
|"webApplicationInfo"| Identifies the add-in's web app as it is known in Azure Active Directory. | **\<WebApplicationInfo\>** | In the current XML manifest, the **\<WebApplicationInfo\>** element is inside **\<VersionOverrides\>**, not the base manifest. |
|"authorization"| Identifies any Microsoft Graph permissions that the add-in needs. | **\<WebApplicationInfo\>** | In the current XML manifest, the **\<WebApplicationInfo\>** element is inside **\<VersionOverrides\>**, not the base manifest. |

The **\<Hosts\>**, **\<Requirements\>**, and **\<ExtendedOverrides\>** elements are part of the base manifest in the current XML manifest. But concepts and purposes associated with these elements are configured inside the "extension" property of the preview unified manifest.

#### "extension" property

The "extension" property in the preview unified manifest primarily represents characteristics of the add-in that would not be relevant to other kinds of Microsoft 365 extensions. For example, the Office applications that the add-in extends (such as, Excel, PowerPoint, Word, and Outlook) are specified inside the "extension" property, as are customizations of the Office application ribbon. The configuration purposes of the "extension" property closely match those of the **\<VersionOverrides\>** element in the current XML manifest.

> [!NOTE]
> The **\<VersionOverrides\>** section of the current XML manifest has a "double jump" system for many string resources. Strings, including URLs, are specified and assigned an ID in the **\<Resources\>** child of **\<VersionOverrides\>**. Elements that require a string have a `resid` attribute that matches the ID of a string in the **\<Resources\>** element. The "extension" property of the preview unified manifest simplifies things by defining strings directly as property values. There is nothing in the unified manifest that is equivalent to the **\<Resources\>** element.

The following table shows a mapping of some high level child properties of the "extension" property in the preview unified manifest to XML elements in the current manifest. Dot notation is used to reference child properties.

|JSON property|Purpose|XML elements|Comments|
|:-----|:-----|:-----|:-----|
| "requirements.capabilities" | Identifies the requirement sets that the add-in needs to be installable. | **\<Requirements\>** and **\<Sets\>** |*None.* |
| "requirements.scopes" | Identifies the Office applications in which the add-in can be installed. | **\<Hosts\>** |*None.* |
| "ribbons" | The ribbons that the add-in customizes. | **\<Hosts\>**, **ExtensionPoints**, and various **\*FormFactor** elements | The "ribbons" property is an array of anonymous objects that each merge the purposes of the these three elements. See ["ribbons" table](#ribbons-table).|
| "alternatives" | Specifies backwards compatibility with an equivalent COM add-in, XLL, or both. | **\<EquivalentAddins\>** | See the [EquivalentAddins - See also](/javascript/api/manifest/equivalentaddins#see-also) for background information. |
| "runtimes"  | Configures various kinds of add-ins that have little or no UI, such as custom function-only add-ins and [function commands](../design/add-in-commands.md#types-of-add-in-commands). | **\<Runtimes\>**. **\<FunctionFile\>**, and **\<ExtensionPoint\>** (of type CustomFunctions) |*None.* |
| "autoRunEvents" | Configures an event handler for a specified event. | **\<Event\>** and **\<ExtensionPoint\>** (of type Events) |*None.* |

##### "ribbons" table

The following table maps the child properties of the anonymous child objects in the "ribbons" array onto XML elements in the current manifest. 

|JSON property|Purpose|XML elements|Comments|
|:-----|:-----|:-----|:-----|
| "contexts" | Specifies the command surfaces that the add-in customizes. | various **\*CommandSurface** elements, such as **PrimaryCommandSurface** and **MessageReadCommandSurface** |*None.* |
| "tabs" | Configures custom ribbon tabs. | **\<CustomTab\>** | The names and hierarchy of the descendant properties of "tabs" closely match the descendants of **\<CustomTab\>**.  |

## Sample preview unified manifest

The following is an example of a preview unified manifest for an add-in.

```json
{
  "$schema": "https://raw.githubusercontent.com/OfficeDev/microsoft-teams-app-schema/op/extensions/MicrosoftTeams.schema.json",
  "id": "00000000-0000-0000-0000-000000000000",
  "version": "1.0.0",
  "manifestVersion": "devPreview",
  "name": {
    "short": "Name of your app (<=30 chars)",
    "full": "Full name of app, if longer than 30 characters (<=100 chars)"
  },
  "description": {
    "short": "Short description of your app (<= 80 chars)",
    "full": "Full description of your app (<= 4000 chars)"
  },
  "icons": {
    "outline": "outline.png",
    "color": "color.png"
  },
  "accentColor": "#230201",
  "developer": {
    "name": "Contoso",
    "websiteUrl": "https://www.contoso.com",
    "privacyUrl": "https://www.contoso.com/privacy",
    "termsOfUseUrl": "https://www.contoso.com/servicesagreement"
  },
  "localizationInfo": {
    "defaultLanguageTag": "en-us",
    "additionalLanguages": [
      {
        "languageTag": "es-es",
        "file": "es-es.json"
      }
    ]
  },
  "webApplicationInfo": {
    "id": "00000000-0000-0000-0000-000000000000",
    "resource": "api://www.contoso.com/prodapp"
  },
  "authorization": {
    "permissions": {
      "resourceSpecific": [
        {
          "name": "Mailbox.ReadWrite.User",
          "type": "Delegated"
        }
      ]
    }
  },
  "extensions": [
    {
      "requirements": {
        "scopes": [ "mail" ],
        "capabilities": [
          {
            "name": "Mailbox", "minVersion": "1.1"
          }
        ]
      },
      "runtimes": [
        {
          "requirements": {
            "capabilities": [
              {
                "name": "MailBox", "minVersion": "1.10"
              }
            ]
          },
          "id": "eventsRuntime",
          "type": "general",
          "code": {
            "page": "https://contoso.com/events.html",
            "script": "https://contoso.com/events.js"
          },
          "lifetime": "short",
          "actions": [
            {
              "id": "onMessageSending",
              "type": "executeFunction"
            },
            {
              "id": "onNewMessageComposeCreated",
              "type": "executeFunction"
            }
          ]
        },
        {
          "requirements": {
            "capabilities": [
              {
                "name": "MailBox", "minVersion": "1.1"
              }
            ]
          },
          "id": "commandsRuntime",
          "type": "general",
          "code": {
            "page": "https://contoso.com/commands.html",
            "script": "https://contoso.com/commands.js"
          },
          "lifetime": "short",
          "actions": [
            {
              "id": "action1",
              "type": "executeFunction"
            },
            {
              "id": "action2",
              "type": "executeFunction"
            },
            {
              "id": "action3",
              "type": "executeFunction"
            }
          ]
        }
      ],
      "ribbons": [
        {
          "contexts": [
            "mailCompose"
          ],
          "tabs": [
            {
              "builtInTabId": "TabDefault",
              "groups": [
                {
                  "id": "dashboard",
                  "label": "Controls",
                  "controls": [
                    {
                      "id": "control1",
                      "type": "button",
                      "label": "Action 1",
                      "icons": [
                        {
                          "size": 16,
                          "file": "test_16.png"
                        },
                        {
                          "size": 32,
                          "file": "test_32.png"
                        },
                        {
                          "size": 80,
                          "file": "test_80.png"
                        }
                      ],
                      "supertip": {
                        "title": "Action 1 Title",
                        "description": "Action 1 Description"
                      },
                      "actionId": "action1"
                    },
                    {
                      "id": "menu1",
                      "type": "menu",
                      "label": "My Menu",
                      "icons": [
                        {
                          "size": 16,
                          "file": "test_16.png"
                        },
                        {
                          "size": 32,
                          "file": "test_32.png"
                        },
                        {
                          "size": 80,
                          "file": "test_80.png"
                        }
                      ],
                      "supertip": {
                        "title": "My Menu",
                        "description": "Menu with 2 actions"
                      },
                      "items": [
                        {
                          "id": "menuItem1",
                          "type": "menuItem",
                          "label": "Action 2",
                          "supertip": {
                            "title": "Action 2 Title",
                            "description": "Action 2 Description"
                          },
                          "actionId": "action2"
                        },
                        {
                          "id": "menuItem2",
                          "type": "menuItem",
                          "label": "Action 3",
                          "icons": [
                            {
                              "size": 16,
                              "file": "test_16.png"
                            },
                            {
                              "size": 32,
                              "file": "test_32.png"
                            },
                            {
                              "size": 80,
                              "file": "test_80.png"
                            }
                          ],
                          "supertip": {
                            "title": "Action 3 Title",
                            "description": "Action 3 Description"
                          },
                          "actionId": "action3"
                        }
                      ]
                    }
                  ]
                }
              ]
            }
          ]
        },
        {
          "contexts": [ "mailRead" ],
          "tabs": [
            {
              "builtInTabId": "TabDefault",
              "groups": [
                {
                  "id": "dashboard",
                  "label": "Controls",
                  "controls": [
                    {
                      "id": "control1",
                      "type": "button",
                      "label": "Action 1",
                      "icons": [
                        {
                          "size": 16,
                          "file": "test_16.png"
                        },
                        {
                          "size": 32,
                          "file": "test_32.png"
                        },
                        {
                          "size": 80,
                          "file": "test_80.png"
                        }
                      ],
                      "supertip": {
                        "title": "Action 1 Title",
                        "description": "Action 1 Description"
                      },
                      "actionId": "action1"
                    }
                  ]
                }
              ]
            }
          ]
        }
      ],
      "autoRunEvents": [
        {
          "requirements": {
            "capabilities": [
              {
                "name": "MailBox", "minVersion": "1.10"
              }
            ]
          },
          "events": [
            {
              "type": "newMessageComposeCreated",
              "actionId": "onNewMessageComposeCreated"
            },
            {
              "type": "messageSending",
              "actionId": "onMessageSending",
              "options": {
                "sendMode": "promptUser"
              }
            }
          ]
        }
      ],
      "alternates": [
        {
          "requirements": {
            "scopes": [ "mail" ]
          },
          "prefer": {
            "comAddin": {
              "progId": "ContosoExtension"
            }
          },
          "hide": {
            "storeOfficeAddin": {
              "officeAddinId": "00000000-0000-0000-0000-000000000000",
              "assetId": "WA000000000"
            }
          }
        }
      ]
    }
  ]
}
```

## Next steps

- [Build an Outlook add-in with the unified Microsoft 365 manifest (preview)](../quickstarts/outlook-quickstart-json-manifest.md).