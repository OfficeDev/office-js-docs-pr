---
title: Office Add-ins with the unified app manifest for Microsoft 365
description: Get an overview of the unified app manifest for Microsoft 365 for Office Add-ins and its uses.
ms.topic: overview
ms.date: 04/12/2024
ms.localizationpriority: high
---

# Office Add-ins with the unified app manifest for Microsoft 365

This article introduces the unified app manifest for Microsoft 365. It assumes that you're familiar with [Office Add-ins manifest](add-in-manifests.md).

> [!TIP]
> - For an overview of the add-in only manifest, see [Office Add-ins with the add-in only manifest](xml-manifest-overview.md).
> - If you're familiar with the add-in only manifest, you might get a grasp on the JSON-formatted unified manifest easier by reading [Compare the add-in only manifest with the unified manifest for Microsoft 365](json-manifest-overview.md).

Microsoft is making a number of improvements to the Microsoft 365 developer platform. These improvements provide more consistency in the development, deployment, installation, and administration of all types of extensions of Microsoft 365, including Office Add-ins. These changes are compatible with existing add-ins.

One important improvement is the ability to create a single unit of distribution for all your Microsoft 365 extensions by using the same manifest format and schema.

We've taken an important first step toward these goals by making it possible for you to create Outlook add-ins with a unified manifest for Microsoft 365.

> [!NOTE]
> - The unified manifest currently only supports Outlook add-ins and only in Office linked to a Microsoft 365 subscription and installed on Windows, on a mobile device, or in Outlook on the web. We're working on extending support to Excel, PowerPoint, and Word, as well as to Outlook on Mac, and to perpetual versions of Office.
> - The unified manifest requires Office Version 2304 (Build 16320.00000) or later.

> [!TIP]
> Ready to get started with the unified manifest? Begin with [Build an Outlook add-in with the unified manifest for Microsoft 365](../quickstarts/outlook-quickstart-json-manifest.md).

## Key properties of the unified manifest

The main reference documentation for the version of the unified app manifest is at [Manifest schema](/microsoftteams/platform/resources/schema/manifest-schema). That article provides information about the critical base manifest properties, but may not include any documentation of the "extensions" property, which is the property where Office Add-ins are configured in the unified manifest. So, in this article, we provide a brief description of the meaning of base properties when the Teams App is (or includes) an Office add-in. This is followed by some basic documentation for the "extensions" property and its descendent properties. There is a full sample manifest for an add-in at [Sample unified manifest](#sample-unified-manifest).

### Base properties

Each of the base properties listed in the following table has more extensive documentation at [Manifest schema](/microsoftteams/platform/resources/schema/manifest-schema). Base properties not included in this table have no meaning for Office Add-ins.

|JSON property|Purpose|
|:-----|:-----|
|"$schema"| Identifies the manifest schema. | 
|"manifestVersion"| Version of the manifest schema. |  
|"id"| GUID of the Teams app/add-in. |
|"version"| Version of the Teams app/add-in. The format must be `n.n.n` where each `n` can be no more than five digits.| 
|"name"| Public short and long names of the Teams app/add-in. The short name appears at the top of an add-in's task pane. | 
|"description"| Public short and long descriptions of the Teams app/add-in. | 
|"developer"| Information about the developer of the Teams app/add-in. | 
|"localizationInfo"| Configures the default locale and other supported locales. | 
|"validDomains" | See [Specify safe domains](#specify-safe-domains). |
|"webApplicationInfo"| Identifies the Teams app/add-in's web app as it is known in Azure Active Directory. | 
|"authorization"| Identifies any Microsoft Graph permissions that the add-in needs. |

### "extensions" property

We're working hard to complete reference documentation for the "extensions" property and its descendent properties. In the meantime, the following provides some basic documentation. Most, but not all, of the properties have an equivalent element (or attribute) in the add-in only manifest for add-ins. For the most part, the description, and restrictions, that apply to the XML element or attribute also apply to its JSON property equivalent in the unified manifest. The tables in the '"extensions" property' section of [Compare the add-in only manifest with the unified manifest for Microsoft 365](json-manifest-overview.md#extensions-property) can help you determine the XML equivalent of a JSON property.


|JSON property|Purpose|
|:-----|:-----|:-----|:-----|
| "requirements.capabilities" | Identifies the [requirement sets](office-versions-and-requirement-sets.md#office-requirement-sets-availability) that the add-in needs to be installable. |
| "requirements.scopes" | Identifies the Office applications in which the add-in can be installed. For example, "mail" means the add-in can be installed in Outlook. | 
| "ribbons" | The ribbons that the add-in customizes. | 
| "ribbons.contexts" | Specifies the command surfaces that the add-in customizes. For example, "mailRead" or "mailCompose". |
| "ribbons.tabs" | Configures custom ribbon tabs. |
| "alternates" | Specifies backwards compatibility with an equivalent COM add-in, XLL, or both. Also specifies the main icons that are used to represent the add-in on older versions of Office. | 
| "runtimes"  | Configures the [embedded runtimes](../testing/runtimes.md) that the add-in uses, including various kinds of add-ins that have little or no UI, such as custom function-only add-ins and [function commands](../design/add-in-commands.md#types-of-add-in-commands). | 
| "autoRunEvents" | Configures an event handler for a specified event. | 

## Specify safe domains

There is a "validDomains" array in the manifest file that is used to tell Office which domains your add-in should be allowed to navigate to. As noted in [Specify domains you want to open in the add-in window](add-in-manifests.md#specify-domains-you-want-to-open-in-the-add-in-window), when running in Office on the web, your task pane can be navigated to any URL. However, in desktop platforms, if your add-in tries to go to a URL in a domain other than the domain that hosts the start page, that URL opens in a new browser window outside the add-in pane of the Office application.

To override this behavior in desktop platforms, add each domain you want to open in the add-in window to the list of domains specified in the "validDomains" array. If the add-in tries to go to a URL in a domain that is in the list, then it opens in the task pane in both Office on the web and desktop. If it tries to go to a URL that isn't in the list, then in Office on desktop, that URL opens in a new browser window (outside the add-in task pane).

## Sample unified manifest

The following is an example of a unified app manifest for an add-in.

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
                          "url": "test_16.png"
                        },
                        {
                          "size": 32,
                          "url": "test_32.png"
                        },
                        {
                          "size": 80,
                          "url": "test_80.png"
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
                          "url": "test_16.png"
                        },
                        {
                          "size": 32,
                          "url": "test_32.png"
                        },
                        {
                          "size": 80,
                          "url": "test_80.png"
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
                              "url": "test_16.png"
                            },
                            {
                              "size": 32,
                              "url": "test_32.png"
                            },
                            {
                              "size": 80,
                              "url": "test_80.png"
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
              ],
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
                          "url": "test_16.png"
                        },
                        {
                          "size": 32,
                          "url": "test_32.png"
                        },
                        {
                          "size": 80,
                          "url": "test_80.png"
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
              ],
              "customMobileRibbonGroups" [
                {
                  "id": "myMobileGroup",
                  "label": "Contoso Actions",
                  "controls": [
                    {
                      "id": "msgReadFunctionButton",
                      "type": "MobileButton",
                      "label": "Action 1",
                      "icons": [
                        {
                          "size": 16,
                          "url": "test_16.png"
                        },
                        {
                          "size": 32,
                          "url": "test_32.png"
                        },
                        {
                          "size": 80,
                          "url": "test_80.png"
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
              "customMobileRibbonGroups": [
                {
                  "id": "mobileDashboard",
                  "label": "Controls",
                  "controls": [
                    {
                      "id": "control1",
                      "type": "MobileButton",
                      "label": "Action 1",
                      "icons": [
                        {
                          "size": 16,
                          "url": "test_16.png"
                        },
                        {
                          "size": 32,
                          "url": "test_32.png"
                        },
                        {
                          "size": 80,
                          "url": "test_80.png"
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
          },
          "alternateIcons": {
            "icon": {
              "size": 64,
              "url": "https://contoso.com/assets/icon64x64.jpg"
            },
            "highResolutionIcon": {
              "size": 64,
              "url": "https://contoso.com/assets/icon128x128.jpg"
            }
          }
        }
      ]
    }
  ]
}
```

## See also

- [Create add-in commands with the unified manifest for Microsoft 365](create-addin-commands-unified-manifest.md)
- [Manifest schema](/microsoftteams/platform/resources/schema/manifest-schema)
