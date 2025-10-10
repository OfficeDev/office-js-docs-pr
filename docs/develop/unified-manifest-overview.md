---
title: Office Add-ins with the unified app manifest for Microsoft 365
description: Get an overview of the unified app manifest for Microsoft 365 for Office Add-ins and its uses.
ms.topic: overview
ms.date: 10/10/2025
ms.localizationpriority: high
---

# Office Add-ins with the unified app manifest for Microsoft 365

This article introduces the unified app manifest for Microsoft 365. It assumes that you're familiar with [Office Add-ins manifest](add-in-manifests.md).

> [!TIP]
>
> - For an overview of the add-in only manifest, see [Office Add-ins with the add-in only manifest](xml-manifest-overview.md).
> - If you're familiar with the add-in only manifest, you might get a grasp on the JSON-formatted unified manifest easier by reading [Compare the add-in only manifest with the unified manifest for Microsoft 365](json-manifest-overview.md).

Microsoft is making a number of improvements to the Microsoft 365 developer platform. These improvements provide more consistency in the development, deployment, installation, and administration of all types of extensions of Microsoft 365, including Office Add-ins. These changes are compatible with existing add-ins.

One important improvement is the ability to create a single unit of distribution for all your Microsoft 365 extensions by using the same manifest format and schema.

We've taken an important first step toward these goals by making it possible for you to create Outlook add-ins with a unified manifest for Microsoft 365.

[!include[Unified manifest host application support note](../includes/unified-manifest-support-note.md)]

> [!TIP]
> Ready to get started with the unified manifest? Begin with [Build an Outlook add-in with the unified manifest for Microsoft 365](../quickstarts/outlook-quickstart-json-manifest.md).

## Key properties of the unified manifest

The main reference documentation for the version of the unified app manifest is at [Microsoft 365 app manifest schema reference](/microsoft-365/extensibility/schema). In this article, we provide a brief description of the meaning of base properties when the App for Microsoft 365 is (or includes) an Office Add-in. This is followed by some basic documentation for the [`"extensions"`](/microsoft-365/extensibility/schema/root#extensions) property and its descendant properties. There is a full sample manifest for an add-in at [Sample unified manifest](#sample-unified-manifest).

### Base properties

Each of the base properties listed in the following table has more extensive documentation at [Microsoft 365 app manifest schema](/microsoft-365/extensibility/schema/root). Base properties not included in this table have no meaning for Office Add-ins.

|JSON property|Purpose|
|:-----|:-----|
|"$schema"| Identifies the manifest schema. |
|[`"manifestVersion"`](/microsoft-365/extensibility/schema/root#manifestversion)| Version of the manifest schema. |  
|`"id"`| GUID of the App for Microsoft 365. |
|[`"version"`](/microsoft-365/extensibility/schema/root#version)| Version of the App for Microsoft 365. The format must be `n.n.n` where each `n` can be no more than five digits.|
|[`"name"`](/microsoft-365/extensibility/schema/root#name)| Public short and long names of the App for Microsoft 365. The short name appears at the top of an add-in's task pane. |
|[`"description"`](/microsoft-365/extensibility/schema/root#description)| Public short and long descriptions of the App for Microsoft 365. |
|[`"developer"`](/microsoft-365/extensibility/schema/root#developer)| Information about the developer of the App for Microsoft 365. |
|[`"localizationInfo"`](/microsoft-365/extensibility/schema/root#localizationinfo)| Configures the default locale and other supported locales. |
|[`"validDomains"`](/microsoft-365/extensibility/schema/root#validdomains) | See [Specify safe domains](#specify-safe-domains). |
|[`"webApplicationInfo"`](/microsoft-365/extensibility/schema/root#webApplicationInfo-property)| Identifies the App for Microsoft 365's web app as it is known in Microsoft Entra ID. |
|[`"authorization"`](/microsoft-365/extensibility/schema/root#authorization)| Identifies any Microsoft Graph permissions that the add-in needs. |

### `"extensions"` property

We're working hard to complete reference documentation for the `"extensions"` property and its descendant properties. In the meantime, the following provides some basic documentation. Most, but not all, of the properties have an equivalent element (or attribute) in the add-in only manifest for add-ins. For the most part, the description, and restrictions, that apply to the XML element or attribute also apply to its JSON property equivalent in the unified manifest. The tables in the '`"extensions"` property' section of [Compare the add-in only manifest with the unified manifest for Microsoft 365](json-manifest-overview.md#extensions-property) can help you determine the XML equivalent of a JSON property.

> [!NOTE]
> This table contains only some selected representative descendant properties of `"extensions"`. *It isn't an exhaustive list of all child properties of `"extensions"`.* For the full reference of the unified manifest, see the [Microsoft 365 app manifest schema reference](/microsoft-365/extensibility/schema).

|JSON property|Purpose|
|:-----|:-----|
| [`"requirements.capabilities"`](/microsoft-365/extensibility/schema/requirements-extension-element-capabilities) | Identifies the [requirement sets](office-versions-and-requirement-sets.md#office-requirement-sets-availability) that the add-in needs to be installable. |
| [`"requirements.scopes"`](/microsoft-365/extensibility/schema/requirements-extension-element#scopes) | Identifies the Office applications in which the add-in can be installed. For example, `"mail"` means the add-in can be installed in Outlook. |
| [`"ribbons"`](/microsoft-365/extensibility/schema/element-extensions#ribbons) | The ribbons that the add-in customizes. |
| `"ribbons.contexts"` | Specifies the command surfaces that the add-in customizes. For example, `"mailRead"` or `"mailCompose"`. |
| `"ribbons.fixedControls"` | Configures and adds the button of an [integrated spam-reporting](../outlook/spam-reporting.md) add-in to the Outlook ribbon. |
| `"ribbons.spamPreProcessingDialog"` | Configures the preprocessing dialog shown after the button of a spam-reporting add-in is selected from the Outlook ribbon. |
| `"ribbons.tabs"` | Configures custom ribbon tabs. |
| [`"alternates"`](/microsoft-365/extensibility/schema/element-extensions#alternates) | Specifies backwards compatibility with an equivalent COM add-in, XLL, or both. Also specifies the main icons that are used to represent the add-in on older versions of Office. |
| [`"runtimes"`](/microsoft-365/extensibility/schema/element-extensions#runtimes)  | Configures the [embedded runtimes](../testing/runtimes.md) that the add-in uses, including various kinds of add-ins that have little or no UI, such as custom function-only add-ins and [function commands](../design/add-in-commands.md#types-of-add-in-commands). |
| [`"autoRunEvents"`](/microsoft-365/extensibility/schema/element-extensions#autorunevents) | Configures an event handler for a specified event. |
| [`"keyboardShortcuts"`](/microsoft-365/extensibility/schema/element-extensions#keyboardshortcuts) (developer preview) | Defines custom keyboard shortcuts or key combinations to run specific actions. |

## Specify safe domains

There is a `"validDomains"` array in the manifest file that is used to tell Office which domains your add-in should be allowed to navigate to. As noted in [Specify domains you want to open in the add-in window](add-in-manifests.md#specify-domains-you-want-to-open-in-the-add-in-window), when running in Office on the web, your task pane can be navigated to any URL. However, in desktop platforms, if your add-in tries to go to a URL in a domain other than the domain that hosts the start page, that URL opens in a new browser window outside the add-in pane of the Office application.

To override this behavior in desktop platforms, add each domain you want to open in the add-in window to the list of domains specified in the `"validDomains"` array. If the add-in tries to go to a URL in a domain that is in the list, then it opens in the task pane in both Office on the web and desktop. If it tries to go to a URL that isn't in the list, then in Office on desktop, that URL opens in a new browser window (outside the add-in task pane).

## Client and platform support

Add-ins that use the unified manifest can be installed if the Office platform *directly* supports it.

To run an add-in on platforms that don't directly support the unified manifest, you must publish the add-in to [Microsoft Marketplace](https://marketplace.microsoft.com/). Then, deploy the add-in in the [Microsoft 365 admin center](../publish/publish.md). This way, an add-in only manifest is generated from the unified manifest and stored. The add-in only manifest is then used to install the add-in on platforms that don't directly support the unified manifest.

The following tables lists which Office platforms directly support add-ins that use the unified manifest.

| Client/platform | Support for add-ins with the unified manifest|
| ----- | ----- |
| Office on the web | Directly supported |
| Office on Windows (Version 2304 (Build 16320.00000) or later) connected to a Microsoft 365 subscription | Directly supported |
| [new Outlook on Windows](https://support.microsoft.com/office/656bb8d9-5a60-49b2-a98b-ba7822bc7627) | Directly supported |
| Office on Windows (prior to Version 2304 (Build 16320.00000)) connected to a Microsoft 365 subscription | Not directly supported |
| Office on Windows (perpetual versions) | Not directly supported |
| Office on Mac | Not directly supported |
| Office on mobile | Not directly supported |

> [!NOTE]
> If you're deploying an add-in in the [Microsoft 365 admin center](../publish/publish.md#integrated-apps-portal-in-the-microsoft-365-admin-center) and require it to run on platforms that don't directly support the unified manifest, the add-in must be a published Microsoft Marketplace add-in. Custom add-ins or line-of-business (LOB) add-ins that use the unified manifest can be deployed in the [Integrated apps portal](/microsoft-365/admin/manage/test-and-deploy-microsoft-365-apps) of the Microsoft 365 admin center, but they won't be installable on Office versions that don't directly support the unified manifest.

## Sample unified manifest

The following is an example of a unified app manifest for an add-in. It doesn't contain every possible manifest property.

```json
{
  "$schema": "https://developer.microsoft.com/json-schemas/teams/vDevPreview/MicrosoftTeams.schema.json",
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
                "name": "MailBox",
                "minVersion": "1.10"
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
                      "type": "mobileButton",
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
                      "type": "mobileButton",
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
- [Microsoft 365 app manifest schema reference](/microsoft-365/extensibility/schema)
