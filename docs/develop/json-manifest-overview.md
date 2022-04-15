---
title: JSON-formatted manifest for Office Add-ins
description: Get an overview of the preview JSON manifest.
ms.date: 05/24/2022
ms.localizationpriority: high
---

# JSON-formatted manifest for Office Add-ins (preview)

Microsoft of making a number of improvements to the Microsoft 365 developer platform. These improvements will provide more consistency in the development, deployment, installation, and administration of all types of extension of Micrsoft 365, including Office Add-ins. These changes are not breaking, so no existing extension will be broken. 

Two of the most important improvements are:

- It will be possible to surface a single web app as multiple types of Microsoft 365 extensions. For example a web app can be both an Office Add-in and a custom tab in Teams.
- All types of Microsoft 365 extensions will use the same manifest format (JSON) and schema. It will be based on the current Teams manifest schema. In support of the first bullet, it will be possible to specify multiple types of extensions in the manifest.  

The new manifest is available for preview and we encourage experienced add-in developers to experiment with it. It should not be used in production add-ins. During the early preview period, the following limitations apply:

- The preview JSON manifest only supports Outlook add-ins. We're working on extending support to Excel, PowerPoint, and Word.
- It is not yet possible to combine an add-in with other Microsoft 365 extension types, such as a Teams tab. We're working on this too.

## Conceptual mapping of the preview JSON and current XML manifests

This section describes the preview JSON manifest for readers that are familiar with the current XML manifest.

### Schemas

There are a total of 7 [Schemas](/openspecs/office_file_formats/ms-owemxml/c6a06390-34b8-4b42-82eb-b28be12494a8) to define the current XML manifest. There is just one schema for the [preview JSON manifest](/microsoftteams/platform/resources/dev-preview/developer-preview-intro.md). 

For the most part ... mapping

"keyboards" property incorps keyboard shortcuts json

Req sets are called "capabilities".

Associated COM app -> "alternatives"


## Sample preview JSON manifest

The following is an example of a preview JSON-manifest for an Outlook Add-in.

```json
{
    "$schema": "../../schema/metaos.public.schema.json",
    "id": "82a9d9c3-4702-4322-bbc4-6fe7f9b01483",
    "version": "1.0.0",
    "manifestVersion": "m365DevPreview",
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