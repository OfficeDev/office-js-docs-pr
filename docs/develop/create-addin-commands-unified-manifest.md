---
title: Create add-in commands with the unified manifest for Microsoft 365
description: Configure the unified manifest for Microsoft 365 to define add-in commands for Excel, Outlook, PowerPoint, and Word. Use add-in commands to create UI elements, add buttons or lists, and perform actions.
ms.date: 05/19/2025
ms.localizationpriority: medium
---

# Create add-in commands with the unified manifest for Microsoft 365

Add-in commands provide an easy way to customize the default Office user interface (UI) with specified UI elements that perform actions. For an introduction to add-in commands, see [Add-in commands](../design/add-in-commands.md).

This article describes how to configure the [Unified manifest for Microsoft 365](unified-manifest-overview.md) to define add-in commands and how to create the code for [function commands](../design/add-in-commands.md#types-of-add-in-commands).

[!include[Unified manifest host application support note](../includes/unified-manifest-support-note.md)]

> [!TIP]
> Instructions for creating add-in commands with the add-in only manifest are in [Create add-in commands with the add-in only manifest](create-addin-commands.md).

[!INCLUDE [non-unified manifest clients note](../includes/non-unified-manifest-clients.md)]

## Starting point and major steps

Both of the tools that create add-in projects with a unified manifest &#8212; the [Office Yeoman generator](yeoman-generator-overview.md) and [Microsoft 365 Agents Toolkit](agents-toolkit-overview.md) &#8212; create projects with one or more add-in commands. The only time you won't already have an add-in command is if you are updating an add-in which previously didn't have one.

## Two decisions

- Decide which of two types of add-in commands you need: [Task pane or function](../design/add-in-commands.md#types-of-add-in-commands)
- Decide which kind of UI element you need: button or menu item. Then carry out the steps in the sections and subsections below that correspond to your decisions.

## Add a task pane command

The following subsections explain how to include a [task pane command](../design/add-in-commands.md#types-of-add-in-commands) in an add-in.

### Configure the runtime for the task pane command

1. Open the unified manifest and find the [`"extensions.runtimes"`](/microsoft-365/extensibility/schema/extension-runtimes-array?view=m365-app-prev&preserve-view=true) array.
1. Ensure that there is a runtime object that has an [`"actions.type"`](/microsoft-365/extensibility/schema/extension-runtimes-actions-item#type) property with the value `"openPage"`. This type of runtime opens a task pane.
1. Ensure that the [`"requirements.capabilities"`](/microsoft-365/extensibility/schema/requirements-extension-element-capabilities) array contains an object that specifies a [Requirement Set](office-versions-and-requirement-sets.md) that supports add-in commands. For Outlook the minimum requirement set for add-in commands is [Mailbox 1.3](/javascript/api/requirement-sets/outlook/requirement-set-1.3/outlook-requirement-set-1.3). For other Office host applications, the minimum requirement set for add-in commands is [AddinCommands 1.1](/javascript/api/requirement-sets/common/add-in-commands-requirement-sets).

1. Ensure that the `"id"` of the runtime object has a descriptive name such as `"TaskPaneRuntime"`.
1. Ensure that the [`"code.page"`](/microsoft-365/extensibility/schema/extension-runtime-code#page) property of the runtime object is set to the URL of the page that should open in the task pane, such as `"https://localhost:3000/taskpane.html"`.
1. Ensure that the [`"actions.view"`](/microsoft-365/extensibility/schema/extension-runtimes-actions-item#view) of the runtime object has a name that describes the content of the page that you set in the preceding step, such as `"homepage"` or `"dashboard"`.
1. Ensure that the [`"actions.id"`](/microsoft-365/extensibility/schema/extension-runtimes-actions-item#id) of the runtime object has a descriptive name such as `"ShowTaskPane"` that indicates what happens when the user selects the add-in command button or menu item.
1. Set the other properties and subproperties of the runtime object as shown in the following completed example of a runtime object. The `"type"` and `"lifetime"` properties are required and in Outlook Add-ins. They always have the values shown in this example.

    ```json
    "runtimes": [
        {
            "requirements": {
                "capabilities": [
                    {
                        "name": "Mailbox",
                        "minVersion": "1.3"
                    }
                ]
            },
            "id": "TaskPaneRuntime",
            "type": "general",
            "code": {
                "page": "https://localhost:3000/taskpane.html"
            },
            "lifetime": "short",
            "actions": [
                {
                    "id": "ShowTaskPane",
                    "type": "openPage",
                    "view": "homepage"
                }
            ]
        }
    ]
    ```

### Configure the UI for the task pane command

1. Ensure that the extension object for which you configured a runtime has a [`"ribbons"`](/microsoft-365/extensibility/schema/element-extensions#ribbons) array property as a peer to the [`"runtimes"`](/microsoft-365/extensibility/schema/element-extensions#runtimes) array. There is typically only one extension object in the [`"extensions"`](/microsoft-365/extensibility/schema/root#extensions) array.
1. Ensure that the array has an object with array properties named `"contexts"` and `"tabs"`, as shown in the following example.

    ```json
    "ribbons": [
        {
            "contexts": [
                // child objects omitted
            ],
            "tabs": [
                // child objects omitted
            ]
        }
    ]
    ```

1. Ensure that the `"contexts"` array has strings that specify the windows or panes in which the UI for the task pane command should appear. For example, `"mailRead"` means that it will appear in the reading pane or message window when an email message is open, but `"mailCompose"` means it will appear when a new message or a reply is being composed. The following are the allowable values:

    - `"mailRead"`
    - `"mailCompose"`
    - `"meetingDetailsOrganizer"`
    - `"meetingDetailsAttendee"`

    The following is an example.

    ```json
    "contexts": [
        "mailRead"
    ],
    ```

1. Ensure that the `"tabs"` array has an object with a `"builtInTabId"` string property that is set to the ID of ribbon tab in which you want your task pane command to appear. Also, ensure that there is a `"groups"` array with at least one object in it. The following is an example.

    ```json
    "tabs": [
        {
            "builtInTabID": "TabDefault",
            "groups": [
                {
                    // properties omitted
                }
            ]
        }
    ]
    ```

    > [!NOTE]
    > For a list of the possible values of the `"builtInTabID"` property, see [Find the IDs of built-in Office ribbon tabs](built-in-ui-ids.md).

1. Ensure that the `"groups"` array has an object to define the custom control group that will hold your add-in command UI controls. The following is an example. Note the following about this JSON:

    - The `"id"` must be unique across all groups in all ribbon objects in the manifest. Maximum length is 64 characters.
    - The `"label"` appears on the group on the ribbon. Although its maximum length is 64 characters, to ensure that the control group fits correctly in the ribbon, we recommend that you limit the `"label"` to 16 characters.
    - One of the `"icons"` appears on the group only if the Office application window, and hence the ribbon, has been sized by the user too small for any of the controls in the group to appear. Office decides when to use one of these icons and which one to use based on the size of the window and the resolution of the device. You cannot control this. You must provide image files for 16, 32, and 80 pixels, while five other sizes are also supported (20, 24, 40, 48, and 64 pixels). You must use Secure Sockets Layer (SSL) for all URLs.

    ```json
    "groups": [
        {
            "id": "msgReadGroup",
            "label": "Contoso Add-in",
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
                    // properties omitted
                }
            ]
        }
    ]
    ```

1. Ensure that there is a control object in the `"controls"` array for each button or custom menu you want. The following is an example. Note the following about this JSON:

    - The `"id"`, `"label"`, and `"icons"` properties have the same purpose and the same restrictions as the corresponding properties of a group object, except that they apply to a specific button or menu within the group.
    - The `"type"` property is set to `"button"` which means that the control will be a ribbon button. You can also configure a task pane command to be run from a menu item. See [Menu and menu items](#menu-and-menu-items).
    - The `"supertip.title"` (maximum length: 64 characters) and `"supertip.description"` (maximum length: 128 characters) appear when the cursor is hovering over the button or menu.
    - The `"actionId"` must be an exact match for the [`"runtimes.actions.id"`](/microsoft-365/extensibility/schema/extension-runtimes-actions-item#id) that you set in [Configure the runtime for the task pane command](#configure-the-runtime-for-the-task-pane-command).

    ```json
    {
        "id": "msgReadOpenPaneButton",
        "type": "button",
        "label": "Show Task Pane",
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
            "title": "Show Contoso Task Pane",
            "description": "Opens the Contoso task pane."
        },
        "actionId": "ShowTaskPane"
    }
    ```

You've now completed adding a task pane command to your add-in. [Sideload and test it](../testing/test-debug-office-add-ins.md#sideload-an-office-add-in-for-testing).

## Add a function command

The following subsections explain how to include a [function command](../design/add-in-commands.md#types-of-add-in-commands) in an add-in.

### Create the code for the function command

1. Ensure that your source code includes a JavaScript or Typescript file with the function that you want to run with your function command. The following is an example. Since this article is about creating add-in commands, and not about teaching the Office JavaScript Library, we provide it with minimal comments, but do note the following:

    - For purposes of this article, the file is named **commands.js**.
    - The function will cause a small notification to appear on an open email message with the text "Action performed".
    - Like all code that call APIs in the Office JavaScript Library, it must begin by [initializing the library](initialize-add-in.md). It does this by calling `Office.onReady`.
    - The last thing the code calls is [Office.actions.associate](/javascript/api/office/office.actions#office-office-actions-associate-member(1)) to tell Office which function in the file should be run when the UI for your function command is invoked. The function maps the function name to an action ID that you configure in the manifest in a later step. If you define multiple function commands in the same file, your code must call `associate` for each one.
    - The function must take a parameter of type [Office.AddinCommands.Event](/javascript/api/office/office.addincommands.event). The last line of the function must call [event.completed](/javascript/api/office/office.addincommands.event#office-office-addincommands-event-completed-member(1)).

    ```javascript
    Office.onReady(function() {
    // Add any initialization code here.
    });

    function setNotification(event) {
    const message = {
        type: Office.MailboxEnums.ItemNotificationMessageType.InformationalMessage,
        message: "Performed action.",
        icon: "Icon.80x80",
        persistent: true,
    };

    // Show a notification message.
    Office.context.mailbox.item.notificationMessages.replaceAsync("ActionPerformanceNotification", message);

    // Be sure to indicate when the add-in command function is complete.
    event.completed();
    }
    
    // Map the function to the action ID in the manifest.
    Office.actions.associate("SetNotification", setNotification);
    ```

1. Ensure that your source code includes an HTML file that is configured to load the function file you created. The following is an example. Note the following about this JSON:

    - For purposes of this article, the file is named **commands.html**.
    - The `<body>` element is empty because the file has no UI. Its only purpose is to load JavaScript files.
    - The Office JavaScript Library and the **commands.js** file that you created in the preceding step is explicitly loaded.

        > [!NOTE]
        > It's common in Office Add-in development to use tools like [webpack](https://webpack.js.org/) and its plugins to automatically inject `<script>` tags into HTML files at build time. If you use such a tool, you shouldn't include any `<script>` tags in your source file that are going to be inserted by the tool.

    ```html
    <!DOCTYPE html>
    <html>
        <head>
            <meta charset="UTF-8" />
            <meta http-equiv="X-UA-Compatible" content="IE=Edge" />

            <!-- Office JavaScript Library -->
            <script type="text/javascript" src="https://appsforoffice.microsoft.com/lib/1/hosted/office.js"></script>
            <!-- Function command file -->
            <script src="commands.js" type="text/javascript"></script>
        </head>
        <body>
        </body>
    </html>
    ```

### Configure the runtime for the function command

1. Open the unified manifest and find the `"extensions.runtimes"` array.
1. Ensure that there is a runtime object that has a `"actions.type"` property with the value `"executeFunction"`.
1. Ensure that the `"requirements.capabilities"` array contains objects that specify any [Requirement Sets](office-versions-and-requirement-sets.md) that are needed to support the APIs add-in commands. For Outlook, the minimum requirement set for add-in commands is [Mailbox 1.3](/javascript/api/requirement-sets/outlook/requirement-set-1.3/outlook-requirement-set-1.3). But if your function command calls that API that is part of later **Mailbox** requirement set, such as **Mailbox 1.5**, then you need to specify the later version (e.g., "1.5") as the `"minVersion"` value. For other Office host applications, the minimum requirement set for add-in commands is [AddinCommands 1.1](/javascript/api/requirement-sets/common/add-in-commands-requirement-sets).

1. Ensure that the `"id"` of the runtime object has a descriptive name such as "CommandsRuntime".
1. Ensure that the [`"code.page"`](/microsoft-365/extensibility/schema/extension-runtime-code#page) property of the runtime object is set to the URL of the UI-less HTML page that loads your function file, such as `"https://localhost:3000/commands.html"`.
1. Ensure that the `"actions.id"` of the runtime object has a descriptive name such as "SetNotification" that indicates what happens when the user selects the add-in command button or menu item.

    > [!IMPORTANT]
    > The value of `"actions.id"` must exactly match the first parameter of the call to `Office.actions.associate` in the function file.

1. Set the other properties and subproperties of the runtime object as shown in the following completed example of a runtime object.

    ```json
    "runtimes": [
        {
            "id": "CommandsRuntime",
            "type": "general",
            "code": {
                "page": "https://localhost:3000/commands.html"
            },
            "lifetime": "short",
            "actions": [
                {
                    "id": "SetNotification",
                    "type": "executeFunction",
                }
            ]
        }       
    ]
    ```

### Configure the UI for the function command

1. Ensure that the extension object for which you configured a runtime has a `"ribbons"` array property as a peer to the `"runtimes"` array. There is typically only one extension object in the `"extensions"` array.
1. Ensure that the array has an object with array properties named `"contexts"` and `"tabs"`, as shown in the following example.

    ```json
    "ribbons": [
        {
            "contexts": [
                // child objects omitted
            ],
            "tabs": [
                // child objects omitted
            ]
        }
    ]
    ```

1. Ensure that the `"contexts"` array has strings that specify the windows or panes in which the UI for the function command should appear. For example, `"mailRead"` means that it will appear in the reading pane or message window when an email message is open, but `"mailCompose"` means it will appear when a new message or a reply is being composed. The following are the allowable values:

    - `"mailRead"`
    - `"mailCompose"`
    - `"meetingDetailsOrganizer"`
    - `"meetingDetailsAttendee"`

    The following is an example.

    ```json
    "contexts": [
        "mailRead"
    ],
    ```

1. Ensure that the `"tabs"` array has an object with a `"builtInTabId"` string property that is set to the ID of ribbon tab in which you want your function command to appear and a `"groups"` array with at least one object in it. The following is an example.

    ```json
    "tabs": [
        {
            "builtInTabID": "TabDefault",
            "groups": [
                {
                    // properties omitted
                }
            ]
        }
    ]
    ```

    > [!NOTE]
    > For a list of the possible values of the `"builtInTabID"` property, see [Find the IDs of built-in Office ribbon tabs](built-in-ui-ids.md).

1. Ensure that the `"groups"` array has an object to define the custom control group that will hold your add-in command UI controls. The following is an example. Note the following about this JSON:

    - The `"id"` must be unique across all groups in all ribbon objects in the manifest. Maximum length is 64 characters.
    - The `"label"` appears on the group on the ribbon. Although its maximum length is 64 characters, to ensure that the control group fits correctly in the ribbon, we recommend that you limit the `"label"` to 16 characters.
    - One of the `"icons"` appears on the group only if the Office application window, and hence the ribbon, has been sized by the user too small for any of the controls in the group to appear. Office decides when to use one of these icons and which one to use based on the size of the window and the resolution of the device. You cannot control this. You must provide image files for 16, 32, and 80 pixels, while five other sizes are also supported (20, 24, 40, 48, and 64 pixels). You must use Secure Sockets Layer (SSL) for all URLs.

    ```json
    "groups": [
        {
            "id": "msgReadGroup",
            "label": "Contoso Add-in",
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
                    // properties omitted
                }
            ]
        }
    ]
    ```

1. Ensure that there is a control object in the `"controls"` array for each button or custom menu you want. The following is an example. Note the following about this JSON:

    - The `"id"`, `"label"`, and `"icons"` properties have the same purpose and the same restrictions as the corresponding properties of a group object, except that they apply to a specific button or menu within the group.
    - The `"type"` property is set to `"button"` which means that the control will be a ribbon button. You can also configure a function command to be run from a menu item. See [Menu and menu items](#menu-and-menu-items).
    - The `"supertip.title"` (maximum length: 64 characters) and `"supertip.description"` (maximum length: 128 characters) appear when the cursor is hovering over the button or menu.
    - The `"actionId"` must be an exact match for the `"runtime.actions.id"` that you set in [Configure the runtime for the function command](#configure-the-runtime-for-the-function-command).

    ```json
    {
        "id": "msgReadSetNotificationButton",
        "type": "button",
        "label": "Set Notification",
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
            "title": "Set Notification",
            "description": "Displays a notification message on the current message."
        },
        "actionId": "SetNotification"
    }
    ```

You've now completed adding a function command to your add-in. [Sideload and test it](../testing/test-debug-office-add-ins.md#sideload-an-office-add-in-for-testing).

## Menu and menu items

In addition to custom buttons, you can also add custom drop down menus to the Office ribbon. This section explains how by using an example with two menu items. One invokes a task pane command. The other invokes a function command.

### Configure the runtimes and code

Carry out the steps of the following sections:

- [Configure the runtime for the task pane command](#configure-the-runtime-for-the-task-pane-command)
- [Create the code for the function command](#create-the-code-for-the-function-command)
- [Configure the runtime for the function command](#configure-the-runtime-for-the-function-command)

### Configure the UI for the menu

1. Ensure that the extension object for which you configured a runtime has a `"ribbons"` array property as a peer to the `"runtimes"` array. There is typically only one extension object in the `"extensions"` array.
1. Ensure that the array has an object with array properties named `"contexts"` and `"tabs"`, as shown in the following example.

    ```json
    "ribbons": [
        {
            "contexts": [
                // child objects omitted
            ],
            "tabs": [
                // child objects omitted
            ]
        }
    ]
    ```

1. Ensure that the `"contexts"` array has strings that specify the windows or panes in which the menu should appear on the ribbon. For example, `"mailRead"` means that it will appear in the reading pane or message window when an email message is open, but `"mailCompose"` means it will appear when a new message or a reply is being composed. The following are the allowable values:

    - `"mailRead"`
    - `"mailCompose"`
    - `"meetingDetailsOrganizer"`
    - `"meetingDetailsAttendee"`

    The following is an example.

    ```json
    "contexts": [
        "mailRead"
    ],
    ```

1. Ensure that the `"tabs"` array has an object with a `"builtInTabId"` string property that is set to the ID of ribbon tab in which you want your task pane command to appear and a `"groups"` array with at least one object in it. The following is an example.

    ```json
    "tabs": [
        {
            "builtInTabID": "TabDefault",
            "groups": [
                {
                    // properties omitted
                }
            ]
        }
    ]
    ```

    > [!NOTE]
    > For a list of the possible values of the `"builtInTabID"` property, see [Find the IDs of built-in Office ribbon tabs](built-in-ui-ids.md).

1. Ensure that the `"groups"` array has an object to define the custom control group that will hold your drop down menu control. The following is an example. Note the following about this JSON:

    - The `"id"` must be unique across all groups in all ribbon objects in the manifest. Maximum length is 64 characters.
    - The `"label"` appears on the group on the ribbon. Although its maximum length is 64 characters, to ensure that the control group fits correctly in the ribbon, we recommend that you limit the `"label"` to 16 characters.
    - One of the `"icons"` appears on the group only if the Office application window, and hence the ribbon, has been sized by the user too small for any of the controls in the group to appear. Office decides when to use one of these icons and which one to use based on the size of the window and the resolution of the device. You cannot control this. You must provide image files for 16, 32, and 80 pixels, while five other sizes are also supported (20, 24, 40, 48, and 64 pixels). You must use Secure Sockets Layer (SSL) for all URLs.

    ```json
    "groups": [
        {
            "id": "msgReadGroup",
            "label": "Contoso Add-in",
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
                    // properties omitted
                }
            ]
        }
    ]
    ```

1. Ensure that there is a control object in the `"controls"` array. The following is an example. Note the following about this JSON:

    - The `"id"`, `"label"`, and `"icons"` properties have the same purpose and the same restrictions as the corresponding properties of a group object, except that they apply to the drop down menu within the group.
    - The `"type"` property is set to `"menu"` which means that the control will be a drop down menu.
    - The `"supertip.title"` (maximum length: 64 characters) and `"supertip.description"` (maximum length: 128 characters) appear when the cursor is hovering over the menu.
    - The `"items"` property contains the JSON for the two menu options. The values are added in later steps.

    ```json
    {
        "id": "msgReadMenu",
        "type": "menu",
        "label": "Contoso Menu",
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
            "title": "Show Contoso Actions",
            "description": "Opens the Contoso menu."
        },
        "items": [
            {
                "id": "",
                "type": "",
                "label": "",
                "supertip": {},
                "actionId": ""
            },
            {
                "id": "",
                "type": "",
                "label": "",
                "supertip": {},
                "actionId": ""
            }
        ]
    }
    ```

1. The first item shows a task pane. The following is an example. Note the following about this code:

    - The `"id"`, `"label"`, and `"supertip"` properties have the same purpose and the same restrictions as the corresponding properties of the parent menu object, except that they apply to just this menu option.
    - The `"icons"` property is optional for menu items and there isn't one in this example. If you include one, it has the same purposes and restrictions as the `"icons"` property of the parent menu, except that the icon appears on the menu item beside the label.
    - The `"type"` property is set to `"menuItem"`.
    - The `"actionId"` must be an exact match for the `"runtimes.actions.id"` that you set in [Configure the runtime for the task pane command](#configure-the-runtime-for-the-task-pane-command).

    ```json
    {
        "id": "msgReadOpenPaneMenuItem",
        "type": "menuItem",
        "label": "Show Task Pane",
        "supertip": {
            "title": "Show Contoso Task Pane",
            "description": "Opens the Contoso task pane."
        },
        "actionId": "ShowTaskPane"
    },
    ```

1. The second item runs a function command. The following is an example. Note the following about this code:

    - The `"actionId"` must be an exact match for the `"runtimes.actions.id"` that you set in [Configure the runtime for the function command](#configure-the-runtime-for-the-function-command).

    ```json
    {
        "id": "msgReadSetNotificationMenuItem",
        "type": "menuItem",
        "label": "Set Notification",
        "supertip": {
            "title": "Set Notification",
            "description": "Displays a notification message on the current message."
        },
        "actionId": "SetNotification"
    }
    ```

You've now completed adding a menu to your add-in. [Sideload and test it](../testing/test-debug-office-add-ins.md#sideload-an-office-add-in-for-testing).

## See also

- [Add-in commands](../design/add-in-commands.md)
- [Unified manifest for Microsoft 365](unified-manifest-overview.md)
