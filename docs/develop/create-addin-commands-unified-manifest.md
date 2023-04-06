---
title: Create add-in commands with the unified Microsoft 365 manifest
description: Configure the unified Microsoft 365 manifest to define add-in commands for Excel, Outlook, PowerPoint, and Word. Use add-in commands to create UI elements, add buttons or lists, and perform actions.
ms.date: 04/04/2023
ms.localizationpriority: medium
---

# Create add-in commands with the unified Microsoft 365 manifest

Add-in commands provide an easy way to customize the default Office user interface (UI) with specified UI elements that perform actions. For an introduction to add-in commands, see [Add-in commands](../design/add-in-commands.md).

This article describes how to edit the unified manifest to define add-in commands and how to create the code for [function commands](../design/add-in-commands.md#types-of-add-in-commands). 

> [!TIP]
> Instructions for creating add-in commands with the XML manifest are in [Create add-in commands with the XML manifest](create-addin-commands.md).

> [!NOTE]
> The unified manifest is currently supported only for Outlook Add-ins. It is in preview and should not be used for production add-ins.

## Starting point and major steps

Both of the tools that create add-in projects with a unified manifest &#8212; the [Office Yeoman generator](yeoman-generator-overview.md) and [Teams Toolkit](teams-toolkit-overview.md)  &#8212;  create projects that initially have one or more add-in commands. So, you will almost always be working with an add-in project that already has at least one add-in command. One exception is if you are updating an add-in which previously did not have an add-in command and from which the code and markup for any add-in commands has been removed. 

??????????????????

For simplicity, this article is written as if none of the relevant markup or code is already present, but you need to This article applies to both scenarios, so when we use the verb "ensure" for the most part, where "ensure" might mean "create", "add", "rename", or "verify" depending on what your starting point is. For example, "Ensure that there is a 'whatever' property in the 'something' object" means "Verify that there is a 'whatever' ..." when there already is a "whatever" property; but it means "Add a 'whatever' property ..." if there isn't one already.

??????????

## Two decisions

Decide which of two types of add-in commands you need: [Task pane or function](../design/add-in-commands.md#types-of-add-in-commands). Decide which kind of UI element you need: button or menu item. Then carry out the steps in the sections and subsections below that correspond to your decisions.

## Add a task pane command

The following subsections explain how to include a task pane command in an add-in.

### Configure the runtime

1. Open the unified manifest and find the "extensions.runtimes" array.
1. Ensure that there is a runtime object that has a "actions.type" property with the value "openPage". (This is the type of runtime that opens a task pane.) There usually is one and it is usually the first object in the "runtimes" array.
1. Ensure that the "requirements.capabilities" array contains an object that specifies a [Requirement Set](office-versions-and-requirement-sets.md) that supports add-in commands. For Outlook the minimum requirement set for add-in commands is [Mailbox 1.3](/javascript/api/requirement-sets/outlook/requirement-set-1.3/outlook-requirement-set-1.3). When support for the unified manifest is extended to other Office host applications, the minimum requirement set for add-in commands will be [AddinCommands 1.1](/javascript/api/requirement-sets/common/add-in-commands-requirement-sets).
1. Ensure that the "id" of the runtime object has a descriptive name such as "TaskPaneRuntime".
1. Ensure that the "code.page" property of the runtime object is set to the URL of the page that should open in the task pane, such as "https://localhost:3000/taskpane.html".
1. Ensure that the "actions.view" of the runtime object has a name that describes the content of the page that you set in the preceding step, such as "homepage" or "dashboard". 
1. Ensure that the "actions.id" of the runtime object has a descriptive name such as "ShowTaskPane" that indicates what happens when the user selects the add-in command button or menu item.
1. Set the other properties and subproperties of the runtime object as shown in the following completed example of a runtime object.

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
                    "pinnable": false,
                    "view": "homepage"
                }
            ]
        }        
    ]
    ```

### Configure the UI

1. Ensure that the extension object for which you configured a runtime has a "ribbons" array property as a peer to the "runtimes" array. (There is typically only one extension object in the "extensions" array.)
1. Ensure that the array has an object with array properties named "contexts" and "tabs" as shown in the following example.

    ```json
    "ribbons": [
        {
            "contexts": [
                // markup omitted
            ],
            "tabs": [
                // markup omitted
            ]
        }
    ]
    ```

1. Ensure that the "contexts" array has strings that specify the windows or panes in which the UI for the task pane command should appear. For example, "mailRead" means that it will appear in the reading pane or message window when an email message is open, but "mailCompose" means it will appear when a new message or a reply is being composed. The following is an example.

    ```json
    "contexts": [
        "mailRead"
    ],
    ```

1. Ensure that the "tabs" array has an object with "builtInTabId" string property that is set to the ID of ribbon tab in which you want your task pane command to appear and a "groups" array with at least one object in it. The following is an example.

    ```json
    "tabs": [
        {
            "builtInTabID": "TabDefault",
            "groups": [
                {
                    // markup omitted                
                }
            ]
        }
    ]
    ```

    > [!NOTE]
    > Only "TabDefault", which in Outlook is either the **Home**, **Message**, or **Meeting** tab, is currently supported for the "builtInTabID" property.

1. Ensure that the "groups" array has an object to define the custom control group that will hold your add-in command UI controls. The following is an example. Note the following about this markup:

    - The "id" must be unique across all groups in all ribbon objects in the manifest. Maximum length is 64 characters.
    - The "label" appears on the group in the ribbon. Maximum length is 64 characters.
    - One of the "icons" appears on the group only if the Office application window, and hence the ribbon, has been sized by the user too small for any of the controls in the group to appear. Office decides when to use one of these icons and which one to use based on the size of the window and the resolution of the device. You cannot control this. You must provide image files for 16, 32, and 80 pixels, while five other sizes are also supported (20, 24, 40, 48, and 64 pixels).
    
    > [!NOTE]
    > The name of the "icons.file" property may change during the preview of the unified manifest for Office Add-ins. If you get intellisense or manifest validation errors, try replacing "file" with "url".


    ```json
    "groups": [
        {
            "id": "msgReadGroup",
            "label": "Contoso Add-in",
            "icons": [
                {
                    "size": 16,
                    "file": "https://localhost:3000/assets/icon-16.png"
                },
                {
                    "size": 32,
                    "file": "https://localhost:3000/assets/icon-32.png"
                },
                {
                    "size": 80,
                    "file": "https://localhost:3000/assets/icon-80.png"
                }
            ],
            "controls": [
                {
                    // markup omitted 
                }
            ]
        }
    ]
    ```

1. Ensure that there is a control object in the "controls" array for each button or custom menu you want. The following is an example. Note the following about this markup:

    - The "id", "label", and "icons" properties have the same purpose and the same restrictions as the corresponding properties of a group object, except that they apply to a specific button or menu within the group.
    - The "type" property is set to "button" which means that the control will be a ribbon button. You can also configure a task pane command to be executed from a menu item. See [Menu and menu items](#menu-and-menu-items).
    - The "supertip.title" (maximum length: 64 characters) and "supertip.descrption" (maximum length: 128 characters) appear when the cursor is hovering over the button or menu.
    - The "actionID" must be an exact match for the "ribbons.actions.id" that you set in [Configure the runtime](#configure-the-runtime).

    ```json
    {
        "id": "msgReadOpenPaneButton",
        "type": "button",
        "label": "Show Task Pane",
        "icons": [
            {
                "size": 16,
                "file": "https://localhost:3000/assets/icon-16.png"
            },
            {
                "size": 32,
                "file": "https://localhost:3000/assets/icon-32.png"
            },
            {
                "size": 80,
                "file": "https://localhost:3000/assets/icon-80.png"
            }
        ],
        "supertip": {
            "title": "Show Contoso Task Pane",
            "description": "Opens the Contoso task pane."
        },
        "actionId": "ShowTaskPane"
    }
    ```

You have now completed adding a task pane command to your add-in.

## Add a function command

### Create the code for the function command

### Configure the runtime for the function command

### Configure the UI for the function command

#### Ribbon button for a function command


You have now completed adding a task pane command to your add-in.

#### Menu and menu items

To open a task pane from an item in a menu, take the following steps:

1. Set the "type" property of the control object to "menu". The example control object in the preceding section would be look like the following.

    ```json
    {
        "id": "msgContosoMenu",
        "type": "menu",

        // "label", "icons", "supertip" properties omitted.
    }
    ```

1. Add an "items" array to the control object to represent the menu choices. It must have at least one object, but a menu normally has more than one item. The following is an example.

    ```json
    {
        "id": "msgReadOpenPaneButton",
        "type": "menu",

        // "label", "icons", "supertip", and "actionId" properties omitted.

        "items" [
            {

            },
            {

            }
        ]
    }
    ```

