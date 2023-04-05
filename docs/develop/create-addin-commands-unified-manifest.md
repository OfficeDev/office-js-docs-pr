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

1. Ensure that the "contexts" array has strings that specify the windows or panes in which the UI for the task pane add-in command should appear. For example, "mailRead" means that it will appear in the reading pane or message window when an email message is open, but "mailCompose" means it will appear when a new message or a reply is being composed. The following is an example.

    ```json
    "contexts": [
        "mailRead"
    ],
    ```

1. Ensure that the "tabs" array has an object with "builtInTabId" string property that is set to the ID of ribbon tab in which you want your task pane add-in command to appear and a "groups" array with at least one object in it. The following is an example.

    ```json
    "tabs": [
        "builtInTabID": "defaultTab",
        "groups: [
            {

            },
        ]
    ],
    ```

    > [!NOTE]
    > Only "tabDefault", which in Outlook is either the **Home**, **Message**, or **Meeting** tab, is currently supported for the "builtInTabID" property.

1. 

#### Ribbon button

#### Menu item

## Add a function command

### Configure the runtime

### Configure the UI

#### Ribbon button

#### Menu item

### Create the code for the function
