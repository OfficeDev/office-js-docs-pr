---
title: Custom keyboard shortcuts in Office Add-ins
description: 'Learn how to add custom keyboard shortcuts, also known as key combinations, to your Office Add-in.'
ms.date: 11/09/2020
localization_priority: Normal
---

# Add Custom keyboard shortcuts to your Office Add-ins (preview)

Keyboard shortcuts, also known as key combinations, enable your add-in's users to work more efficiently and they improve the add-in's accessibility for users with disabilities by providing an alternative to the mouse.

[!include[Keyboard shortcut prerequisites](../includes/keyboard-shortcuts-prerequisites.md)]

> [!NOTE]
> To start with a working version of an add-in with keyboard shortcuts already enabled, clone and run the sample [Excel Keyboard Shortcuts](https://github.com/OfficeDev/PnP-OfficeAddins/tree/master/Samples/excel-keyboard-shortcuts). When you are ready to add keyboard shortcuts to your own add-in, continue with this article.

There are three steps to add keyboard shortcuts to an add-in:

1. [Configure the add-in's manifest](#configure-the-manifest).
1. [Create or edit the shortcuts JSON file](#create-or-edit-the-shortcuts-json-file) to define actions and their keyboard shortcuts.
1. [Add one or more runtime calls](#create-a-mapping-of-actions-to-their-functions) of the [Office.actions.associate](/javascript/api/office/office.actions#associate) API to map a function to each action.

## Configure the manifest

There are two small changes to the manifest to make. One is to enable the add-in to use a shared runtime and the other is to point to a JSON-formatted file where you defined the keyboard shortcuts.

### Configure the add-in to use a shared runtime

Adding custom keyboard shortcuts requires your add-in to use the shared runtime. For more information, [Configure an add-in to use a shared runtime](../excel/configure-your-add-in-to-use-a-shared-runtime.md).

### Link the mapping file to the manifest

Immediately *below* (not inside) the `<VersionOverrides>` element in the manifest, add an [ExtendedOverrides](../reference/manifest/extendedoverrides.md) element. Set the `Url` attribute to the full URL of a JSON file in your project that you will create in a later step.

```xml
    ...
    </VersionOverrides>  
    <ExtendedOverrides Url="https://contoso.com/addin/shortcuts.json"></ExtendedOverrides>
</OfficeApp>
```

## Create or edit the shortcuts JSON file

Create a JSON file in your project. Be sure the path of the file matches the location you specified for the `Url` attribute of the [ExtendedOverrides](../reference/manifest/extendedoverrides.md) element. This file will describe your keyboard shortcuts, and the actions that they will invoke.

1. Inside the JSON file, add the following JSON:

    ```json
    {
        "actions": [
        ],
        "shortcuts": [
        ]
    }
    ```

1. The actions array will contain objects that define the actions to be invoked. Here is an example.

    ```json
        "actions": [
            {
                "id": "SHOWTASKPANE",
                "type": "ExecuteFunction",
                "name": "Show task pane for add-in"
            },
            {
                "id": "HIDETASKPANE",
                "type": "ExecuteFunction",
                "name": "Hide task pane for add-in"
            }
        ]
    ```

    For more information about the actions objects, see [Constructing the action objects](#constructing-the-action-objects). The complete schema for the shortcuts JSON is at [extended-manifest.schema.json](https://developer.microsoft.com/en-us/json-schemas/office-js/extended-manifest.schema.json).

1. The shortcuts array will contain objects that map key combinations onto actions. Here is an example.

    ```json
        "shortcuts": [
            {
                "action": "SHOWTASKPANE",
                "key": {
                    "default": "CTRL+SHIFT+UP"
                }
            },
            {
                "action": "HIDETASKPANE",
                "key": {
                    "default": "CTRL+SHIFT+DOWN"
                }
            }
        ]
    ```

    For more information about the shortcuts objects, including restrictions about property values, see [Constructing the shortcut objects](#constructing-the-shortcut-objects). The complete schema for the shortcuts JSON is at [extended-manifest.schema.json](https://developer.microsoft.com/en-us/json-schemas/office-js/extended-manifest.schema.json).

    > [!NOTE]
    > You can use "CONTROL" in place of "CTRL" throughout this article.

    In a later step, the actions will themselves be mapped to functions that you write. In this example, you will later map SHOWTASKPANE to a function that calls the `Office.addin.showAsTaskpane` method and HIDETASKPANE to a function that calls the `Office.addin.hide` method.

## Create a mapping of actions to their functions

1. In your project, open the JavaScript file loaded by your HTML page in the `<FunctionFile>` element.
1. In the JavaScript file, use the [Office.actions.associate](/javascript/api/office/office.actions#associate) API to map each action that you specified in the JSON file to a JavaScript function. Add the following JavaScript to the file. Note the following about the code:

    - The first parameter is one of the actions from the JSON file.
    - The second parameter is the function that runs when a user presses the key combination that is mapped to the action in the JSON file.

    ```javascript
    Office.actions.associate('-- action ID goes here--', function () {

    });
    ```

1. To continue the example, use `'SHOWTASKPANE'` as the first parameter.
1. For the body of the function, use the [Office.addin.showTaskpane](/javascript/api/office/office.addin.md#showastaskpane--) method to open the add-in's task pane. When you are done, the code should look like the following:

    ```javascript
    Office.actions.associate('SHOWTASKPANE', function () {
        return Office.addin.showAsTaskpane()
            .then(function () {
                return;
            })
            .catch(function (error) {
                return error.code;
            });
    });
    ```

1. Add a second call of `Office.actions.associate` function to map the `HIDETASKPANE` action to a function that calls [Office.addin.hide](/javascript/api/office/office.addin.md#hide--). The following is an example:

    ```javascript
    Office.actions.associate('HIDETASKPANE', function () {
        return Office.addin.hide()
            .then(function () {
                return;
            })
            .catch(function (error) {
                return error.code;
            });
    });
    ```

Following the previous steps lets your add-in toggle the visibility of the task pane by pressing **Ctrl+Shift+Up arrow key** and **Ctrl+Shift+Down arrow key**. This is the same behavior as shown in the [sample excel keyboard shortcuts add-in](https://github.com/OfficeDev/PnP-OfficeAddins/tree/master/Samples/excel-keyboard-shortcuts).

## Details and restrictions

### Constructing the action objects

Use the following guidelines when specifying the objects in the `action` array of the shortcuts.json:

- The property names `id` and `name` are mandatory.
- The `id` property is used to uniquely identify the action to invoke using a keyboard shortcut.
- The `name` property must be a user friendly string describing the action. It must be a combination of the characters A - Z, a - z, 0 - 9, and the punctuation marks "-", "_", and "+".
- The `type` property is optional. Currently only `ExecuteFunction` type is supported.

The following is an example:

```json
    "actions": [
        {
            "id": "SHOWTASKPANE",
            "type": "ExecuteFunction",
            "name": "Show task pane for add-in"
        },
        {
            "id": "HIDETASKPANE",
            "type": "ExecuteFunction",
            "name": "Hide task pane for add-in"
        }
    ]
```

The complete schema for the shortcuts JSON is at [extended-manifest.schema.json](https://developer.microsoft.com/en-us/json-schemas/office-js/extended-manifest.schema.json).

### Constructing the shortcut objects

Use the following guidelines when specifying the objects in the `shortcuts` array of the shortcuts.json:

- The property names `action`, `key`, and `default` are required.
- The value of the `action` property is a string and must match one of the `id` properties in the action object.
- The `default` property can be any combination of the characters A - Z, a -z, 0 - 9, and the punctuation marks "-", "_", and "+". (By convention, lower case letters are not used in these properties.)
- The `default` property must contain the name of at least one modifier key (ALT, CTRL, SHIFT) and only one other key.
- For Macs, we also support the COMMAND modifier key.
- For Macs, ALT is mapped to the OPTION key. For Windows, COMMAND is mapped to the CTRL key.
- When two characters are linked to the same physical key in a standard keyboard, then they are synonyms in the `default` property; for example, ALT+a and ALT+A are the same shortcut, so are CTRL+- and CTRL+\_ because "-" and "_" are the same physical key.
- The "+" character indicates that the keys on either side of it are pressed simultaneously.

The following is an example:

```json
    "shortcuts": [
        {
            "action": "SHOWTASKPANE",
            "key": {
                "default": "CTRL+SHIFT+UP"
            }
        },
        {
            "action": "HIDETASKPANE",
            "key": {
                "default": "CTRL+SHIFT+DOWN"
            }
        }
    ]
```

The complete schema for the shortcuts JSON is at [extended-manifest.schema.json](https://developer.microsoft.com/en-us/json-schemas/office-js/extended-manifest.schema.json).

> [!NOTE]
> Keytips, also known as sequential key shortcuts, such as the Excel shortcut to choose a fill color **Alt+H, H**, are not supported in Office add-ins.

### Using shortcuts when the focus is in the task pane

Currently, the keyboard shortcuts for an Office add-in can only be invoked when the user's focus is in the worksheet. When the user's focus is inside the Office UI (such as the task pane), none of the add-in's shortcuts are ignored. As a workaround, the add-in can define keyboard handlers that can invoke certain actions when the user's focus is inside of the add-in UI.

## Using key combinations that are already used by Office or another add-in

During the preview period, there is no system for determining what happens when a user presses a key combination that is registered by an add-in and also by Office or by another add-in. Behavior is undefined.

Currently, there is no workaround when two or more add-ins have registered the same keyboard shortcut, but you can minimize conflicts with Excel with these good practices:

- Use only keyboard shortcuts with the following pattern in your add-in: **Ctrl+Shift+Alt+*x***, where *x* is some other key.
- If you need more keyboard shortcuts, check the [list of Excel keyboard shortcuts](https://support.microsoft.com/office/keyboard-shortcuts-in-excel-1798d9d5-842a-42b8-9c99-9b7213f0040f), and avoid using any of them in your add-in.

## Browser shortcuts that cannot be overridden

You cannot use any of the following keyboard combinations. They are used by browsers and cannot be overridden. This list is a work in progress. If you discover other combinations that cannot be overridden, please let us know by using the feedback tool at the bottom of this page.

- Ctrl+N
- Ctrl+Shift+N
- Ctrl+T
- Ctrl+Shift+T
- Ctrl+W
- Ctrl+PgUp/PgDn

## Next Steps

- See the sample add-in [excel-keyboard-shortcuts](https://github.com/OfficeDev/PnP-OfficeAddins/tree/master/Samples/excel-keyboard-shortcuts).
