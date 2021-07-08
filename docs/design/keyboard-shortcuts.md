---
title: Custom keyboard shortcuts in Office Add-ins
description: 'Learn how to add custom keyboard shortcuts, also known as key combinations, to your Office Add-in.'
ms.date: 06/02/2021
localization_priority: Normal
---

# Add custom keyboard shortcuts to your Office Add-ins

Keyboard shortcuts, also known as key combinations, enable your add-in's users to work more efficiently. Keyboard shortcuts also improve the add-in's accessibility for users with disabilities by providing an alternative to the mouse.

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

Adding custom keyboard shortcuts requires your add-in to use the shared runtime. For more information, [Configure an add-in to use a shared runtime](../develop/configure-your-add-in-to-use-a-shared-runtime.md).

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

1. Inside the JSON file, there are two arrays. The actions array will contain objects that define the actions to be invoked and the shortcuts array will contain objects that map key combinations onto actions. Here is an example:

    ```json
    {
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
        ],
        "shortcuts": [
            {
                "action": "SHOWTASKPANE",
                "key": {
                    "default": "Ctrl+Alt+Up"
                }
            },
            {
                "action": "HIDETASKPANE",
                "key": {
                    "default": "Ctrl+Alt+Down"
                }
            }
        ]
    }
    ```

    For more information about the JSON objects, see [Construct the action objects](#construct-the-action-objects) and [Construct the shortcut objects](#construct-the-shortcut-objects). The complete schema for the shortcuts JSON is at [extended-manifest.schema.json](https://developer.microsoft.com/json-schemas/office-js/extended-manifest.schema.json).

    > [!NOTE]
    > You can use "CONTROL" in place of "Ctrl" throughout this article.

    In a later step, the actions will themselves be mapped to functions that you write. In this example, you will later map SHOWTASKPANE to a function that calls the `Office.addin.showAsTaskpane` method and HIDETASKPANE to a function that calls the `Office.addin.hide` method.

## Create a mapping of actions to their functions

1. In your project, open the JavaScript file loaded by your HTML page in the `<FunctionFile>` element.
1. In the JavaScript file, use the [Office.actions.associate](/javascript/api/office/office.actions#associate) API to map each action that you specified in the JSON file to a JavaScript function. Add the following JavaScript to the file. Note the following about the code.

    - The first parameter is one of the actions from the JSON file.
    - The second parameter is the function that runs when a user presses the key combination that is mapped to the action in the JSON file.

    ```javascript
    Office.actions.associate('-- action ID goes here--', function () {

    });
    ```

1. To continue the example, use `'SHOWTASKPANE'` as the first parameter.
1. For the body of the function, use the [Office.addin.showTaskpane](/javascript/api/office/office.addin#showastaskpane--) method to open the add-in's task pane. When you are done, the code should look like the following:

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

1. Add a second call of `Office.actions.associate` function to map the `HIDETASKPANE` action to a function that calls [Office.addin.hide](/javascript/api/office/office.addin#hide--). The following is an example.

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

Following the previous steps lets your add-in toggle the visibility of the task pane by pressing **Ctrl+Alt+Up** and **Ctrl+Alt+Down**. The same behavior is shown in the [Excel keyboard shortcuts](https://github.com/OfficeDev/PnP-OfficeAddins/tree/master/Samples/excel-keyboard-shortcuts) sample in the Office Add-ins PnP repo in GitHub.

## Details and restrictions

### Construct the action objects

Use the following guidelines when specifying the objects in the `actions` array of the shortcuts.json.

- The property names `id` and `name` are mandatory.
- The `id` property is used to uniquely identify the action to invoke using a keyboard shortcut.
- The `name` property must be a user friendly string describing the action. It must be a combination of the characters A - Z, a - z, 0 - 9, and the punctuation marks "-", "_", and "+".
- The `type` property is optional. Currently only `ExecuteFunction` type is supported.

The following is an example.

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

The complete schema for the shortcuts JSON is at [extended-manifest.schema.json](https://developer.microsoft.com/json-schemas/office-js/extended-manifest.schema.json).

### Construct the shortcut objects

Use the following guidelines when specifying the objects in the `shortcuts` array of the shortcuts.json.

- The property names `action`, `key`, and `default` are required.
- The value of the `action` property is a string and must match one of the `id` properties in the action object.
- The `default` property can be any combination of the characters A - Z, a -z, 0 - 9, and the punctuation marks "-", "_", and "+". (By convention, lower case letters are not used in these properties.)
- The `default` property must contain the name of at least one modifier key (Alt, Ctrl, Shift) and only one other key. 
- Shift cannot be used as the only modifier key. Combine Shift with either Alt or Ctrl.
- For Macs, we also support the Command modifier key.
- For Macs, Alt is mapped to the Option key. For Windows, Command is mapped to the Ctrl key.
- When two characters are linked to the same physical key in a standard keyboard, then they are synonyms in the `default` property; for example, Alt+a and Alt+A are the same shortcut, so are Ctrl+- and Ctrl+\_ because "-" and "_" are the same physical key.
- The "+" character indicates that the keys on either side of it are pressed simultaneously.

The following is an example.

```json
    "shortcuts": [
        {
            "action": "SHOWTASKPANE",
            "key": {
                "default": "Ctrl+Alt+Up"
            }
        },
        {
            "action": "HIDETASKPANE",
            "key": {
                "default": "Ctrl+Alt+Down"
            }
        }
    ]
```

The complete schema for the shortcuts JSON is at [extended-manifest.schema.json](https://developer.microsoft.com/json-schemas/office-js/extended-manifest.schema.json).

> [!NOTE]
> KeyTips, also known as sequential key shortcuts, such as the Excel shortcut to choose a fill color **Alt+H, H**, are not supported in Office Add-ins.

## Avoid key combinations in use by other add-ins

There are many keyboard shortcuts that are already in use by Office. Avoid registering keyboard shortcuts for your add-in that are already in use, however there may be some instances where it is necessary to override existing keyboard shortcuts or handle conflicts between multiple add-ins that have registered the same keyboard shortcut.

In the case of a conflict, the user will see a dialog box the first time they attempt to use a conflicting keyboard shortcut, note that the action name that is displayed in this dialog is the `name` property in the action object in `shortcuts.json` file.

![Illustration showing a conflict modal with two different actions for a single shortcut.](../images/add-in-shortcut-conflict-modal.png)

The user can select which action the keyboard shortcut will take. After making the selection, the preference is saved for future uses of the same shortcut. The shortcut preferences are saved per user, per platform. If the user wishes to change their preferences, they can invoke the **Reset Office Add-ins shortcut preferences** command from the **Tell me** search box. Invoking the command clears all of the user's add-in shortcut preferences and the user will again be prompted with the conflict dialog box the next time they attempt to use a conflicting shortcut:

![The Tell me search box in Excel showing the reset Office Add-in shortcut preferences action.](../images/add-in-reset-shortcuts-action.png)

For the best user experience, we recommend that you minimize conflicts with Excel with these good practices:

- Use only keyboard shortcuts with the following pattern: **Ctrl+Shift+Alt+*x***, where *x* is some other key.
- If you need more keyboard shortcuts, check the [list of Excel keyboard shortcuts](https://support.microsoft.com/office/keyboard-shortcuts-in-excel-1798d9d5-842a-42b8-9c99-9b7213f0040f), and avoid using any of them in your add-in.
- When the keyboard focus is inside the add-in UI, **Ctrl+Spacebar** and **Ctrl+Shift+F10** will not work as these are essential accessibility shortcuts.
- On a Windows or Mac computer, if the "Reset Office Add-ins shortcut preferences" command is not available on the search menu, the user can manually add the command to the ribbon by customizing the ribbon through the context menu.

## Customize the keyboard shortcuts per platform

It's possible to customize shortcuts to be platform-specific. The following is an example of the `shortcuts` object that customizes the shortcuts for each of the following platforms: `windows`, `mac`, `web`. Note that you must still have a `default` shortcut key for each shortcut.

In the following example, the `default` key is the fallback key for any platform that is not specified. The only platform not specified is Windows, so the `default` key will only apply to Windows.

```json
    "shortcuts": [
        {
            "action": "SHOWTASKPANE",
            "key": {
                "default": "Ctrl+Alt+Up",
                "mac": "Command+Shift+Up",
                "web": "Ctrl+Alt+1",
            }
        },
        {
            "action": "HIDETASKPANE",
            "key": {
                "default": "Ctrl+Alt+Down",
                "mac": "Command+Shift+Down",
                "web": "Ctrl+Alt+2"
            }
        }
    ]
```

## Localize the keyboard shortcuts JSON

If your add-in supports multiple locales, you'll need to localize the `name` property of the action objects. Also, if any of the locales that the add-in supports have alphabets or different writing systems, and hence different keyboards, you may need to localize the shortcuts also. For information about how to localize the keyboard shortcuts JSON, see [Localize extended overrides](../develop/localization.md#localize-extended-overrides).

## Browser shortcuts that cannot be overridden

When using custom keyboard shortcuts on the web, some keyboard shortcuts that are used by the browser cannot be overridden by add-ins. This list is a work in progress. If you discover other combinations that cannot be overridden, please let us know by using the feedback tool at the bottom of this page.

- Ctrl+N
- Ctrl+Shift+N
- Ctrl+T
- Ctrl+Shift+T
- Ctrl+W
- Ctrl+PgUp/PgDn

## Next Steps

- See the [Excel keyboard shortcuts](https://github.com/OfficeDev/PnP-OfficeAddins/tree/master/Samples/excel-keyboard-shortcuts) sample add-in.
- Get an overview of working with extended overrides in [Work with extended overrides of the manifest](../develop/extended-overrides.md).
