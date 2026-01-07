---
title: Custom keyboard shortcuts in Office Add-ins
description: Learn how to add custom keyboard shortcuts, also known as key combinations, to your Office Add-in.
ms.date: 01/06/2026
ms.topic: how-to
ms.localizationpriority: medium
---

# Add custom keyboard shortcuts to your Office Add-ins

Keyboard shortcuts, also known as key combinations, make it possible for your add-in's users to work more efficiently. Keyboard shortcuts also improve the add-in's accessibility for users with disabilities by providing an alternative to the mouse.

There are three steps to add keyboard shortcuts to an add-in.

1. [Configure the add-in's manifest to use a shared runtime](#define-custom-keyboard-shortcuts).
1. [Define custom keyboard shortcuts](#define-custom-keyboard-shortcuts) and the actions they'll run.
1. [Map custom actions to their functions](#map-custom-actions-to-their-functions) using the [Office.actions.associate](/javascript/api/office/office.actions#office-office-actions-associate-member) API.

## Prerequisites

Keyboard shortcuts are currently only supported in the following platforms and build of **Excel** and **Word**.

- Office on the web
- Office on Windows
  - **Excel**: Version 2102 (Build 13801.20632) and later
  - **Word**: Version 2408 (Build 17928.20114) and later
- Office on Mac
  - **Excel**: Version 16.55 (21111400) and later
  - **Word**: Version 16.88 (24081116) and later

Additionally, keyboard shortcuts only work on platforms that support the following requirement sets. For information about requirement sets and how to work with them, see [Specify Office applications and API requirements](../develop/specify-office-hosts-and-api-requirements.md).

- [SharedRuntime 1.1](/javascript/api/requirement-sets/common/shared-runtime-requirement-sets)
- [KeyboardShortcuts 1.1](/javascript/api/requirement-sets/common/keyboard-shortcuts-requirement-sets) (required if the add-in provides its users with the option to customize keyboard shortcuts)

> [!TIP]
> To start with a working version of an add-in with keyboard shortcuts already configured, clone and run the [Use keyboard shortcuts for Office Add-in actions](https://github.com/OfficeDev/Office-Add-in-samples/tree/main/Samples/office-keyboard-shortcuts) sample. When you're ready to add keyboard shortcuts to your own add-in, continue with this article.

## Define custom keyboard shortcuts

The process to define custom keyboard shortcuts for your add-in varies depending on the type of manifest your add-in uses. Select the tab for the type of manifest you're using.

> [!TIP]
> To learn more about manifests for Office Add-ins, see [Office Add-ins manifest](../develop/add-in-manifests.md).

# [Unified app manifest for Microsoft 365](#tab/jsonmanifest)

> [!NOTE]
> Implementing keyboard shortcuts with the unified app manifest for Microsoft 365 is in public developer preview. This shouldn't be used in production add-ins. We invite you to try it out in test or development environments. For more information, see the [Microsoft 365 app manifest schema reference](/microsoft-365/extensibility/schema/?view=m365-app-prev&preserve-view=true).

If your add-in uses the unified app manifest for Microsoft 365, custom keyboard shortcuts and their actions are defined in the manifest.

1. In your add-in project, open the **manifest.json** file.
1. Add the following object to the [`"extensions.runtimes"`](/microsoft-365/extensibility/schema/extension-runtimes-array?view=m365-app-prev&preserve-view=true) array. Note the following about this markup.
    - The [`"actions"`](/microsoft-365/extensibility/schema/extension-runtimes-actions-item?view=m365-app-prev&preserve-view=true) objects specify the functions your add-in can run. In the following example, an add-in will be able to show and hide a task pane. You'll create these functions in a later section. Currently, custom keyboard shortcuts can only run actions that are of type `"executeFunction"`.
    - While the [`"actions.displayName"`](/microsoft-365/extensibility/schema/extension-runtimes-actions-item?view=m365-app-prev&preserve-view=true#displayname) property is optional, it's required if a custom keyboard shortcut will be created for the action. This property is used to describe the action of a keyboard shortcut. The description you provide appears in the dialog that's shown to a user when there's a shortcut conflict between multiple add-ins or with Microsoft 365. Office appends the name of the add-in in parentheses at the end of the description. For more information on how conflicts with keyboard shortcuts are handled, see [Avoid key combinations in use by other add-ins](#avoid-key-combinations-in-use-by-other-add-ins).

    ```json
    {
        "id": "TaskPaneRuntime",
        "type": "general",
        "code": {
            "page": "https://localhost:3000/taskpane.html"
        },
        "lifetime": "long",
        "actions": [
            {
                "id": "ShowTaskpane",
                "type": "executeFunction",
                "displayName": "Show task pane"
            },
            {
                "id": "HideTaskpane",
                "type": "executeFunction",
                "displayName": "Hide task pane"
            }
        ]
    }
    ```

1. Add the following to the [`"extensions"`](/microsoft-365/extensibility/schema/root#extensions) array. Note the following about the markup.
    - The SharedRuntime 1.1 requirement set is specified in the [`"requirements.capabilities"`](/microsoft-365/extensibility/schema/requirements-extension-element-capabilities?view=m365-app-prev&preserve-view=true) object to support custom keyboard shortcuts.
    - Each [`"shortcuts"`](/microsoft-365/extensibility/schema/extension-shortcut) object represents a single action that's invoked by a keyboard shortcut. It specifies the supported key combinations for various platforms, such as Office on the web, on Windows, and on Mac. For guidance on how to create custom key combinations, see [Guidelines for custom key combinations](#guidelines-for-custom-key-combinations).
    - A default key combination must be specified. It's used on all supported platforms if there isn't a specific combination configured for a particular platform.
    - The value of the `"actionId"` property must match the value specified in the `"id"` property of the applicable `"extensions.runtimes.actions"` object.

    ```json
    "keyboardShortcuts": [
        {
            "requirements": {
                "capabilities": [
                    {
                        "name": "SharedRuntime",
                        "minVersion": "1.1"
                    }
                ]
            },
            "shortcuts": [
                {
                    "key": {
                        "default": "Ctrl+Alt+Up",
                        "mac": "Command+Shift+Up",
                        "web": "Ctrl+Alt+1",
                        "windows": "Ctrl+Alt+Up"
                    },
                    "actionId": "ShowTaskpane"
                },
                {
                    "key": {
                        "default": "Ctrl+Alt+Down",
                        "mac": "Command+Shift+Down",
                        "web": "Ctrl+Alt+2",
                        "windows": "Ctrl+Alt+Down"
                    },
                    "actionId": "HideTaskpane"
                }
            ]
        }
    ]
    ```

> [!NOTE]
> If you've defined keyboard shortcuts for an add-in that uses the unified manifest and want to publish it to [Microsoft Marketplace](../publish/publish-office-add-ins-to-appsource.md), you must specify JSON resource files for the custom shortcuts and their localized strings (if applicable) in the manifest. These resource files are used for backward compatibility on platforms that don't directly support the unified manifest. To learn how to configure this in your manifest, see [Support backward compatibility for add-ins with a unified manifest in Microsoft Marketplace](#support-backward-compatibility-for-add-ins-with-a-unified-manifest-in-microsoft-marketplace).

# [Add-in only manifest](#tab/xmlmanifest)

### Configure the manifest to use a shared runtime

To customize keyboard shortcuts for your add-in, you must first configure the add-in manifest to use a [shared runtime](../testing/runtimes.md#shared-runtime). For guidance on how to configure your add-in to use a shared runtime, see [Configure an add-in to use a shared runtime](../develop/configure-your-add-in-to-use-a-shared-runtime.md).

### Create or edit the shortcuts JSON file

If your add-in uses an add-in only manifest, custom keyboard shortcuts are defined in a JSON file. This file describes your keyboard shortcuts and the actions that they'll invoke. The complete schema for the JSON file is at [extended-manifest.schema.json](https://developer.microsoft.com/json-schemas/office-js/extended-manifest.schema.json).

1. In your add-in project, create a JSON file.
1. Add the following markup to the file. Note the following about the code.
    - The `"actions"` array contains objects that define the actions to be invoked. The [`"actions.id"`](/microsoft-365/extensibility/schema/extension-runtimes-actions-item#id) and `"actions.name"` properties are required.
    - The `"actions.id"` property uniquely identifies the action to invoke using a keyboard shortcut.
    - The `"actions.name"` property must describe the action of a keyboard shortcut. The description you provide appears in the dialog that's shown to a user when there's a shortcut conflict between multiple add-ins or with Microsoft 365. Office appends the name of the add-in in parentheses at the end of the description. For more information on how conflicts with keyboard shortcuts are handled, see [Avoid key combinations in use by other add-ins](#avoid-key-combinations-in-use-by-other-add-ins).
    - The `"type"` property is optional. Currently, only the `"ExecuteFunction"` type is supported.
    - The specified actions will be mapped to functions that you create in a later step. In the example, you'll later map `"ShowTaskpane"` to a function that calls the `Office.addin.showAsTaskpane` method and `"HideTaskpane"` to a function that calls the `Office.addin.hide` method.
    - The `"shortcuts"` array contains objects that map key combinations to actions. The `"shortcuts.action"`, `"shortcuts.key"`, and `"shortcuts.key.default"` properties are required.
    - The value of the `"shortcuts.action"` property must match the `"actions.id"` property of the applicable action object.
    - It's possible to customize shortcuts to be platform-specific. In the example, the `"shortcuts"` object customizes shortcuts for each of the following platforms: `"windows"`, `"mac"`, and `"web"`. You must define a default shortcut key for each shortcut. This is used as a fallback key if a key combination isn't specified for a particular platform.

    > [!TIP]
    > For guidance on how to create custom key combinations, see [Guidelines for custom key combinations](#guidelines-for-custom-key-combinations).

    ```json
    {
        "actions": [
            {
                "id": "ShowTaskpane",
                "type": "ExecuteFunction",
                "name": "Show task pane"
            },
            {
                "id": "HideTaskpane",
                "type": "ExecuteFunction",
                "name": "Hide task pane"
            }
        ],
        "shortcuts": [
            {
                "action": "ShowTaskpane",
                "key": {
                    "default": "Ctrl+Alt+Up",
                    "mac": "Command+Shift+Up",
                    "web": "Ctrl+Alt+1",
                    "windows": "Ctrl+Alt+Up"
                }
            },
            {
                "action": "HideTaskpane",
                "key": {
                    "default": "Ctrl+Alt+Down",
                    "mac": "Command+Shift+Down",
                    "web": "Ctrl+Alt+2",
                    "windows": "Ctrl+Alt+Up"
                }
            }
        ]
    }
    ```

### Link the mapping file to the manifest

1. In your add-in project, open the **manifest.xml** file.
1. Immediately *below* (not inside) the `<VersionOverrides>` element in the manifest, add an [ExtendedOverrides](/javascript/api/manifest/extendedoverrides) element. Set the `Url` attribute to the full URL of the JSON file you created in a previous step.

```xml
    ...
    </VersionOverrides>
    <ExtendedOverrides Url="https://contoso.com/addin/shortcuts.json"></ExtendedOverrides>
</OfficeApp>
```

---

## Map custom actions to their functions

1. In your project, open the JavaScript file loaded by the HTML page specified in your manifest.

1. In the JavaScript file, use the [Office.actions.associate](/javascript/api/office/office.actions#office-office-actions-associate-member) API to map each action you specified in an earlier step to a JavaScript function. Add the following JavaScript to the file. Note the following about the code.
    - The first parameter is the name of an action that you mapped to a keyboard shortcut. The location of the name of the action depends on the type of manifest your add-in uses.
        - **Unified app manifest for Microsoft 365**: The value of the `"extensions.keyboardShortcuts.shortcuts.actionId"` property in the **manifest.json** file.
        - **Add-in only manifest**: The value of the `"actions.id"` property in the shortcuts JSON file.
    - The second parameter is the function that runs when a user presses the key combination that's mapped to an action.

    ```javascript
    Office.actions.associate("ShowTaskpane", () => {
        return Office.addin.showAsTaskpane()
            .then(() => {
                return;
            })
            .catch((error) => {
                return error.code;
            });
    });
    ```

    ```javascript
    Office.actions.associate("HideTaskpane", () => {
        return Office.addin.hide()
            .then(() => {
                return;
            })
            .catch((error) => {
                return error.code;
            });
    });
    ```

## Guidelines for custom key combinations

Use the following guidelines to create custom key combinations for your add-ins.

- A keyboard shortcut must include at least one modifier key (<kbd>Alt</kbd>/<kbd>Option</kbd>, <kbd>Ctrl</kbd>/<kbd>Cmd</kbd>, <kbd>Shift</kbd>) and only one other key. These keys must be joined with a `+` character.
- The <kbd>Cmd</kbd> modifier key is supported on the macOS platform.
- On macOS, the <kbd>Alt</kbd> key is mapped to the <kbd>Option</kbd> key. On Windows, the <kbd>Cmd</kbd> key is mapped to the <kbd>Ctrl</kbd> key.
- The <kbd>Shift</kbd> key can't be used as the only modifier key. It must be combined with either <kbd>Alt</kbd>/<kbd>Option</kbd> or <kbd>Ctrl</kbd>/<kbd>Cmd</kbd>.
- Key combinations can include characters "A-Z", "a-z", "0-9", and the punctuation marks "-", "_", and "+". By convention, lowercase letters aren't used in keyboard shortcuts.
- When two characters are linked to the same physical key on a standard keyboard, then they're synonyms in a custom keyboard shortcut. For example, <kbd>Alt</kbd>+<kbd>a</kbd> and <kbd>Alt</kbd>+<kbd>A</kbd> are the same shortcut, as well as <kbd>Ctrl</kbd>+<kbd>-</kbd> and <kbd>Ctrl</kbd>+<kbd>\_</kbd> ("-" and "_" are linked to the same physical key).

> [!NOTE]
> Custom keyboard shortcuts must be pressed simultaneously. KeyTips, also known as sequential key shortcuts (for example, <kbd>Alt</kbd>+<kbd>H</kbd>, <kbd>H</kbd>), aren't supported in Office Add-ins.

### Browser shortcuts that can't be overridden

When using custom keyboard shortcuts on the web, some keyboard shortcuts that are used by the browser can't be overridden by add-ins. The following list is a work in progress. If you discover other combinations that can't be overridden, please let us know by using the feedback tool at the bottom of this page.

- <kbd>Ctrl</kbd>+<kbd>N</kbd>
- <kbd>Ctrl</kbd>+<kbd>Shift</kbd>+<kbd>N</kbd>
- <kbd>Ctrl</kbd>+<kbd>T</kbd>
- <kbd>Ctrl</kbd>+<kbd>Shift</kbd>+<kbd>T</kbd>
- <kbd>Ctrl</kbd>+<kbd>W</kbd>
- <kbd>Ctrl</kbd>+<kbd>PgUp</kbd>/<kbd>PgDn</kbd>

### Avoid key combinations in use by other add-ins

There are many keyboard shortcuts that are already in use by Microsoft 365. Avoid registering keyboard shortcuts for your add-in that are already in use. However, there may be some instances where it's necessary to override existing keyboard shortcuts or handle conflicts between multiple add-ins that have registered the same keyboard shortcut.

In the case of a conflict, the user will see a dialog box the first time they attempt to use a conflicting keyboard shortcut. Note that the source of the text for the add-in option that's displayed in this dialog varies depending on the type of manifest your add-in uses.
    - **Unified app manifest for Microsoft 365**: The value of the `"extensions.runtimes.actions.displayName"` property in the **manifest.json** file.
    - **Add-in only manifest**: The value of the `"actions.name"` property in the shortcuts JSON file.

:::image type="content" source="../images/add-in-shortcut-conflict-modal.png" alt-text="A conflict modal with two different actions for a single shortcut.":::

The user can select which action the keyboard shortcut will take. After making the selection, the preference is saved for future uses of the same shortcut. The shortcut preferences are saved per user, per platform. If the user wishes to change their preferences, they can invoke the **Reset Office Add-ins shortcut preferences** command from the **Tell me** search box. Invoking the command clears all of the user's add-in shortcut preferences and the user will again be prompted with the conflict dialog box the next time they attempt to use a conflicting shortcut.

:::image type="content" source="../images/add-in-reset-shortcuts-action.png" alt-text="The Tell me search box in Excel showing the reset Office Add-in shortcut preferences action.":::

For the best user experience, we recommend that you minimize keyboard shortcut conflicts with these good practices.

- Use only keyboard shortcuts with the following pattern: <kbd>Ctrl</kbd>+<kbd>Shift</kbd>+<kbd>Alt</kbd>+*x*, where *x* is some other key.
- Avoid using established keyboard shortcuts in Excel and Word. For a list, see the following:
  - [Keyboard shortcuts in Excel](https://support.microsoft.com/office/1798d9d5-842a-42b8-9c99-9b7213f0040f)
  - [Keyboard shortcuts in Word](https://support.microsoft.com/office/95ef89dd-7142-4b50-afb2-f762f663ceb2)
- When the keyboard focus is inside the add-in UI, <kbd>Ctrl</kbd>+<kbd>Space</kbd> and <kbd>Ctrl</kbd>+<kbd>Shift</kbd>+<kbd>F10</kbd> won't work as these are essential accessibility shortcuts.
- On a Windows or Mac computer, if the **Reset Office Add-ins shortcut preferences** command isn't available on the search menu, the user can manually add the command to the ribbon by customizing the ribbon through the context menu.

## Localize the description of a keyboard shortcut

You may need to localize your custom keyboard shortcuts in the following scenarios.

- Your add-in supports another locale.
- Your add-in supports different alphabets, writing systems, or keyboard layouts.

Guidance on how to localize your keyboard shortcuts varies depending on the type of manifest your add-in uses.

# [Unified app manifest for Microsoft 365](#tab/jsonmanifest)

To learn how to localize your custom keyboard shortcuts with the unified app manifest for Microsoft 365, see [Localize strings in your app manifest](/microsoftteams/platform/concepts/build-and-test/apps-localization).

> [!NOTE]
> If you've defined keyboard shortcuts for an add-in that uses the unified manifest and want to publish it to [Microsoft Marketplace](../publish/publish-office-add-ins-to-appsource.md), you must specify JSON resource files for the custom shortcuts and their localized strings (if applicable) in the manifest. These resource files are used for backward compatibility on platforms that don't directly support the unified manifest. To learn how to configure this in your manifest, see [Support backward compatibility for add-ins with a unified manifest in Microsoft Marketplace](#support-backward-compatibility-for-add-ins-with-a-unified-manifest-in-microsoft-marketplace).

# [Add-in only manifest](#tab/xmlmanifest)

### Update the shortcuts JSON file

To define an alternative keyboard binding for another locale, you must specify tokens in your add-in's shortcuts JSON file. The tokens name reference strings in the shortcuts JSON file and the localization resource file, which you'll create in a later step. The following is an example that assigns a keyboard shortcut to a function (defined elsewhere) that displays the add-in's task pane. Note the following about this markup.

- The tokens must have the format **${resource.*name-of-resource*}**. The resource name must match the applicable string specified in the shortcuts and localization resource files.
- Default strings *must be defined in the shortcuts JSON file itself*. Default strings are used when the locale of the Microsoft 365 host application doesn't match the other *ll-cc* property in the localization resource file. Defining the default strings directly in the shortcuts file ensures that Microsoft 365 doesn't download the localization resource file when the locale of the Microsoft 365 application matches the default locale of the add-in (as specified in the manifest).

```json
{
    "actions": [
        {
            "id": "ShowTaskpane",
            "type": "ExecuteFunction",
            "name": "${resource.showTaskpane_action_name}"
        }
    ],
    "shortcuts": [
        {
            "action": "ShowTaskpane",
            "key": {
                "default": "${resource.showTaskpane_default_key}"
            }
        }
    ],
    "resources": { 
        "default": {
            "showTaskpane_action_name": {
                "value": "Show task pane",
                "comment": "Display name for the ShowTaskpane action."
            },
            "showTaskpane_default_key": { 
                "value": "Ctrl+Shift+A",
                "comment": "Default shortcut to show the task pane."
            }
        }
    }
}
```

### Create a localization resource file

While the default shortcuts and strings are defined in the shortcuts JSON file, the localization resource file configures alternative keyboard shortcuts for one additional locale. The localized strings defined in this file are used when the language of the Microsoft 365 host application matches the **ll-cc** property specified in the file.

Similar to the shortcuts file, the localization resource file is also JSON-formatted and includes strings for the alternative locale. A string is assigned to each token that was used in the shortcuts JSON file. The following is an example of alternative strings for `es-es`. Note that keyboard shortcuts may differ from the default when localizing for locales that have a different alphabet or writing system, and hence a different keyboard.

```json
{
    "showTaskpane_action_name": {
        "value": "(es-es) Mostrar panel de tareas",
        "comment": "Display name for the ShowTaskpane action."
    },
    "showTaskpane_default_key": {
        "value": "Ctrl+Shift+A",
        "comment": "(es-es) Shortcut to show the task pane."
    }
}
```

### Specify the localization resource file in the manifest

Use the `ResourcesUrl` attribute of the [ExtendedOverrides](/javascript/api/manifest/extendedoverrides) element to point Microsoft 365 to the localization resource file. The following is an example.

```xml
    ...
    </VersionOverrides>  
    <ExtendedOverrides Url="https://contoso.com/addin/shortcuts.json"
                       ResourcesUrl="https://contoso.com/addin/localization.json">
    </ExtendedOverrides>
</OfficeApp>
```

---

## Turn on shortcut customization for specific users

> [!NOTE]
> The APIs described in this section require the [KeyboardShortcuts 1.1](/javascript/api/requirement-sets/common/keyboard-shortcuts-requirement-sets) requirement set.

Users of your add-in can reassign the actions of the add-in to alternate keyboard combinations.

Use the [Office.actions.replaceShortcuts](/javascript/api/office/office.actions#office-office-actions-replaceshortcuts-member) method to assign a user's custom keyboard combinations to your add-ins actions. The method takes a parameter of type `{[actionId:string]: string|null}`, where the `actionId`s are a subset of the action IDs that must be defined in the add-in's extended manifest JSON. The values are the user's preferred key combinations. The value can also be `null`, which will remove any customization for that `actionId` and revert to the specified default keyboard combination.

If the user is logged into Microsoft 365, the custom combinations are saved in the user's roaming settings per platform. Customizing shortcuts aren't currently supported for anonymous users.

```javascript
const userCustomShortcuts = {
    ShowTaskpane: "Ctrl+Shift+1",
    HideTaskpane: "Ctrl+Shift+2"
};

Office.actions.replaceShortcuts(userCustomShortcuts)
    .then(() => {
        console.log("Successfully registered shortcut.");
    })
    .catch((error) => {
        if (error.code == "InvalidOperation") {
            console.log("ActionId doesn't exist or shortcut combination is invalid.");
        }
    });
```

To find out what shortcuts are already in use for the user, call the [Office.actions.getShortcuts](/javascript/api/office/office.actions#office-office-actions-getshortcuts-member) method. This method returns an object of type `[actionId:string]:string|null}`, where the values represent the current keyboard combination the user must use to invoke the specified action. The values can come from three different sources.

- If there was a conflict with the shortcut and the user has chosen to use a different action (either native or another add-in) for that keyboard combination, the value returned will be `null` since the shortcut has been overridden and there's no keyboard combination the user can currently use to invoke that add-in action.
- If the shortcut has been customized using the [Office.actions.replaceShortcuts](/javascript/api/office/office.actions#office-office-actions-replaceshortcuts-member) method, the value returned will be the customized keyboard combination.
- If the shortcut hasn't been overridden or customized, the value returned varies depending on the type of manifest the add-in uses.
  - **Unified app manifest for Microsoft 365**: The shortcut specified in the **manifest.json** file of the add-in.
  - **Add-in only manifest**: The shortcut specified in the shortcuts JSON file of the add-in.

The following is an example.

```javascript
Office.actions.getShortcuts()
    .then((userShortcuts) => {
       for (const action in userShortcuts) {
           let shortcut = userShortcuts[action];
           console.log(action + ": " + shortcut);
       }
    });
```

As described in [Avoid key combinations in use by other add-ins](#avoid-key-combinations-in-use-by-other-add-ins), it's a good practice to avoid conflicts in shortcuts. To discover if one or more key combinations are already in use, pass them as an array of strings to the [Office.actions.areShortcutsInUse](/javascript/api/office/office.actions#office-office-actions-areshortcutsinuse-member) method. The method returns a report containing key combinations that are already in use in the form of an array of objects of type `{shortcut: string, inUse: boolean}`. The `shortcut` property is a key combination, such as "Ctrl+Shift+1". If the combination is already registered to another action, the `inUse` property is set to `true`. For example, `[{shortcut: "Ctrl+Shift+1", inUse: true}, {shortcut: "Ctrl+Shift+2", inUse: false}]`. The following code snippet is an example.

```javascript
const shortcuts = ["Ctrl+Shift+1", "Ctrl+Shift+2"];
Office.actions.areShortcutsInUse(shortcuts)
    .then((inUseArray) => {
        const availableShortcuts = inUseArray.filter((shortcut) => {
            return !shortcut.inUse;
        });
        console.log(availableShortcuts);
        const usedShortcuts = inUseArray.filter((shortcut) => {
            return shortcut.inUse;
        });
        console.log(usedShortcuts);
    });
```

## Implement custom keyboard shortcuts across supported Microsoft 365 apps

You can implement a custom keyboard shortcut to be used across supported Microsoft 365 apps, such as Excel and Word. If the implementation to perform the same task is different on each app, you must use the `Office.actions.associate` method to call a different callback function for each app. The following code is an example.

```javascript
const host = Office.context.host;
if (host === Office.HostType.Excel) {
    Office.actions.associate("ChangeFormat", changeFormatExcel);
} else if (host === Office.HostType.Word) {
    Office.actions.associate("ChangeFormat", changeFormatWord);
}
...
```

## Support backward compatibility for add-ins with a unified manifest in Microsoft Marketplace

To publish an add-in that uses the unified manifest and implements custom keyboard shortcuts to Microsoft Marketplace, you must specify JSON resource files for the shortcuts and their localized strings (if applicable) in the manifest. This ensures your add-in's keyboard shortcuts and its localized resources work on platforms that don't directly support the unified manifest (for information on supported clients and platforms, see [Office Add-ins with the unified app manifest for Microsoft 365](../develop/unified-manifest-overview.md#client-and-platform-support)).

### Create JSON resource files

For guidance on how to create a shortcuts JSON file, see [Create or edit the shortcuts JSON file](#create-or-edit-the-shortcuts-json-file). If your custom shortcuts are supplemented with localized strings, you must define resource tokens in your shortcuts JSON file and create a localization resource file. For guidance, see [Update the shortcuts JSON file](#update-the-shortcuts-json-file) and [Create a localization resource file](#create-a-localization-resource-file).

### Specify the JSON resource files in the manifest

In your add-in's manifest, specify the JSON resource files in the [`"extensions.keyboardShortcuts.keyMappingFiles"`](/microsoft-365/extensibility/schema/extension-keyboard-shortcut?view=m365-app-prev&preserve-view=true#keymappingfiles) object.

- For the shortcuts JSON file, provide its full HTTPS URL in the [`"extensions.keyboardShortcuts.keyMappingFiles.shortcutsUrl"`](/microsoft-365/extensibility/schema/keyboard-shortcuts-mapping-files?view=m365-app-prev&preserve-view=true#shortcutsurl) property.
- For the localization resource file, provide its full HTTPS URL in the [`"extensions.keyboardShortcuts.keyMappingFiles.localizationResourceUrl"`](/microsoft-365/extensibility/schema/keyboard-shortcuts-mapping-files?view=m365-app-prev&preserve-view=true#localizationresourceurl) property.

The following is an example of how to specify shortcuts and localization resource files in the manifest.

```json
"keyboardShortcuts": [
    {
        ...
        "shortcuts": [
            ...
        ],
        "keyMappingFiles": {
            "shortcutsUrl": "https://contoso.com/addin/shortcuts.json",
            "localizationResourceUrl": "https://contoso.com/addin/localization.json"
        }
    }
]
```

## See also

- [Office Add-in sample: Use keyboard shortcuts for Office Add-in actions](https://github.com/OfficeDev/Office-Add-in-samples/tree/main/Samples/office-keyboard-shortcuts)
- [Shared runtime requirement sets](/javascript/api/requirement-sets/common/shared-runtime-requirement-sets)
- [Keyboard shortcuts requirement sets](/javascript/api/requirement-sets/common/keyboard-shortcuts-requirement-sets)
