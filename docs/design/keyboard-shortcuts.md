---
title: Custom keyboard shortcuts in Office Add-ins
description: 'Learn how to add custom keyboard shortcuts, also known as key combinations, to your Office Add-in.'
ms.date: 11/06/2020
localization_priority: Normal
---

# Custom keyboard shortcuts in Office Add-ins (preview)

Keyboard shortcuts, also known as key combinations, enable your add-in's users to work more efficiently and they improve the add-in's accessibility for users with disabilities by providing an alternative to the mouse.

> [!IMPORTANT]
> Keyboard shortcuts are in preview. Please experiment with them in a development or testing environment but don't add them to a production add-in.

Keyboard shortcuts are only supported on Excel and only on these platforms and builds:

* Excel on Windows: Version 2009 (Build 13231.20262)
* Excel on Mac: 16.41.20091302
* Excel on the web

> [!NOTE]
> Keyboard shortcuts work only on platforms that support the following requirement sets. For more about requirement sets and how to work with them, see [Specify Office applications and API requirements](../develop/specify-office-hosts-and-api-requirements.md).
>
> - [SharedRuntime 1.1](../reference/requirement-sets/shared-runtime-requirement-sets.md)

There are four major steps to adding keyboard shortcuts to an add-in:

1. Configure the add-in's manifest.
1. Create or edit the extended overrides JSON file to map action names to keyboard combinations.
1. Add one or more runtime calls of the [Office.actions.associate](/javascript/api/office/office.actions#associate) API to map a function to each action name.

Each step is described in more detail below.

## Configure the manifest

> [!NOTE]
> If your add-in's manifest already has an [ExtendedOverrides](../reference/manifest/extendedoverrides.md) element (which would be just below the `<VersionOverrides>` element) then you are already using a feature that leverages extended overrides and you can skip this section. Continue with [Create or edit the extended overrides JSON file](#create-or-edit-the-extended-overrides-json-file).

### Configure the add-in to use a shared runtime

The keyboard shortcuts feature requires that the add-in use a shared runtime. To configure the add-in, see [Configure an add-in to use a shared runtime](../excel/configure-your-add-in-to-use-a-shared-runtime.md).

### Link the mapping file to the manifest

Immediately *below* (not inside) the `<VersionOverrides>` element in the manifest, add an [ExtendedOverrides](../reference/manifest/extendedoverrides.md) element. Set the `Url` attribute to the full URL of a JSON file in your project that you will create in a later step. Example: `https://localhost:3000/add-in/extended-overrides.json`. When you are ready for staging and then production, you will need to change this value. Example: `https://contoso.com/addin/extended-overrides.json`.

## Create or edit the extended overrides JSON file

If there isn't one already, create a JSON file at the path that you use in development for the `Url` attribute of the [ExtendedOverrides](../reference/manifest/extendedoverrides.md) element.

1. Be sure there is an outermost pair of braces (`{ }`)in the file.
1. Just inside this outermost object, add the following JSON markup. Note that the file must be proper JSON, not simple a JavaScript object, so the property names must be within quotation marks.

    ```javascript
    {
        "shortcuts": [
        ]
    }
    ```

1. The shortcuts array will contain objects that map key combinations onto action names. Here is an example. Note the following about this markup:

    - The property names you see here, `action`, `key`, and `default` are mandatory.
    - The value of the `action` property can be any string, and the `default` property can be any combination of the characters A - Z, a -z, 0 - 9, and the punctuation marks "-", "_", and "+". (By convention lower case letters are not used in these properties.)
    - The `default` property must contain the name of at least one modifier key (ALT, CTRL, SHIFT) and at least one other key. (An additional modifier key is possible on Mac.)
    - When two characters are linked to the same physical key in a standard keyboard, then they are synonyms in the `default` property; for example, ALT+a and ALT+A are the same shortcut, so are CTRL+- and CTRL+\_ because "-" and "_" are the same physical key.
    - The "+" character indicates that the keys on either side of it are pressed simultaneously.
    - In a later step, the action names will themselves be mapped to functions that you write. In this example, you will later map SHOWTASKPANE to a function that calls the `Office.addin.showAsTaskpane` method and HIDETASKPANE to a function that calls the `Office.addin.hide` method. 

    > [!NOTE] The complete schema for the shortcuts JSON is at [extended-manifest.schema.json](https://developer.microsoft.com/en-us/json-schemas/office-js/extended-manifest.schema.json).

    ```javascript
    {
        "shortcuts": [
            {
                "action": "SHOWTASKPANE",
                "key": {
                    "default": "CTRL+SHIFT+UP"
                }
            },
            {
                "action": "HIDETASKPANE",
                "key": {
                    "default": "CTRL+SHIFT+DOWN"
                }
            }
        ]
    }
    ```

    > [!NOTE]
    > Keytips, also known as sequential key shortcuts, such as the Excel shortcut to choose a fill color **Alt+H, H**, are not supported in Office add-ins.

1. Optionally, you can vary the key combination for Office on the web, Office on Windows, or Office on Mac with additional properties on the `"key"` property. The following is an example. The `"default"` combination is used on any platform that doesn't have it's own specified combination. Note that on Mac, you can use the modifier key COMMAND.

    ```javascript
    {
        "shortcuts": [
            {
                "action": "SHOWTASKPANE",
                "key": {
                    "default": "CTRL+SHIFT+UP",
                    "web": "CTRL+SHIFT+P",
                    "Win32": "CTRL+SHIFT+R",
                    "Mac": "COMMAND+SHIFT+S"
                }
            }
        ]
    }
    ```

## Create a mapping of functions to actions

The last major step is to map your custom functions onto the action names.

1. Be sure that the HTML file that the `<FunctionFile>` element in the manifest points to has a `<script>` tag that loads a custom JavaScript file.
1. In the JavaScript file, use calls of the [Office.actions.associate](/javascript/api/office/office.actions#associate) API to map a function to each action name that you used in the JSON file. To begin add the following to the file. Note the following about this code:

    - The first parameter is one of the action names from the JSON file.
    - The second parameter is the function that runs when a user presses the key combination that is mapped to the action name in the JSON file.

    ```javascript
    Office.actions.associate('-- acton name goes here--', function () {

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

1. Sideload the add-in and toggle the visibility of the task pane by pressing Ctrl+Shift+UpArrow and Ctrl+Shift+DownArrow.

An example of a simple add-in that uses several custom keyboard shortcuts is at [shortcut-sample-revision1](https://github.com/OfficeDev/testing-assets/tree/master/addins/shortcut-sample-revision1).

## Using key combinations that are already used by Office or another add-in

During the preview period, there is no system for determining what happens when a user presses a key combination that is registered by an add-in and also by Office or by another add-in. Behavior is undefined.

Currently, there is no workaround when two or more add-ins have registered the same keyboard shortcut, but you can minimize conflicts with Excel with these good practices:

- Use only keyboard shortcuts with the following patterns in your add-in.

    - **Alt-*n***, where *n* is a numeral from 1 to 9.
    - **Ctrl+Shift+Alt+*n***, where *n* is a numeral from 1 to 9.

- If you need more keyboard shortcuts, check the [list of Excel keyboard shortcuts](https://support.microsoft.com/office/keyboard-shortcuts-in-excel-1798d9d5-842a-42b8-9c99-9b7213f0040f), and avoid using any of them in your add-in.

### Browser shortcuts that cannot be overridden

You cannot use any of the following keyboard combinations. They are used by browsers and cannot be overridden. This list is a work in progress. If you discover other combinations that cannot be overridden, please let us know by using the feedback tool at the bottom of this page.

- Ctrl+N
- Ctrl+Shift+N
- Ctrl+T
- Ctrl+Shift+T
- Ctrl+W
- Ctrl+PgUp/PgDn