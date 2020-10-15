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

1. The shortcuts array will contain objects that map key combinations onto action names. Here is an example. The property names you see here, `action`, `key`, and `default` are mandatory. The values of the `action` and `default` properties are all capitalized by convention. In a later step, the action names will themselves be mapped to functions that you will write. In this case, SHOWTASKPANE will be mapped to a function that calls the `Office.addin.showAsTaskpane` method and HIDETASKPANE will be mapped to a function that calls the `Office.addin.hide` method. The complete schema for the shortcuts JSON is at [extended-manifest.schema.json](https://developer.microsoft.com/en-us/json-schemas/office-js/extended-manifest.schema.json).

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

1. Optionally, you can vary the key combination for Office on the web, Office on Windows, or Office on Mac with additional properties on the `"key"` property. The following is an example. The `"default"` combination is used on any platform that doesn't have it's own specified combination.

    ```javascript
    {
        "shortcuts": [
            {
                "action": "SHOWTASKPANE",
                "key": {
                    "default": "CTRL+SHIFT+UP"
                    "web": "CTRL+SHIFT+P"
                    "Win32": "CTRL+SHIFT+R"
                    "Mac": "CTRL+SHIFT+S"
                }
            }
        ]
    }
    `


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

You may want to override a key combination that is used by Office, or you may inadvertently use a combination that is used by Office or by another add-in. In either case, the first time a user presses a key combination that is registered by your add-in and by Office or by another add-in, then the user will be prompted to choose which action should be taken by the combination. The prompt will include a brief description of each option. For example, if your add-in uses a key combination to turn the selected cell value red, but Office uses the same combination to bold the selected cell value, and another add-in uses it to overwrite the selected cell with imported data; then Office will prompt the user with three options with labels like "Format red", "Bold", and "Overwrite".

### Browser shortcuts that cannot be overridden

You cannot use any of the following keyboard combinations. They are used by browsers and cannot be overridden. This list is a work in progress. If you discover other combinations that cannot be overridden, please let us know by using the feedback tool at the bottom of this page.

- Ctrl+N
- Ctrl+Shift+N
- Ctrl+T
- Ctrl+Shift+T
- Ctrl+W
- Ctrl+PgUp/PgDn