---
title: "Tutorial: Add keyboard shortcuts to an Office Add-in"
description: 'Learn how to add custom keyboard shortcuts to your add-in.'
ms.date: 11/06/2020
ms.prod: excel
localization_priority: Normal
---

# Tutorial: Add keyboard shortcuts to an Office Add-in (preview)

There are three steps to add keyboard shortcuts to an add-in:

> [!div class="checklist"]
> * Configure the add-in's manifest.
> * Create the shortcuts JSON file to define actions and their keyboard shortcuts.
> * Map functions to runtime calls with the `associate` method.

> [!NOTE]
> To start with a working version of the add-in with keyboard shortcuts already enabled. Clone and run the [Keyboard Shortcuts PnP](https://github.com/OfficeDev/PnP-OfficeAddins/tree/master/Samples/excel-keyboard-shortcuts) and follow along using the instructions below.

[!include[Keyboard shortcut prerequisites](../includes/keyboard-shortcuts-prerequisites.md)]

## Create an add-in project with a shared runtime

Adding custom keyboard shortcuts requires your add-in to use the shared runtime. Follow the instructions in [Configure an add-in to use a shared runtime](../excel/configure-your-add-in-to-use-a-shared-runtime.md) to begin.

Once you have completed the steps in that tutorial, return here to add keyboard shortcuts to that add-in.

## Link the mapping file to the manifest

1. Immediately *below* (not inside) the `<VersionOverrides>` element in the manifest, add an [ExtendedOverrides](../reference/manifest/extendedoverrides.md) element. Set the `Url` attribute to the full URL of a JSON file in your project that you will create in a later step.

```xml
    ...
    </VersionOverrides>
    <ExtendedOverrides Url="https://localhost:3000/shortcuts.json"></ExtendedOverrides>
</OfficeApp>
```

## Create the shortcuts JSON file

This file describes your keyboard shortcuts, and the actions they invoke.

1. In the base folder of your project, create a JSON file called **shortcuts.json**.
1. Inside the **shortcuts.json** file, the actions array will contain objects that define the actions to be invoked and the shortcuts array will contain objects that map key combinations onto actions. Example:

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
    }
    ```

## Create a mapping of actions to their functions

1. In your project, open **./src/taskpane.js**.
1. The [Office.actions.associate](/javascript/api/office/office.actions#associate) API maps actions that you specify in the JSON file to a JavaScript function. Add the following JavaScript to the file to map the `SHOWTASKPANE` action to a function that calls [Office.addin.showAsTaskpane](/javascript/api/office/office.addin.md#showastaskpane--):

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

1. Add a second call of `Office.actions.associate` function to map the `HIDETASKPANE` action to a function that calls [Office.addin.hide](/javascript/api/office/office.addin.md#hide--).

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

1. Save your changes and rebuild the project.

   ```command&nbsp;line
   npm run build
   ```

## Run your Office Add-in

Start the local web server, which runs in Node.js. Toggle the visibility of the task pane by pressing **Ctrl+Shift+Up arrow** and **Ctrl+Shift+Down arrow**.

# [Excel on Windows](#tab/excel-windows)

To test your add-in in Excel on Windows, run the following command. When you run this command, the local web server will start and Excel will open with your add-in loaded.

```command&nbsp;line
npm run start:desktop
```

# [Excel on the web](#tab/excel-online)

To test your add-in in Excel on the web, run the following command. When you run this command, the local web server will start.

```command&nbsp;line
npm run start:web
```

To try your add-in, open a new workbook in Excel on a browser. In this workbook, complete the following steps to sideload your add-in.

1. In Excel, choose the **Insert** tab and then choose **Add-ins**.

   ![Insert ribbon in Excel on the web with the My Add-ins icon highlighted](../images/excel-cf-online-register-add-in-1.png)

2. Choose **Manage My Add-ins** and select **Upload My Add-in**.

3. Choose **Browse...** and navigate to the root directory of the project that the Yeoman generator created.

4. Select the file **manifest.xml** and choose **Open**, then choose **Upload**.

---

## Next steps

- Learn more about keyboard shortcuts in [Custom keyboard shortcuts in Office Add-ins](../design/keyboard-shortcuts.md).
- See the sample add-in [excel-keyboard-shortcuts](https://github.com/OfficeDev/PnP-OfficeAddins/tree/master/Samples/excel-keyboard-shortcuts).
