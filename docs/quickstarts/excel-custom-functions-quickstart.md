---
ms.date: 11/16/2022
description: Developing custom functions in Excel quick start guide.
title: Custom functions quick start
ms.prod: excel
ms.localizationpriority: high
---

# Get started developing Excel custom functions

With custom functions, developers can add new functions to Excel by defining them in JavaScript or TypeScript as part of an add-in. Excel users can access custom functions just as they would any native function in Excel, such as `SUM()`.

## Prerequisites

[!include[Set up requirements](../includes/set-up-dev-environment-beforehand.md)]
[!include[Yeoman generator prerequisites](../includes/quickstart-yo-prerequisites.md)]

- Office connected to a Microsoft 365 subscription (including Office on the web).

  > [!NOTE]
  > If you don't already have Office, you can [join the Microsoft 365 developer program](https://aka.ms/M365devprogram) to get a free, 90-day renewable Microsoft 365 subscription to use during development.

## Build your first custom functions project

To start, you'll use the Yeoman generator to create the custom functions project. This will set up your project with the correct folder structure, source files, and dependencies to begin coding your custom functions.

1. [!include[Yeoman generator create project guidance](../includes/yo-office-command-guidance.md)]

    - **Choose a project type:** `Excel Custom Functions using a Shared Runtime`
    - **Choose a script type:** `JavaScript`
    - **What do you want to name your add-in?** `My custom functions add-in`

    :::image type="content" source="../images/yo-office-excel-cf-quickstart.png" alt-text="Screenshot of the Yeoman Office Add-in generator command line interface prompts for custom functions projects.":::

    The Yeoman generator will create the project files and install supporting Node components.

1. The Yeoman generator will give you some instructions in your command line about what to do with the project, but ignore them and continue to follow our instructions. Navigate to the root folder of the project.

    ```command&nbsp;line
    cd "My custom functions add-in"
    ```

1. Build the project.

    ```command&nbsp;line
    npm run build
    ```

1. Start the local web server, which runs in Node.js. You can try out the custom function add-in in Excel. You may be prompted to open the add-in's task pane, although this is optional. You can still run your custom functions without opening your add-in's task pane.

# [Excel on Windows or Mac](#tab/excel-windows)

To test your add-in in Excel on Windows or Mac, run the following command. When you run this command, the local web server will start and Excel will open with your add-in loaded.

```command&nbsp;line
npm run start:desktop
```

[!INCLUDE [alert use https](../includes/alert-use-https.md)]

# [Excel on the web](#tab/excel-online)

To test your add-in in Excel on the web, run the following command. When you run this command, the local web server will start. Replace "{url}" with the URL of an Excel document on your OneDrive or a SharePoint library to which you have permissions.

[!INCLUDE [npm start:web command syntax](../includes/start-web-sideload-instructions.md)]

[!INCLUDE [alert use https](../includes/alert-use-https.md)]

---

## Try out a prebuilt custom function

The custom functions project that you created by using the Yeoman generator contains some prebuilt custom functions, defined within the **./src/functions/functions.js** file. The **./manifest.xml** file in the root directory of the project specifies that all custom functions belong to the `CONTOSO` namespace.

In your Excel workbook, try out the `ADD` custom function by completing the following steps.

1. Select a cell and type `=CONTOSO`. Notice that the autocomplete menu shows the list of all functions in the `CONTOSO` namespace.

1. Run the `CONTOSO.ADD` function, using numbers `10` and `200` as input parameters, by typing the value `=CONTOSO.ADD(10,200)` in the cell and pressing enter.

The `ADD` custom function computes the sum of the two numbers that you specify as input parameters. Typing `=CONTOSO.ADD(10,200)` should produce the result **210** in the cell after you press enter.

[!include[Manually register an add-in](../includes/excel-custom-functions-manually-register.md)]

## Next steps

Congratulations, you've successfully created a custom function in an Excel add-in! Next, build a more complex add-in with streaming data capability. The following link takes you through the next steps in the Excel add-in with custom functions tutorial.

> [!div class="nextstepaction"]
> [Excel custom functions add-in tutorial](../tutorials/excel-tutorial-create-custom-functions.md#create-a-custom-function-that-requests-data-from-the-web)

## Troubleshooting

You may encounter issues if you run the quick start multiple times. If the Office cache already has an instance of a function with the same name, your add-in gets an error when it sideloads. You can prevent this by [clearing the Office cache](../testing/clear-cache.md) before running `npm run start`.

:::image type="content" source="../images/custom-function-already-exists-error.png" alt-text="An error message in Excel titled 'Error installing functions'. It contains the text 'This add-in wasn't installed because a custom function with the same name already exists'.":::

## See also

- [Custom functions overview](../excel/custom-functions-overview.md)
- [Custom functions metadata](../excel/custom-functions-json.md)
- [Runtime for Excel custom functions](../excel/custom-functions-runtime.md)
- [Using Visual Studio Code to publish](../publish/publish-add-in-vs-code.md#using-visual-studio-code-to-publish)
