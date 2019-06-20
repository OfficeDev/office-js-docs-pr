---
ms.date: 06/20/2019
description: Developing custom functions in Excel quick start guide.
title: Custom functions quick start
ms.prod: excel
localization_priority: Normal
---

# Get started developing Excel custom functions

With custom functions, developers can now add new functions to Excel by defining them in JavaScript or Typescript as part of an add-in. Excel users can access custom functions just as they would any native function in Excel, such as `SUM()`.

## Prerequisites

[!include[Yeoman generator prerequisites](../includes/quickstart-yo-prerequisites.md)]

* Excel on Windows (version 1810 or later) or Excel Online

## Build your first custom functions project

To start, you'll use the Yeoman generator to create the custom functions project. This will set up your project with the correct folder structure, source files, and dependencies to begin coding your custom functions.

1. In a folder of your choice, run the following command and then answer the prompts as follows.

    ```command&nbsp;line
    yo office
    ```

    - **Choose a project type:** `Excel Custom Functions Add-in project`
    - **Choose a script type:** `JavaScript`
    - **What do you want to name your add-in?** `starcount`

    ![Yeoman generator for Office Add-ins prompts for custom functions](../images/starcountPrompt.png)

    The Yeoman generator will create the project files and install supporting Node components.

2. The Yeoman generator will give you some instructions in your command line about what to do with the project, but ignore them and continue to follow our instructions. Navigate to the root folder of the project.

    ```command&nbsp;line
    cd starcount
    ```

3. Build the project. 

    ```command&nbsp;line
    npm run build
    ```

    > [!NOTE]
    > Office Add-ins should use HTTPS, not HTTP, even when you are developing. If you are prompted to install a certificate after you run `npm run build`, accept the prompt to install the certificate that the Yeoman generator provides.

4. Start the local web server, which runs in Node.js. You can try out the custom function add-in in Excel on Windows or Excel Online. You may be prompted to open the add-in's task pane, although this is optional. You can still run your custom functions without opening your add-in's task pane.

# [Excel on Windows](#tab/excel-windows)

To test your add-in in Excel on Windows, run the following command. When you run this command, the local web server will start and Excel will open with your add-in loaded.

```command&nbsp;line
npm run start:desktop
```

# [Excel Online](#tab/excel-online)

To test your add-in in Excel Online, run the following command. When you run this command, the local web server will start.

```command&nbsp;line
npm run start:web
```

To use your custom functions add-in, open a new workbook in Excel Online. In this workbook, complete the following steps to sideload your add-in.

1. In Excel Online, choose the **Insert** tab and then choose **Add-ins**.

   ![Insert ribbon in Excel Online with the My Add-ins icon highlighted](../images/excel-cf-online-register-add-in-1.png)
   
2. Choose **Manage My Add-ins** and select **Upload My Add-in**.

3. Choose **Browse...** and navigate to the root directory of the project that the Yeoman generator created.

4. Select the file **manifest.xml** and choose **Open**, then choose **Upload**.

---

## Try out a prebuilt custom function

The custom functions project that you created by using the Yeoman generator contains some prebuilt custom functions, defined within the **./src/functions/functions.js** file. The **./manifest.xml** file in the root directory of the project specifies that all custom functions belong to the `CONTOSO` namespace.

In your Excel workbook, try out the `ADD` custom function by completing the following steps:

1. Select a cell and type `=CONTOSO`. Notice that the autocomplete menu shows the list of all functions in the `CONTOSO` namespace.

2. Run the `CONTOSO.ADD` function, using numbers `10` and `200` as input parameters, by typing the value `=CONTOSO.ADD(10,200)` in the cell and pressing enter.

The `ADD` custom function computes the sum of the two numbers that you specify as input parameters. Typing `=CONTOSO.ADD(10,200)` should produce the result **210** in the cell after you press enter.

## Next steps

Congratulations, you've successfully created a custom function in an Excel add-in! Next, build a more complex add-in with streaming data capability. The following link takes you through the next steps in the Excel add-in with custom functions tutorial.

> [!div class="nextstepaction"]
> [Excel custom functions add-in tutorial](../tutorials/excel-tutorial-create-custom-functions.md#create-a-custom-function-that-requests-data-from-the-web
)

## See also

* [Custom functions overview](../excel/custom-functions-overview.md)
* [Custom functions metadata](../excel/custom-functions-json.md)
* [Runtime for Excel custom functions](../excel/custom-functions-runtime.md)
* [Custom functions best practices](../excel/custom-functions-best-practices.md)
