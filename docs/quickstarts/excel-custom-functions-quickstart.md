---
ms.date: 01/08/2019
description: Learn best practices for developing custom functions in Excel.
title: Custom functions best practices (preview)
localization_priority: Normal
---

# Get started developing Excel Custom Functions

Custom functions enable developers to add new functions to Excel by defining those functions in JavaScript or Typescript as part of an add-in. Users within Excel can access custom functions just as they would any native function in Excel, such as `SUM()`.

## Platforms

Custom functions are currently in preview and subject to change. They are not supported for use in production environments. For preview purposes you can try custom functions on the following platforms.

- Excel Online
- Excel for Windows (64-bit version 1810 or later). At present, Excel for Windows 32-bit may not work for all scenarios.
- Excel for Mac (version 13.329 or later)

To use custom functions within Excel Online, login by using either your Office 365 subscription or a [Microsoft account](https://account.microsoft.com/account).

To use custom functions within Excel for Windows or Excel for Mac, you must have an Office 365 subscription, join the [Office Insider](https://products.office.com/office-insider) program (**Insider** level -- formerly called "Insider Fast"), and use a sufficiently recent build of Excel (as specified earlier).

If you are using a version of Office on your desktop which you downloaded from the Windows Store, you must be part of the [Windows Insider](https://insider.windows.com/) program at the **Insider** level (formerly called "Insider Fast"), running the April 2018 Update version or later to use custom functions. This is a new change as of January 2019.

## Subscribe to Office 365

If you don't already have an Office 365 subscription, you can get one by joining the [Office 365 Developer Program](https://developer.microsoft.com/en-us/office/dev-program).

## Set up your development environment

You'll need the following tools and related resources to begin creating custom functions.

- [Node.js](https://nodejs.org/en/) (version 8.0.0 or later)

- [Git Bash](https://git-scm.com/downloads) (or another Git client)

- The latest version of [Yeoman](https://yeoman.io/) and the [Yeoman generator for Office Add-ins](https://www.npmjs.com/package/generator-office). To install these tools globally, run the following command via the command prompt:

    ```
    npm install -g yo generator-office
    ```

    > [!NOTE]
    > Even if you have previously installed the Yeoman generator, we recommend updating your package to the latest version from npm.

- Excel for Windows (64-bit version 1810 or later) or Excel Online. See exact specifications in the [platform section](#plaforms).

- Join the [Office Insider program](https://products.office.com/office-insider) (**Insider** level -- formerly called "Insider Fast")

## Build your first custom functions project

To start, you'll use the Yeoman generator to create the custom functions project. This will set up your project with the correct folder structure, source files, and dependencies to begin coding your custom functions.

1. Run the following command and then answer the prompts as follows.

    ```
    yo office
    ```

    - Choose a project type: `Excel Custom Functions Add-in project (...)`

    - Choose a script type: `JavaScript`

    - What do you want to name your add-in? `stock-ticker`

    ![Yeoman generator for Office Add-ins prompts for custom functions](../images/12-10-fork-cf-pic.jpg)

    The Yeoman generator will create the project files and install supporting Node components. The project files come from the [Excel-Custom-Functions](https://github.com/OfficeDev/Excel-Custom-Functions) GitHub repository.

2. Go to the project folder.

    ```
    cd stock-ticker
    ```

3. Trust the self-signed certificate that is needed to run this project. For detailed instructions for either Windows or Mac, see [Adding Self Signed Certificates as Trusted Root Certificate](https://github.com/OfficeDev/generator-office/blob/master/src/docs/ssl.md).  

4. Build the project.

    ```
    npm run build
    ```

5. Start the local web server, which runs in Node.js.

    - If you'll be using Excel for Windows to test your custom functions, run the following command to start the local web server, launch Excel, and sideload the add-in:

        ```
         npm run start
        ```
        After running this command, your command prompt will show details about what has been done, another npm window will open showing the details of the build, and Excel will start with your add-in loaded. If you add-in does not load, check that you have completed step 3 properly.

    - If you'll be using Excel Online to test your custom functions, run the following command to start the local web server:

        ```
        npm run start-web
        ```

         After running this command, another window will open showing you the details of the build. To use your functions, open a new workbook in Office Online.

## Try out the prebuilt custom functions

The custom functions project that you created by using the Yeoman generator contains some prebuilt custom functions, defined within the **src/customfunctions.js** file. The **manifest.xml** file in the root directory of the project specifies that all custom functions belong to the `CONTOSO` namespace.

In your Excel workbook, try out the `ADD` custom function by completing the following steps in Excel:

1. Within a cell, type `=CONTOSO`. Notice that the autocomplete menu shows the list of all functions in the `CONTOSO` namespace.

2. Run the `CONTOSO.ADD` function, with numbers `10` and `200` as input parameters, by typing the value `=CONTOSO.ADD(10,200)` in the cell and pressing enter.

The `ADD` custom function computes the sum of the two numbers that you specify as input parameters. Typing `=CONTOSO.ADD(10,200)` should produce the result **210** in the cell after you press enter.

## Next steps

Congratulations, you've successfully created a custom function in an Excel add-in! Next, learn more about the capabilities of custom functions and build a more complex add-in by following along with the Excel custom functions add-in tutorial.

> [!div class="nextstepaction"]
> [Excel custom functions add-in tutorial](../tutorials/excel-tutorial-create-custom-functions.md)

## See also

* [Custom functions overview](../excel/custom-functions-overview.md)
* [Custom functions metadata](../excel/custom-functions-json.md)
* [Runtime for Excel custom functions](../excel/custom-functions-runtime.md)
* [Custom functions best practices](../excel/custom-functions-best-practices.md)
