---
title: 'Tutorial: Create custom functions in Excel'
description: Build an Excel add-in with JavaScript custom functions that calculate values, retrieve web data, and stream real-time updates.
ms.date: 07/28/2026
ms.service: excel
ms.topic: tutorial
# Customer intent: As an add-in developer, I want to create custom functions in Excel to increase user productivity.
ms.localizationpriority: high
ai-usage: ai-assisted
---

# Tutorial: Create custom functions in Excel

Build an Excel add-in that provides JavaScript custom functions alongside built-in functions such as `SUM`. You'll create functions that perform a calculation, retrieve data from the web, and stream real-time updates into a worksheet.

In this tutorial, you:

> [!div class="checklist"]
>
> - Create a custom function add-in by using the [Yeoman generator for Office Add-ins](../develop/yeoman-generator-overview.md).
> - Use a prebuilt custom function to perform a simple calculation.
> - Create a custom function that gets data from the web.
> - Create a custom function that streams real-time data from the web.

## Prerequisites

[!INCLUDE [Yeoman generator prerequisites](../includes/quickstart-yo-prerequisites.md)]

## Create a custom functions project

Create the code project for your custom function add-in. The [Yeoman generator for Office Add-ins](../develop/yeoman-generator-overview.md) sets up the project with prebuilt custom functions to try. If you already generated a project in the custom functions quickstart, use that project and continue at [Create a custom function that requests data from the web](#create-a-custom-function-that-requests-data-from-the-web).

> [!NOTE]
> If you recreate the Yo Office project, you might get an error because the Office cache already has an instance of a function with the same name. To prevent this error, [clear the Office cache](../testing/clear-cache.md) before running `npm run start`.

1. [!INCLUDE [Yeoman generator create project guidance](../includes/yo-office-command-guidance.md)]

    - **Choose a project type:** `Excel Custom Functions using a Shared Runtime`
    - **Choose a script type:** `JavaScript`
    - **What do you want to name your add-in?** `My custom functions add-in`

    :::image type="content" source="../images/yo-office-excel-cf-quickstart.png" alt-text="The Yeoman Office Add-in generator command line interface prompts for custom functions projects.":::

    The Yeoman generator creates the project files and installs supporting Node components.

1. Go to the root folder of the project.

    ```command&nbsp;line
    cd "My custom functions add-in"
    ```

1. Build the project.

    ```command&nbsp;line
    npm run build
    ```

    > [!NOTE]
    > Office Add-ins should use HTTPS, not HTTP, even when you're developing. If you're prompted to install a certificate after you run `npm run build`, accept the prompt to install the certificate that the Yeoman generator provides.

1. Start the local web server, which runs in Node.js. You can try out the custom function add-in in Excel.

# [Excel on Windows or Mac](#tab/excel-windows)

The command to test your add-in in Excel on Windows or Mac depends on when you created the project. If the `"scripts"` section of the project's package.json file has a `start:desktop` script, run `npm run start:desktop`. Otherwise, run `npm run start`. The local web server starts and Excel opens with your add-in loaded.

[!INCLUDE [alert use https](../includes/alert-use-https.md)]

# [Excel on the web](#tab/excel-online)

To test your add-in in Excel on the web, run the following command. When you run this command, the local web server starts. Replace `{url}` with the URL of an Excel document on your OneDrive or a SharePoint library to which you have permissions.

[!INCLUDE [npm start on web command syntax](../includes/start-web-sideload-instructions.md)]

[!INCLUDE [alert use https](../includes/alert-use-https.md)]

---

## Try a prebuilt custom function

The project contains prebuilt custom functions in **./src/functions/functions.js**. The **./manifest.xml** file assigns them to the `CONTOSO` namespace, which you use to access the functions in Excel.

Next, try the `ADD` custom function by completing the following steps.

1. In Excel, go to any cell and enter `=CONTOSO`. Notice that the autocomplete menu shows the list of all functions in the `CONTOSO` namespace.

1. Enter `=CONTOSO.ADD(10,200)` in the cell and then select <kbd>Enter</kbd>.

The `ADD` custom function returns `210`.

[!INCLUDE [Manually register an add-in](../includes/excel-custom-functions-manually-register.md)]

> [!NOTE]
> See the [Troubleshooting](#troubleshooting) section of this article if you encounter errors when sideloading the add-in.

## Create a custom function that requests data from the web

Integrating data from the web is a great way to extend Excel through custom functions. Create a `getStarCount` custom function that retrieves the number of stars for a GitHub repository.

1. In the **My custom functions add-in** project, open **./src/functions/functions.js** in your code editor.

1. In **functions.js**, add the following code.

    ```js
    /**
     * Gets the star count for a GitHub repository.
     * @customfunction
     * @param {string} userName GitHub user or organization name.
     * @param {string} repoName GitHub repository name.
     * @returns {number} Number of stars given to the GitHub repository.
     */
    async function getStarCount(userName, repoName) {
      try {
        const url = `https://api.github.com/repos/${userName}/${repoName}`;
        const response = await fetch(url);

        if (!response.ok) {
          throw new Error(response.statusText);
        }

        const jsonResponse = await response.json();
        return jsonResponse.stargazers_count;
      } catch (error) {
        throw new CustomFunctions.Error(CustomFunctions.ErrorCode.notAvailable, String(error));
      }
    }
    ```

1. Run the following command to rebuild the project.

    ```command&nbsp;line
    npm run build
    ```

1. Complete the following steps (for Excel on the web, Windows, or Mac) to re-register the add-in in Excel. You must complete these steps before the new function is available.

### [Excel on Windows or Mac](#tab/excel-windows)

1. Close Excel and then reopen Excel.

1. In the Excel ribbon, select **Home** > **Add-ins**.

1. Under the **Developer Add-ins** section, select **My custom functions add-in** to register it.

    :::image type="content" source="../images/excel-cf-select-add-in.png" alt-text="The My Add-ins dialog that shows active add-ins, with the My custom function add-in button highlighted.":::

1. In cell **B1**, enter `=CONTOSO.GETSTARCOUNT("OfficeDev", "Office-Add-in-Samples")`, and then select <kbd>Enter</kbd>. The cell displays the current number of stars for the [Office-Add-in-Samples repository](https://github.com/OfficeDev/Office-Add-in-Samples).

# [Excel on the web](#tab/excel-online)

1. Select **Home** > **Add-ins**, and then select **More Settings**.

1. On the **Office Add-ins** dialog, select **Upload My Add-in**.

1. Choose **Browse...** and go to the root directory of the project that the Yeoman generator created.

1. Select the **manifest.xml** file and choose **Open**, and then choose **Upload**.

1. In cell **B1**, enter `=CONTOSO.GETSTARCOUNT("OfficeDev", "Excel-Custom-Functions")`, and then select <kbd>Enter</kbd>. The cell shows the current number of stars for the [Excel-Custom-Functions repository](https://github.com/OfficeDev/Excel-Custom-Functions).

---

> [!NOTE]
> See the [Troubleshooting](#troubleshooting) section of this article if you encounter errors when sideloading the add-in.

## Create a streaming asynchronous custom function

The `getStarCount` function returns the number of stars at a specific moment. A streaming function, in contrast, can update a cell repeatedly. It includes an `invocation` parameter that represents the cell that called the function.

The following sample contains two functions. `currentTime` returns the current time as a string. The streaming `clock` function calls `invocation.setResult` to update the cell every second and uses `invocation.onCanceled` to stop the timer when Excel cancels the function.

The **My custom functions add-in** project already contains these functions in **./src/functions/functions.js**.

```js
/**
 * Returns the current time
 * @returns {string} String with the current time formatted for the current locale.
 */
function currentTime() {
  return new Date().toLocaleTimeString();
}

/**
 * Displays the current time once a second.
 * @customfunction
 * @param {CustomFunctions.StreamingInvocation<string>} invocation Custom function invocation
 */
function clock(invocation) {
  const timer = setInterval(() => {
    const time = currentTime();
    invocation.setResult(time);
  }, 1000);

  invocation.onCanceled = () => {
    clearInterval(timer);
  };
}
```

To try the streaming function, enter `=CONTOSO.CLOCK()` in cell **C1**, and then select <kbd>Enter</kbd>. The cell displays the current time and updates every second. You can use the same timer pattern with functions that request real-time data from the web.

## Troubleshooting

You might encounter problems if you run the tutorial multiple times. If the Office cache already has an instance of a function with the same name, your add-in gets an error when it sideloads.

To prevent this conflict, [clear the Office cache](../testing/clear-cache.md) before running `npm run start`. If your npm process is already running, enter `npm run stop`, clear the Office cache, and then restart npm.

:::image type="content" source="../images/custom-function-already-exists-error.png" alt-text="An error message in Excel titled 'Error installing functions'. It contains the text 'This add-in wasn't installed because a custom function with the same name already exists'.":::

## Next steps

You created a new custom functions project, tried out a prebuilt function, created a custom function that requests data from the web, and created a custom function that streams data. Next, learn how to [Share custom function data with the task pane](share-data-and-events-between-custom-functions-and-the-task-pane-tutorial.md).
