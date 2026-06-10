---
title: Publish an Office Add-in from Visual Studio Code to Azure
description: Deploy an Office Add-in web app from Visual Studio Code to Azure, update your manifest, and publish for users.
ms.date: 06/09/2026
ms.custom: vscode-azure-extension-update-completed
ms.localizationpriority: medium
---

# Publish an Office Add-in from Visual Studio Code to Azure

One way to publish your Office Add-in is by deploying its web app to Azure. This article shows the end-to-end flow for add-ins that use the add-in only manifest and that were created with the Yeoman Generator for Office Add-ins.

> [!NOTE]
>
> - For information about publishing an Office Add-in that you created using Visual Studio, see [Publish your add-in using Visual Studio](package-your-add-in-using-visual-studio.md).
> - The process described in this article doesn't apply to add-ins that use the [unified manifest for Microsoft 365](../develop/unified-manifest-overview.md). Add-ins created using Microsoft 365 Agents Toolkit use the unified manifest. For information about publishing an add-in that you created using Agents Toolkit, see [Deploy Teams app to the cloud](/microsoftteams/platform/toolkit/deploy?pivots=visual-studio-code) and [Deploy your first Teams app](/microsoftteams/platform/sbs-gs-javascript?tabs=vscode%2Cvsc%2Cviscode). The latter article is about Teams tab apps, but it is applicable to Office Add-ins created with Agents Toolkit.

## Publish for other users

An Office Add-in includes a manifest file and a web app. The manifest defines key details, such as the Office apps that support your add-in and the web app URL.

During development, you run the add-in on localhost. To publish it for other users, deploy the web app and update the manifest to use the deployed URL.

You can complete this process directly in Visual Studio Code by using the Azure Storage extension.

## Prerequisites

- An Office Add-in project created with the Yeoman generator for Office Add-ins that uses the add-in only manifest.
- [Visual Studio Code](https://code.visualstudio.com).
- An Azure account and permission to create Azure Storage accounts.

## Using Visual Studio Code to publish

> [!NOTE]
> These steps only work for projects created with the Yeoman generator for Office Add-ins, and that use the add-in only manifest. They don't apply if you created the add-in using Agents Toolkit or created it with the Yeoman generator and it uses the unified manifest for Microsoft 365.

1. Open your project from its root folder in Visual Studio Code (VS Code).
1. Select **View** > **Extensions** (<kbd>Ctrl</kbd>+<kbd>Shift</kbd>+<kbd>X</kbd>) to open the Extensions view.
1. Search for the **Azure Storage** extension and install it.
1. Once installed, an Azure icon is added to the **Activity Bar**. Select it to access the extension. If the **Activity Bar** is hidden, open it by selecting **View** > **Appearance** > **Activity Bar**.
1. Select **Sign in to Azure** to sign in to your Azure account. If you don't already have an Azure account, create one by selecting **Create an Azure Account**. Follow the provided steps to set up your account.

    :::image type="content" source="../images/azure-extension-sign-in.png" alt-text="Sign in to Azure button selected in the Azure extension.":::

1. Once you're signed in, your Azure Storage accounts appear in the extension. If you don't already have one, create one by using the **Create Storage Account** command. Use a globally unique account name with only `a-z` and `0-9`. By default, this creates a storage account and a resource group with the same name in West US. You can change these settings in [Azure portal](https://portal.azure.com/).

    :::image type="content" source="../images/azure-extension-create-storage-account.png" alt-text="Selecting Storage accounts > Create Storage Account in the Azure extension.":::

1. Right-click (or select and hold) your storage account and select **Configure Static Website**. You'll be asked to enter the index document name and the 404 document name. Change the index document name from the default `index.html` to **`taskpane.html`**. You may also change the 404 document name but aren't required to.
1. Right-click (or select and hold) your storage account again and this time select **Browse Static Website**. From the browser window that opens, copy the website URL.
1. Open your project's manifest file and change all references to your localhost URL (such as `https://localhost:3000`) to the URL you've copied. This endpoint is the static website URL for your newly created storage account. Save the changes to your manifest file.
1. Open a command line prompt or terminal window and go to the root directory of your add-in project. Run the following command to prepare all files for production deployment.

    ```bash
    npm run build
    ```

    When the build completes, the **dist** folder in the root directory of your add-in project will contain the files that you'll deploy in subsequent steps.

1. In VS Code, go to the Explorer and right-click (or select and hold) the **dist** folder, and select **Deploy to Static Website via Azure Storage**. When prompted, select the storage account you created previously.

    :::image type="content" source="../images/deploy-to-static-website.png" alt-text="Select the dist folder, right-click (or select and hold), and select Deploy to Static Website via Azure Storage.":::

1. When deployment is complete, right-click (or select and hold) the storage account that you created previously and select **Browse Static Website**. This opens the static website and displays the task pane.

1. Finally, [sideload the manifest file](../testing/sideload-office-add-ins-for-testing.md). The add-in then loads from the static website you deployed.

## Deploy custom functions for Excel

If your add-in has custom functions, there are a few more steps to enable them on the Azure Storage account. First, enable CORS so that Office can access the functions.json file.

1. Right-click (or select and hold) the Azure storage account and select **Open in Portal**.
1. In the Settings group, select **Resource sharing (CORS)**. You can also use the search box to find this.
1. Create a new CORS rule for the **Blob service** with the following settings.

    | Property | Value |
    |--|--|
    | Allowed origins | \* |
    | Allowed methods | GET |
    | Allowed headers | \* |
    | Exposed headers | Access-Control-Allow-Origin |
    | Max age | 200 |

1. Select **Save**.

> [!CAUTION]
> This CORS configuration assumes all files on your server are publicly available to all domains.  

Next, add a MIME type for JSON files.

1. Create a new file in the `src` folder named **web.config**.
1. Insert the following XML and save the file.

    ```xml
    <?xml version="1.0"?>
    <configuration>
      <system.webServer>
        <staticContent>
          <mimeMap fileExtension=".json" mimeType="application/json" />
        </staticContent>
      </system.webServer>
    </configuration> 
    ```

1. Open the **webpack.config.js** file.
1. Add the following code in the list of `plugins` to copy the web.config into the bundle when the build runs.

    ```javascript
    new CopyWebpackPlugin({
      patterns: [
      {
        from: "src/web.config",
        to: "src/web.config",
      },
     ],
    }),
    ```

1. Open a command line prompt and go to the root directory of your add-in project. Then run the following command to prepare all files for deployment.

  ```bash
    npm run build
    ```

    When the build completes, the **dist** folder in the root directory of your add-in project will contain the files that you'll deploy.

1. To deploy, in the VS Code **Explorer**, right-click (or select and hold) the **dist** folder and select **Deploy to Static Website via Azure Storage**. When prompted, select the storage account you created previously. If you already deployed the **dist** folder, you'll be prompted if you want to overwrite the files in the Azure storage with the latest changes.

## Deploy updates

[!INCLUDE [General statements about updating an add-in](../includes/deploy-updates-general.md)]

## See also

- [Develop Office Add-ins with Visual Studio Code](../develop/develop-add-ins-vscode.md)
- [Deploy and publish your Office Add-in](../publish/publish.md)
- [Cross-Origin Resource Sharing (CORS) support for Azure Storage](/rest/api/storageservices/cross-origin-resource-sharing--cors--support-for-the-azure-storage-services)
