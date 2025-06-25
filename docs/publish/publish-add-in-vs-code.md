---
title: Publish an add-in using Visual Studio Code and Azure
description: How to publish an add-in using Visual Studio Code and Azure Active Directory
ms.date: 05/19/2025
ms.custom: vscode-azure-extension-update-completed
ms.localizationpriority: medium
---

# Publish an add-in developed with Visual Studio Code

This article describes how to publish an Office Add-in that you created using the Yeoman generator and developed with [Visual Studio Code (VS Code)](https://code.visualstudio.com) or any other editor.

> [!NOTE]
>
> - For information about publishing an Office Add-in that you created using Visual Studio, see [Publish your add-in using Visual Studio](package-your-add-in-using-visual-studio.md).
> - The process described in this article doesn't apply to add-ins that use the [unified manifest for Microsoft 365](../develop/unified-manifest-overview.md). Add-ins created using Microsoft 365 Agents Toolkit use the unified manifest. For information about publishing an add-in that you created using Agents Toolkit, see [Deploy Teams app to the cloud](/microsoftteams/platform/toolkit/deploy?pivots=visual-studio-code) and [Deploy your first Teams app](/microsoftteams/platform/sbs-gs-javascript?tabs=vscode%2Cvsc%2Cviscode). The latter article is about Teams tab apps, but it is applicable to Office Add-ins created with Agents Toolkit.

## Publishing an add-in for other users to access

The simplest Office Add-in is made up of a manifest file and an HTML page. The manifest file describes the add-in's characteristics, such as its name, what Office applications it can run in, and the URL for the add-in's HTML page. The HTML page is contained in a web app that users interact with when they install and run your add-in within an Office application. You can host the web app of an Office Add-in on any web hosting platform, including Azure.

While you're developing, you can run the add-in on your local web server (`localhost`). When you're ready to publish it for other users to access, you'll need to deploy the web application and update the manifest to specify the URL of the deployed application.

When your add-in is working as desired, you can publish it directly through Visual Studio Code using the Azure Storage extension.

## Using Visual Studio Code to publish

>[!NOTE]
> These steps only work for projects created with the Yeoman generator, and that use the add-in only manifest. They don't apply if you created the add-in using Agents Toolkit or created it with the Yeoman generator and it uses the unified manifest for Microsoft 365.

1. Open your project from its root folder in Visual Studio Code (VS Code).
1. Select **View** > **Extensions** (<kbd>Ctrl</kbd>+<kbd>Shift</kbd>+<kbd>X</kbd>) to open the Extensions view.
1. Search for the **Azure Storage** extension and install it.
1. Once installed, an Azure icon is added to the **Activity Bar**. Select it to access the extension. If the **Activity Bar** is hidden, open it by selecting **View** > **Appearance** > **Activity Bar**.
1. Select **Sign in to Azure** to sign in to your Azure account. If you don't already have an Azure account, create one by selecting **Create an Azure Account**. Follow the provided steps to set up your account.

    :::image type="content" source="../images/azure-extension-sign-in.png" alt-text="Sign in to Azure button selected in the Azure extension.":::

1. Once you're signed in, you'll see your Azure storage accounts appear in the extension. If you don't already have a storage account, create one using the **Create Storage Account** option in the command palette. Name your storage account a globally unique name, using only 'a-z' and '0-9'. Note that by default, this creates a storage account and a resource group with the same name. It automatically puts the storage account in West US. This can be adjusted online through [your Azure account](https://portal.azure.com/).

    :::image type="content" source="../images/azure-extension-create-storage-account.png" alt-text="Selecting Storage accounts > Create Storage Account in the Azure extension.":::

1. Right-click (or select and hold) your storage account and select **Configure Static Website**. You'll be asked to enter the index document name and the 404 document name. Change the index document name from the default `index.html` to **`taskpane.html`**. You may also change the 404 document name but aren't required to.
1. Right-click (or select and hold) your storage account again and this time select **Browse Static Website**. From the browser window that opens, copy the website URL.
1. Open your project's manifest file and change all references to your localhost URL (such as `https://localhost:3000`) to the URL you've copied. This endpoint is the static website URL for your newly created storage account. Save the changes to your manifest file.
1. Open a command line prompt or terminal window and go to the root directory of your add-in project. Run the following command to prepare all files for production deployment.

    ```command&nbsp;line
    npm run build
    ```

    When the build completes, the **dist** folder in the root directory of your add-in project will contain the files that you'll deploy in subsequent steps.

1. In VS Code, go to the Explorer and right-click (or select and hold) the **dist** folder, and select **Deploy to Static Website via Azure Storage**. When prompted, select the storage account you created previously.

    :::image type="content" source="../images/deploy-to-static-website.png" alt-text="Select the dist folder, right-click (or select and hold), and select Deploy to Static Website via Azure Storage.":::

1. When deployment is complete, right-click (or select and hold) the storage account that you created previously and select **Browse Static Website**. This opens the static web site and displays the task pane.

1. Finally, [sideload the manifest file](../testing/sideload-office-add-ins-for-testing.md) and the add-in will load from the static web site you just deployed.

## Deploy custom functions for Excel

If your add-in has custom functions, there are a few more steps to enable them on the Azure Storage account. First, enable CORS so that Office can access the functions.json file.

1. Right-click (or select and hold) the Azure storage account and select **Open in Portal**.
1. In the Settings group, select **Resource sharing (CORS)**. You can also use the search box to find this.
1. Create a new CORS rule for the **Blob service** with the following settings.

    |Property        |Value                        |
    |----------------|-----------------------------|
    |Allowed origins | \*                          |
    |Allowed methods | GET                         |
    |Allowed headers | \*                          |
    |Exposed headers | Access-Control-Allow-Origin |
    |Max age         | 200                         |

1. Select **Save**.

> [!CAUTION]
> This CORS configuration assumes all files on your server are publicly available to all domains.  

Next, add a MIME type for JSON files.

1. Create a new file in the /src folder named **web.config**.
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

1. Open a command line prompt and go to the root directory of your add-in project. Then, run the following command to prepare all files for deployment.

    ```command&nbsp;line
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
