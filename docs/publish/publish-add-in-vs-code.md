---
title: Publish an add-in using Visual Studio Code and Azure
description: How to publish an add-in using Visual Studio Code and Azure Active Directory
ms.date: 08/19/2022
ms.custom: vscode-azure-extension-update-completed
ms.localizationpriority: medium
---

# Publish an add-in developed with Visual Studio Code

This article describes how to publish an Office Add-in that you created using the Yeoman generator and developed with [Visual Studio Code (VS Code)](https://code.visualstudio.com) or any other editor.

> [!NOTE]
> For information about publishing an Office Add-in that you created using Visual Studio, see [Publish your add-in using Visual Studio](package-your-add-in-using-visual-studio.md).

## Publishing an add-in for other users to access

An Office Add-in consists of a web application and a manifest file. The web application defines the add-in's user interface and functionality, while the manifest specifies the location of the web application and defines settings and capabilities of the add-in.

While you're developing, you can run the add-in on your local web server (`localhost`). When you're ready to publish it for other users to access, you'll need to deploy the web application and update the manifest to specify the URL of the deployed application.

When your add-in is working as desired, you can publish it directly through Visual Studio Code using the Azure Storage extension.

## Using Visual Studio Code to publish

>[!NOTE]
> These steps only work for projects created with the Yeoman generator.

1. Open your project from its root folder in Visual Studio Code (VS Code).
2. From the Extensions view in VS Code, search for the Azure Storage extension and install it.
3. Once installed, an Azure icon is added to the Activity Bar. Select it to access the extension. If your Activity Bar is hidden, you won't be able to access the extension. Show the Activity Bar by selecting **View > Appearance > Activity Bar**.
4. Run the extension and select **Sign in to Azure** to sign in to your Azure account. If you don't already have an Azure account, create one by selecting **Create an Azure Account**. Follow the provided steps to set up your account.
5. Once you have signed in to your Azure account, you'll see your Azure storage accounts appear in the extension. If you don't already have a storage account, create one using the **Create Storage Account** option in the command palette. Name your storage account a globally unique name, using only 'a-z' and '0-9'. Note that by default, this creates a storage account and a resource group with the same name. It automatically puts the storage account in West US. This can be adjusted online through [your Azure account](https://portal.azure.com/).
6. Select and hold (right-click) your storage account and choose **Configure Static Website**. You'll be asked to enter the index document name and the 404 document name. Change the index document name from the default `index.html` to **`taskpane.html`**. You may also change the 404 document name but are not required to.
7. Select and hold (right-click) your storage again, this time choosing **Browse Static Website**. Copy the website URL from the browser window that opens.
8. In VS Code, open your project's manifest file (`manifest.xml`) and change any reference to your localhost URL (such as `https://localhost:3000`) to the URL you've copied. This endpoint is the static website URL for your newly created storage account. Save the changes to your manifest file.
9. Open a command line prompt and navigate to the root directory of your add-in project. Then run the following command to prepare all files for production deployment.

    ```command&nbsp;line
    npm run build
    ```

    When the build completes, the **dist** folder in the root directory of your add-in project will contain the files that you'll deploy in subsequent steps.

10. To deploy, select the Files explorer, select and hold (right-click) your **dist** folder, and choose **Deploy to Static Website via Azure Storage**. When prompted, select the storage account you created previously.

    :::image type="content" source="../images/deploy-to-static-website.png" alt-text="Selecting the dist folder, right-clicking, and choosing Deploy to Static Website via Azure Storage.":::

11. When deployment is complete, right-click the storage account that you created previously, and choose **Browse Static Website**. This will open the static web site and display the task pane.

## Deploy custom functions for Excel

If your add-in has custom functions there are a few more steps to enable them on the Azure Storage account. First you need to enable CORS so that Office can access the functions.json file.

1. Right-click the Azure storage account and choose **Open in Portal**.
1. In the Settings group, choose **Resource sharing (CORS)**. You can also use the search box to find this.
1. Create a new CORS rule with the following settings.

    |Property        |Value                        |
    |----------------|-----------------------------|
    |Allowed origins | \*                          |
    |Allowed methods | GET                         |
    |Allowed headers | \*                          |
    |Exposed headers | Access-Control-Allow-Origin |
    |Max age         | 200                         |

1. Choose **Save**.

> [!CAUTION]
> This CORS configuration assumes all files on your server are publicly available to all domains.  

Next you need to add a MIME type for JSON files.

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

1. Open a command line prompt and go to the root directory of your add-in project. Then run the following command to prepare all files for deployment.

    ```command&nbsp;line
    npm run build
    ```

    When the build completes, the **dist** folder in the root directory of your add-in project will contain the files that you'll deploy.

1. To deploy, select the Files explorer, select and hold (right-click) your **dist** folder, and choose **Deploy to Static Website via Azure Storage**. When prompted, select the storage account you created previously. If you already deployed the **dist** folder, you'll be prompted if you want to overwrite the files in the Azure storage with the latest changes.

## See also

- [Develop Office Add-ins with Visual Studio Code](../develop/develop-add-ins-vscode.md)
- [Deploy and publish your Office Add-in](../publish/publish.md)
- [Cross-Origin Resource Sharing (CORS) support for Azure Storage](/rest/api/storageservices/cross-origin-resource-sharing--cors--support-for-the-azure-storage-services)
