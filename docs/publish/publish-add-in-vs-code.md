---
title: Publish an add-in using Visual Studio Code and Azure
description: How to publish an add-in using Visual Studio Code and Azure Active Directory
ms.date: 08/12/2020
localization_priority: Normal
---

# Publish an add-in developed with Visual Studio Code

This article describes how to publish an Office Add-in that you created using the Yeoman generator and developed with [Visual Studio Code (VS Code)](https://code.visualstudio.com) or any other editor.

> [!NOTE]
> For information about publishing an Office Add-in that you created using Visual Studio, see [Publish your add-in using Visual Studio](package-your-add-in-using-visual-studio.md).

## Publishing pre-requisites

An Office Add-in consists of a web application and a manifest file. The web application defines the add-in's user interface and functionality, while the manifest specifies the location of the web application and defines settings and capabilities of the add-in.

While you're developing your add-in, you can run the add-in on your local web server (`localhost`), but when you're ready to publish it for other users to access, you'll need to deploy the web application and update the manifest to specify the URL of the deployed application.

When your add-in is working as desired and you're ready to publish it for other users to access, ensure you have the following pre-requisites:

- An [Azure account](https://azure.microsoft.com/free/). You can start one for free if you don't already have an account.
- An Azure storage account. If you need to set one up after creating your Azure account, do so using instructions in [Create an Azure Storage Account](/azure/developer/javascript/tutorial-vscode-static-website-node-03).

## Publishing an add-in for other users to access

1. Open your browser and navigate to https://portal.azure.com/, then select **Storage accounts** to see your storage accounts. Select the one you would like to use for your add-in.

2. On this same page under the settings menu, choose the **Static website** option. You'll notice **Index document name** and **Error document path** are pre-populated with the filename index.html. Change both fields from index.html to **taskpane.html**. Next, copy the **primary endpoint** URL that you did not change.

![Static website settings in Azure](../images/static-website-in-azure.png)

3. Open your project from its root folder in VS Code. Next, open your project's manifest file (`manifest.xml`) and change any reference to your localhost URL (such as `https://localhost:3000`) to the primary endpoint information you've copied. This endpoint is the static website URL for your newly created storage account. Save the changes you've made to your manifest file.

4. Open a command line prompt. From the command line, navigate to the root directory of your add-in project. Then run the following command to prepare all files for production deployment:

    ```command&nbsp;line
    npm run build
    ```

    When the build completes, the **dist** folder in the root directory of your add-in project will contain the files that you'll deploy in subsequent steps.

5. Next go to the **Azure Storage** explorer in VS Code, expand your subscription, and expand the node for the Azure Storage account that you created in the previous step. Expand the **Blob Containers** node. The $web container is where you deploy your app code.

![Storage nodes listed in the Blob Containers node](../images/azure-storage-container.png)

6. To deploy, select the Files explorer, select and hold (right-click) on your **dist** folder, and choose **Deploy to Static Website**. When prompted, select the storage account you created previously.

![Deploying to a static website](../images/deploy-to-static-website.png)

7. When deployment is complete, a message appears with a **Browse to website** button. Select that button to open the primary endpoint of the deployed app code.

## See also

- [Develop Office Add-ins with Visual Studio Code](../develop/develop-add-ins-vscode.md)
- [Deploy and publish your Office Add-in](../publish/publish.md)
