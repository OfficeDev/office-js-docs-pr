---
title: Stage an Office Add-in on Microsoft Azure
description: Learn how to deploy an add-in web app to Azure and sideload the add-in for testing in an Office client application.
ms.date: 08/05/2024
ms.localizationpriority: medium
---

# Stage an Office Add-in on Microsoft Azure

The simplest Office Add-in is made up of an manifest file and an HTML page. The manifest file describes the add-in's characteristics, such as its name, what Office desktop clients it can run in, and the URL for the add-in's HTML page. The HTML page is contained in a web app that users interact with when they install and run your add-in within an Office client application. You can host the web app of an Office Add-in on any web hosting platform, including Azure.

This article describes how to deploy an add-in web app to Azure and [sideload the add-in](../testing/create-a-network-shared-folder-catalog-for-task-pane-and-content-add-ins.md) for testing in an Office client application.

## Prerequisites

1. Install [Visual Studio 2019](https://www.visualstudio.com/downloads) and choose to include the **Azure development** workload.

    > [!NOTE]
    > If you've previously installed Visual Studio 2019, [use the Visual Studio Installer](/visualstudio/install/modify-visual-studio) to ensure that the **Azure development** workload is installed.

2. Install Office.

    > [!NOTE]
    > If you don't already have Office, you can [register for a free 1-month trial](https://www.microsoft.com/microsoft-365/try).

3. Obtain an Azure subscription.

    > [!NOTE]
    > If don't already have an Azure subscription, you can [get one as part of your Visual Studio subscription](https://azure.microsoft.com/pricing/member-offers/visual-studio-subscriptions/) or [register for a free trial](https://azure.microsoft.com/pricing/free-trial).

## Step 1: Create a shared folder to host your add-in manifest file

1. Open File Explorer on your development computer.

2. Right-click (or select and hold) the C:\ drive and then choose **New** > **Folder**.

3. Name the new folder AddinManifests.

4. Right-click (or select and hold) the AddinManifests folder and then choose **Share with** > **Specific people**.

5. In **File Sharing**, choose the drop-down arrow and then choose **Everyone** > **Add** > **Share**.

> [!NOTE]
> In this walkthrough, you're using a local file share as a trusted catalog where you'll store the add-in manifest file. In a real-world scenario, you might instead choose to [deploy the manifest file to a SharePoint catalog](../publish/publish-task-pane-and-content-add-ins-to-an-add-in-catalog.md) or [publish the add-in to AppSource](/partner-center/marketplace-offers/submit-to-appsource-via-partner-center).

## Step 2: Add the file share to the Trusted Add-ins catalog

1. Start Word and create a document.

    > [!NOTE]
    > Although this example uses Word, you can use any Office application that supports Office Add-ins such as Excel, Outlook, PowerPoint, or Project.

2. Choose **File** > **Options**.

3. In the **Word Options** dialog box, choose **Trust Center** and then choose **Trust Center Settings**.

4. In the **Trust Center** dialog box, choose **Trusted Add-in Catalogs**. Enter the universal naming convention (UNC) path for the file share you created earlier as the **Catalog URL** (for example, \\\YourMachineName\AddinManifests), and then choose **Add catalog**.

5. Select the check box for **Show in Menu**.

    > [!NOTE]
    > When you store an add-in manifest file on a share that is specified as a trusted web add-in catalog, the add-in appears under **Shared Folder** in the **Office Add-ins** dialog box that launches from **Home** > **Add-ins** > **Get Add-ins**.

6. Close Word.

## Step 3: Create a web app in Azure using the Azure portal

To create the web app using the Azure portal, complete the following steps.

1. Log on to the [Azure portal](https://portal.azure.com/) using your Azure credentials.

2. Under **Azure Services** select **Web Apps**.

3. On the **App Service** page, select **Add**. Provide this information:

      - Choose the **Subscription** to use for creating this site.

      - Choose the **Resource Group** for your site. If you create a new group, you also need to name it.

      - Enter a unique **App name** for your site. Azure verifies that the site name is unique across the azureweb apps.net domain.

      - Choose whether to publish using code or a docker container.

      - Specify a **Runtime stack**.

      - Choose the **OS** for your site.

      - Choose a **Region**.

      - Choose the **App Service plan** to use for creating this site.

      - Choose **Create**.

4. The next page will let you know that your deployment is underway and when it completes. When it is completed, select **Go to resource**.  

5. In the **Overview** section, choose the URL that is displayed under **URL**. Your browser opens and displays a webpage with the message "Your App Service app is up and running."

    > [!IMPORTANT]
    > [!include[HTTPS guidance](../includes/https-guidance.md)] Azure websites automatically provide an HTTPS endpoint.

## Step 4: Create an Office Add-in in Visual Studio

1. Start Visual Studio as an administrator.

2. Choose **Create a new project**.

3. Using the search box, enter **add-in**.

4. Choose **Word Web Add-in** as the project type, and then choose **Next** to accept the default settings.

Visual Studio creates a basic Word add-in that you'll be able to publish as-is, without making any changes to its web project. To make an add-in for a different Office application, such as Excel, repeat the steps and choose a project type with your desired Office application.

## Step 5: Publish your Office Add-in web app to Azure

1. With your add-in project open in Visual Studio, expand the solution node in **Solution Explorer**, then select **App Service**.

2. Right-click (or select and hold) the web project and then choose **Publish**. The web project contains Office Add-in web app files so this is the project that you publish to Azure.

3. On the **Publish** tab:

      - Choose **Microsoft Azure App Service**.

      - Choose **Select Existing**.

      - Choose **Publish**.

4. Visual Studio publishes the web project for your Office Add-in to your Azure web app. When Visual Studio finishes publishing the web project, your browser opens and shows a webpage with the text "Your App Service app has been created." This is the current default page for the web app.

5. Copy the root URL (for example: `https://YourDomain.azurewebsites.net`); you'll need it when you edit the add-in manifest file later in this article.

## Step 6: Edit and deploy the add-in manifest file

1. In Visual Studio with the sample Office Add-in open in **Solution Explorer**, expand the solution so that both projects show.

2. Expand the Office Add-in project (for example WordWebAddIn), right-click (or select and hold) the manifest folder, and then choose **Open**. The add-in manifest file opens.

3. In the manifest file, find and replace all instances of "~remoteAppUrl" with the root URL of the add-in web app on Azure. This is the URL that you copied earlier after you published the add-in web app to Azure (for example: `https://YourDomain.azurewebsites.net`).

4. Choose **File** and then choose **Save All**. Next, Copy the add-in manifest file (for example, WordWebAddIn.xml).

5. Using the **File Explorer** program, browse to the network file share that you created in [Step 1: Create a shared folder](../publish/host-an-office-add-in-on-microsoft-azure.md#step-1-create-a-shared-folder-to-host-your-add-in-manifest-file) and paste the manifest file into the folder.

## Step 7: Insert and run the add-in in the Office client application

1. Start Word and create a document.

2. Select **Home** > **Add-ins**, then select **Get Add-ins**.

3. In the **Office Add-ins** dialog box, choose **SHARED FOLDER**. Word scans the folder that you listed as a trusted add-ins catalog (in [Step 2: Add the file share to the Trusted Add-ins catalog](../publish/host-an-office-add-in-on-microsoft-azure.md#step-2-add-the-file-share-to-the-trusted-add-ins-catalog)) and shows the add-ins in the dialog box. You should see an icon for your sample add-in.

4. Choose the icon for your add-in and then choose **Add**. A **Show Taskpane** button for your add-in is added to the ribbon.

5. On the ribbon of the **Home** tab, choose the **Show Taskpane** button. The add-in opens in a task pane to the right of the current document.

6. Verify that the add-in works by selecting some text in the document and choosing the **Highlight!** button in the task pane.

## Deploy updates

[!INCLUDE [deploy-updates-that-require-admin-consent](../includes/deploy-updates-that-require-admin-consent.md)]

## See also

- [Publish your Office Add-in](../publish/publish.md)
- [Publish your add-in using Visual Studio](../publish/package-your-add-in-using-visual-studio.md)
