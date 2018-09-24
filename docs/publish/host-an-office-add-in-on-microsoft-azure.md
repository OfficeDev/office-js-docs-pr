---
title: Host an Office Add-in on Microsoft Azure
description: ''
ms.date: 01/25/2018
---



# Host an Office Add-in on Microsoft Azure

The simplest Office Add-in is made up of an XML manifest file and an HTML page. The XML manifest file describes the add-in's characteristics, such as its name, what Office client applications it can run in, and the URL for the add-in's HTML page. The HTML page is contained in a web app that users interact with when they install and run your add-in within an Office client application. You can host the web app of an Office Add-in on any web hosting platform, including Azure.

This article describes how to deploy an add-in web app to Azure and [sideload the add-in](../testing/create-a-network-shared-folder-catalog-for-task-pane-and-content-add-ins.md) for testing in an Office client application.

## Prerequisites 

1. Install [Visual Studio 2017](https://www.visualstudio.com/downloads) and choose to include the **Azure development** workload.

    > [!NOTE]
    > If you've previously installed Visual Studio 2017, [use the Visual Studio Installer](https://docs.microsoft.com/visualstudio/install/modify-visual-studio) to ensure that the **Azure development** workload is installed. 

2. Install Office. 
    
    > [!NOTE]
    > If you don't already have Office, you can [register for a free 1-month trial](http://office.microsoft.com/try/?WT%2Eintid1=ODC%5FENUS%5FFX101785584%5FXT104056786).

3.  Obtain an Azure subscription.
    
    > [!NOTE]
    > If don't already have an Azure subscription, you can [get one as part of your MSDN subscription](http://www.windowsazure.com/pricing/member-offers/msdn-benefits/) or [register for a free trial](https://azure.microsoft.com/pricing/free-trial). 

## Step 1: Create a shared folder to host your add-in XML manifest file

1. Open File Explorer on your development computer.
    
2. Right-click the C:\ drive and then choose **New** > **Folder**.
    
3. Name the new folder AddinManifests.
    
4. Right-click the AddinManifests folder and then choose **Share with** > **Specific people**.
    
5. In **File Sharing**, choose the drop-down arrow and then choose **Everyone** > **Add** > **Share**.
    
> [!NOTE]
> In this walkthrough, you're using a local file share as a trusted catalog where you'll store the add-in XML manifest file. In a real-world scenario, you might instead choose to [deploy the XML manifest file to a SharePoint catalog](../publish/publish-task-pane-and-content-add-ins-to-an-add-in-catalog.md) or [publish the add-in to AppSource](https://docs.microsoft.com/office/dev/store/submit-to-the-office-store).

## Step 2: Add the file share to the Trusted Add-ins catalog

1.  Start Word and create a document.

    > [!NOTE]
    > Although this example uses Word, you can use any Office application that supports Office Add-ins such as Excel, Outlook, PowerPoint, or Project.
    
2.  Choose **File** > **Options**.
    
3.  In the **Word Options** dialog box, choose **Trust Center** and then choose **Trust Center Settings**. 
    
4.  In the **Trust Center** dialog box, choose **Trusted Add-in Catalogs**. Enter the universal naming convention (UNC) path for the file share you created earlier as the **Catalog URL** (for example, \\\YourMachineName\AddinManifests), and then choose **Add catalog**. 
    
5. Select the check box for **Show in Menu**. 

    > [!NOTE]
    > When you store an add-in XML manifest file on a share that is specified as a trusted web add-in catalog, the add-in appears under **Shared Folder** in the **Office Add-ins** dialog box when the user navigates to the **Insert** tab in the ribbon and chooses **My Add-ins**.

6. Close Word.

## Step 3: Create a web app in Azure

Create an empty web app in Azure either by using [Visual Studio 2017](../publish/host-an-office-add-in-on-microsoft-azure.md#using-visual-studio-2017) or by using the [Azure portal](../publish/host-an-office-add-in-on-microsoft-azure.md#using-the-azure-portal).

### Using Visual Studio 2017

To create the web app using Visual Studio 2017, complete the following steps.

1. In Visual Studio, in the **View** menu, choose **Server Explorer**. Right-click **Azure** and choose **Connect to Microsoft Azure subscription**. Follow the instructions for connecting to your Azure subscription.
    
2. In Visual Studio, in **Server Explorer**, expand **Azure**, right-click **App Service**, and then choose **Create New App Service**.
    
3. In the **Create App Service** dialog box, provide this information:
    
      - Enter a unique **Web App Name** for your site. Azure verifies that the site name is unique across the azurewebsites.net domain.

      - Choose the **Subscription** to use for creating this site.

      - Choose the **Resource Group** for your site. If you create a new group, you also need to name it.
    
      - Choose the **App Service Plan** to use for creating this site. If you create a new plan, you also need to name it.
       
      - Choose **Create**.

    The new web app appears in **Server Explorer** under **Azure** >> **App Service** >> (the chosen resouce group).
    
4. Right-click the new web app and then choose **View in Browser**. Your browser opens and displays a webpage with the message "Your App Service app has been created."
    
5. In the browser address bar, change the URL for the web app so that it uses HTTPS and press **Enter** to confirm that the HTTPS protocol is enabled. 

    > [!IMPORTANT]
    > [!include[HTTPS guidance](../includes/https-guidance.md)] Azure websites automatically provide an HTTPS endpoint.
    
### Using the Azure portal

To create the web app using the Azure portal, complete the following steps.

1. Log on to the [Azure portal](https://portal.azure.com/) using your Azure credentials.
    
2. Choose **New** > **Web + Mobile** > **Web App**. 

3. In the **Web App Create** dialog box, provide this information:
    
      - Enter a unique **App name** for your site. Azure verifies that the site name is unique across the azureweb apps.net domain.

      - Choose the **Subscription** to use for creating this site.

      - Choose the **Resource Group** for your site. If you create a new group, you also need to name it.

      - Choose the **OS** for your site.
    
      - Choose the **App Service plan** to use for creating this site. If you create a new plan, you also need to name it.
       
      - Choose **Create**.

4. Choose **Notifications** (the bell icon that is located along the top edge of the Azure portal) and then choose the **Deployments succeeded** notification to open the site's **Overview** page in the Azure portal.

    > [!NOTE]
    > The notification will change from **Deployment in progress** to **Deployments succeeded** when the site deployment completes.

5. In the **Essentials** section of the site's **Overview** page in the Azure portal, choose the URL that is displayed under **URL**. Your browser opens and displays a webpage with the message "Your App Service app has been created." 
    
6. In the browser address bar, change the URL for the web app so that it uses HTTPS and press **Enter** to confirm that the HTTPS protocol is enabled. 

    > [!IMPORTANT]
    > [!include[HTTPS guidance](../includes/https-guidance.md)] Azure websites automatically provide an HTTPS endpoint.

## Step 4: Create an Office Add-in in Visual Studio

1. Start Visual Studio as an administrator.
    
2. Choose **File** > **New** > **Project**.
    
3. Under **Templates**, expand **Visual C#** (or **Visual Basic**), expand **Office/SharePoint**, and then choose **Add-ins**.
    
4. Choose **Word Web Add-in**, and then choose **OK** to accept the default settings.
       
Visual Studio creates a basic Word add-in that you'll be able to publish as-is, without making any changes to its web project.

## Step 5: Publish your Office Add-in web app to Azure

1. With your add-in project open in Visual Studio, expand the solution node in **Solution Explorer** so that you see both projects for the solution.
    
2. Right-click the web project and then choose **Publish**. The web project contains Office Add-in web app files so this is the project that you publish to Azure.
    
3. On the **Publish** tab:

      - Choose **Microsoft Azure App Service**.
      
      - Choose **Select Existing**.

      - Choose **Publish**. 

6. In the **App Service** dialog box, find and choose the web app that you created in [Step 3: Create a web app in Azure](../publish/host-an-office-add-in-on-microsoft-azure.md#step-3-create-a-web-app-in-azure) and then choose **OK**. 

    Visual Studio publishes the web project for your Office Add-in to your Azure web app. When Visual Studio finishes publishing the web project, your browser opens and shows a webpage with the text "Your App Service app has been created." This is the current default page for the web app.

7. To see the webpage for your add-in, change the URL so that it uses HTTPS and specifies the path of your add-in's HTML page (for example: https://YourDomain.azurewebsites.net/Home.html). This confirms that your add-in's web app is now hosted on Azure. Copy the root URL (for example: https://YourDomain.azurewebsites.net); you'll need it when you edit the add-in manifest file later in this article.
    
## Step 6: Edit and deploy the add-in XML manifest file

1. In Visual Studio with the sample Office Add-in open in **Solution Explorer**, expand the solution so that both projects show.
    
2. Expand the Office Add-in project (for example WordWebAddIn), right-click the manifest folder, and then choose **Open**. The add-in XML manifest file opens.
    
3. In the XML manifest file, find and replace all instances of "~remoteAppUrl" with the root URL of the add-in web app on Azure. This is the URL that you copied earlier after you published the add-in web app to Azure (for example: https://YourDomain.azurewebsites.net). 
    
4. Choose **File** and then choose **Save All**. Close the add-in XML manifest file.
    
5. Back in **Solution Explorer**, right-click the manifest folder and choose **Open Folder In File Explorer**.
    
6. Copy the add-in XML manifest file (for example, WordWebAddIn.xml). 
    
7. Browse to the network file share that you created in [Step 1: Create a shared folder](../publish/host-an-office-add-in-on-microsoft-azure.md#step-1-create-a-shared-folder-to-host-your-add-in-xml-manifest-file) and paste the manifest file into the folder.

## Step 7: Insert and run the add-in in the Office client application

1. Start Word and create a document.
    
2. On the ribbon, choose **Insert** > **My Add-ins**. 
    
3. In the **Office Add-ins** dialog box, choose **SHARED FOLDER**. Word scans the folder that you listed as a trusted add-ins catalog (in [Step 2: Add the file share to the Trusted Add-ins catalog](../publish/host-an-office-add-in-on-microsoft-azure.md#step-2-add-the-file-share-to-the-trusted-add-ins-catalog)) and shows the add-ins in the dialog box. You should see an icon for your sample add-in.
    
4. Choose the icon for your add-in and then choose **Add**. A **Show Taskpane** button for your add-in is added to the ribbon. 

5. On the ribbon of the **Home** tab, choose the **Show Taskpane** button. The add-in opens in a task pane to the right of the current document.
    
6. Verify that the add-in works by selecting some text in the document and choosing the **Highlight!** button in the task pane. 

## See also

- [Publish your Office Add-in](../publish/publish.md)
- [Package your add-in using Visual Studio to prepare for publishing](../publish/package-your-add-in-using-visual-studio.md)
    
