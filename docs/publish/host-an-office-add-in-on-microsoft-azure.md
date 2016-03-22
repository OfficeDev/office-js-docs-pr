
# Host an Office Add-in on Microsoft Azure

The simplest Office Add-in is made up of an XML manifest file and an HTML page. The XML manifest file describes the add-in's characteristics, such as its name, what Office client applications it can run in, and the URL for the add-in's HTML page. The HTML page is contained in an Office Add-in website and users see it and interact with it when they install and run your add-in. For more information about Office Add-ins in general, see [Office Add-ins platform overview](../../docs/overview/office-add-ins.md).

You can host the website of an Office Add-in on many of web hosting platforms, including Azure. To host an Office Add-in on Azure, you publish the Office Add-in to a Azure website. 

This topic assumes that you don't have experience using Azure. When you finish, you'll have a simple Office Add-in that has its website hosted on Azure.
You'll learn:

- How to add a trusted add-in catalog to Office 2013
    
- How to create a website in Azure using Visual Studio 2015 or the Azure management portal
    
- How to publish to and host an Office Add-in on a Azure website
    

**A sample Office Add-in website hosted on Azure**


![App for Office website hosted in Microsoft Azure](../../images/off15app_HowToPublishA4OtoAzure_fig17.png)


## Set up your development computer with Azure SDK for .NET, an Azure subscription, and Office 2013



1. Install the Azure SDK for .NET from the [Azure downloads page](http://azure.microsoft.com/en-us/downloads/). If you don't have Visual Studio installed, Visual Studio Express for Web is installed with the SDK.
    
    - Under  **Languages**, choose  **.NET**.
    
    - Choose the version of the Azure .NET SDK that matches your version of Visual Studio, if you already have Visual Studio installed.
    
    - When you're asked whether to run or save the installation executable, choose  **Run**.
    
    - In the Web Platform Installer window, choose  **Install**.
    
2. Install Office 2013 if you haven't already. 
    
     >**Note**  You can get a [trial version for one month](http://office.microsoft.com/en-us/try/?WT%2Eintid1=ODC%5FENUS%5FFX101785584%5FXT104056786).
3. Get your Azure account.
    
     >**Note**  If you're a Microsoft Developer Network (MSDN) subscriber, [you get an Azure subscription as part of your MSDN subscription](http://www.windowsazure.com/en-us/pricing/member-offers/msdn-benefits/).If you're not an MSDN subscriber, you can still [get a free trial of Azure at the Windows Azure website](https://azure.microsoft.com/en-us/pricing/free-trial/). 
To keep the walkthrough simple and focused on using Azure with an Office Add-in, you'll use a local file share as a trusted catalog where you'll store the add-in's XML manifest file. For an add-in that you intend to be used in a business or more than one business, you might keep the add-in manifest file in SharePoint, or publish the add-in to the Office Store. See Publishing basics in [Office Add-ins platform overview](../../docs/overview/office-add-ins.md).


## Step 1: Create a network file share to host your add-in manifest file



1. Open File Explorer (or Windows Explorer if you're using Windows 7 or an earlier version of Windows) on your development computer.
    
2. Right-click the C:\ drive, and then choose  **New** > **Folder**.
    
3. Name the new folder AddinManifests.
    
4. Right-click the AddinManifests folder, and then choose  **Share with** > **Specific people**.
    
5. In  **File Sharing**, click the drop-down arrow and then choose  **Everyone**>  **Add** > **Share**.
    

## Step 2: Add the file share to the Trusted Add-ins catalog so that Office client applications will trust the location where you install Office Add-ins



1.  Start Word 2013 and create a document. (Although we're using Word 2013 in this example, you could use any Office application that supports Office Add-ins like Excel, Outlook, PowerPoint, or Project 2013.)
    
2.  Choose **File** > **Options**.
    
3.  In **Word Options**, choose  **Trust Center**, and then choose  **Trust Center Setting**. 
    
4.  In the **Trust Center**, click  **Trusted Add-in Catalogs**. Enter the universal naming convention (UNC) path for the file share you created earlier as the  **Catalog URL**. For example, \\YourMachineName\AddinManifests. Then choose  **Add catalog**. 
    
5. Select the check box for  **Show in Menu**. When you store an add-in XML manifest file on a share that is a trusted add-in catalog, the add-in appears under  **Shared Folder** in the **Office Add-ins** dialog box.
    

## Step 3: Create a website in Azure


There are a couple of ways you can create an empty Azure website. If you're using Visual Studio 2015, follow the steps in [Using Visual Studio 2015 ](../publish/host-an-office-add-in-on-microsoft-azure.md#bk_usingVS2013) to create a Azure website from within the Visual Studio IDE. You can also follow the steps in [Using the Azure management portal](../publish/host-an-office-add-in-on-microsoft-azure.md#bk_createwebsiteusingAzureportal) to create the Azure website.


### Using Visual Studio 2015



1. In Visual Studio, in the  **View** menu choose **Server Explorer**. Right click  **Azure** and choose **Connect to Microsoft Azure subscription**. Follow the instructions for connecting to your Azure subscription.
    
2. In Visual Studio, in  **Server Explorer**, expand  **Azure**, right-click  **App Service**, and then choose  **Create New Web App**.
    
3. In the  **Create Web App on Windows Azure** dialog box, provide this information:
    
      - Enter a unique  **Web App name** for your site. Azure verifies that the site name is unique across the azurewebsites.net domain.
    
  - Choose the  **App Service plan** you're using to authorize creating this website. If you create a new plan, you also need to name it.
    
  - Choose the  **Resource group** for your site. If you create a new group, you also need to name it.
    
  - Choose a geographical  **Region** appropriate for you.
    
  - For  **Database server:**, accept the default of  **No database** and then choose **Create**.
    

    The new website appears under the chosen resource group under  **App Service** under **Azure** in **Server Explorer**.
    
4. Right-click the new website, and then choose  **View in Browser**. Your browser opens and displays a webpage with the message "This web site has been successfully created."
    
5. In the browser address bar, change the URL for the website so that it uses HTTPS and press  **Enter** to confirm that the HTTPS protocol is enabled. The Office Add-in model requires add-ins to use the HTTPS protocol.
    
6. In Visual Studio 2015, right-click the new website in  **Server Explorer**, choose  **Download Publish Profile** and then save the profile to your computer. The publish profile contains your credentials and enables you to [Step 5: Publish your Office Add-in to the Azure website](../publish/host-an-office-add-in-on-microsoft-azure.md#bk_publishA4OtoAzure).
    

### Using the Azure management portal



1. Log in to the [Azure management portal](https://manage.windowsazure.com/) using your Azure account.
    
2. Choose  **NEW**>  **COMPUTE**>  **WEB APP**>  **QUICK CREATE**. 
    
3. Under  **URL**, enter a unique site name to complete the URL for the website. The management portal verifies that the site name is unique across the azurewebsites.net domain.
    
4. Choose a geographical  **REGION** appropriate for your site.
    
5. Choose  **CREATE WEB APP**. The Azure management portal creates the website and redirects to the  **web sites** page where you can see the website's status.
    
    When the website status is  **Running**, choose the URL for the website under the  **NAME** column. Your browser opens and display a webpage with the message **Your web app has been created!**. 
    
    In the browser address bar, change the URL for the website so that it uses HTTPS and press  **Enter** to confirm that the HTTPS protocol is enabled. The Office Add-in model requires add-ins to use the HTTPS protocol.
    
6. On the  **web apps** page, choose the new website.
    
7. Under  **Publish your app**, choose  **Download the publish profile**, which saves the publish profile to your computer. Remember the file name and the location because you will need this later.
    
    The publish profile contains your credentials and enables you to securely publish to Azure. 
    

## Step 4: Create an Office Add-in in Visual Studio



1. Start Visual Studio as an administrator.
    
2. Choose  **File**>  **New** > **Project**.
    
3. Under  **Templates**, expand  **Visual C#** (or **Visual Basic**), expand  **Office/SharePoint**, and then choose  **Office Add-ins**.
    
4. Choose  **Office Add-in**, and then choose **OK** to accept the default settings.
    
5. When  **Create Office Add-in** appears, leave the default choice for a Task pane add-in and choose **Next**.
    
6. On the next page, clear all check boxes except for Word, then choose  **Finish**.
    
Your basic Office Add-in is created and ready to publish to Azure.
Because we're focusing on how to publish to Azure, you won't make any changes to the sample add-in that you created by using the standard Office Add-in template in Visual Studio.

## Step 5: Publish your Office Add-in to the Azure website



1. With your sample add-in open in Visual Studio, expand the solution node in  **Solution Explorer** so that you see both projects for the solution.
    
2. Right-click the web project, and then choose  **Publish**. 
    
    The web project contains Office Add-in website files so this is the project that you publish to Azure.
    
3. In  **Publish Web**, choose  **Import**. 
    
4. In  **Import Publish Settings**, choose  **Browse**, and then browse to the place where you saved your publish profile earlier in this topic. Choose  **OK** to import your profile.
    
5. In  **Publish Web**, on the  **Connection** tab, accept the defaults and choose **Next**. 
    
    Choose  **Next ** again to accept the default settings.
    
6. On the  **Preview** tab, choose **Start Preview**. The preview shows you all the files in the web project that will be published to the Azure website.
    
7. Choose  **Publish**. Visual Studio publishes the web project for your Office Add-in to your Azure Web Site. 
    
8. When Visual Studio finishes publishing the web project, your browser opens and shows a webpage with the text "This web app has been successfully created." This is the current default page for the website.
    
    To see the webpage for your add-in, change the URL to use https: and add the path of your add-in's default HTML page. For example, the changed URL should look like https://YourDomain.azurewebsites.net/Addin/Home/Home.html. This confirms that your add-in's website is now hosted on Azure. Copy this URL because you'll need it when you edit the add-in manifest file later in this topic.
    

## Step 6: Edit the add-in manifest file to point to the Office Add-in on Azure



1. In Visual Studio with the sample Office Add-in open in  **Solution Explorer**, expand the solution so that both projects show.
    
2. Expand the Office Add-in project, for example  **OfficeAdd-in1**, right-click the manifest folder, and then choose  **Open**. The add-in manifest properties page shows.
    
3. For  **Source Location:**, enter the URL for the add-in's main HTML page that you copied in the previous step after you published the add-in, for example, https://YourDomain.azurewebsites.net/Addin/Home/Home.html. 
    
4. Choose  **File**, and then choose  **Save All**. Close the add-in manifest properties page.
    
5. Back in  **Solution Explorer**, right-click the manifest folder and choose  **Open Folder In File Explorer**.
    
6. Copy the add-in manifest file, for example OfficeAdd-in1.xml. 
    
7. Browse to the network file share that you created earlier in the topic and paste the manifest file into the folder.
    

## Step 7: Insert and run the add-in in the Office client application



1. Start Word and open a new document.
    
2. On the Ribbon, choose  **Insert**>  **My Apps**, and then choose  **See all**.
    
3. In the  **Apps for Office** dialog box, choose **SHARED FOLDER**. Office client applications that work with the Office Add-ins model scan the folder that you list as a trusted add-in catalog and show the add-ins in the dialog. You should see the icon for your sample add-in.
    
4. Choose the icon for your add-in and then choose  **Insert**. The add-in is inserted on the side of the client application.
    
5. Test that the add-in is working by creating some text in the document, then selecting the text, and then choosing  **Get data from selection**.
    

## Additional resources



- [Publish your Office Add-in](../publish/publish.md)
    
- [Package your add-in using Napa or Visual Studio to prepare for publishing](../publish/package-your-add-in-using-napa-or-visual-studio.md)
    
