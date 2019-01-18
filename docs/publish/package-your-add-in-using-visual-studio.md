---
title: Package your add-in using Visual Studio to prepare for publishing | Microsoft Docs
description: How to deploy your web project and package your add-in by using Visual Studio 2017.
ms.date: 01/25/2018
localization_priority: Priority
---


# Package your add-in using Visual Studio to prepare for publishing

Your Office Add-in package contains an XML [manifest file](../develop/add-in-manifests.md) that you'll use to publish the add-in. You'll have to publish the web application files of your project separately. This article describes how to deploy your web project and package your add-in by using Visual Studio 2017.

## To deploy your web project using Visual Studio 2017

Complete the following steps to deploy your web project using Visual Studio 2017.

1. In  **Solution Explorer**, open the shortcut menu for the add-in project, and then choose  **Publish**.
    
    The  **Publish your add-in** page appears.
    
2. In the  **Current profile** drop-down list, select a profile or choose **New ...** to create a new profile.
    
    > [!NOTE]
    > A publish profile specifies the server you are deploying to, the credentials needed to log on to the server, the databases to deploy, and other deployment options.

    If you choose  **New ...**, a wizard appears with the **Create publishing profile** page. You can use this wizard to import a publishing profile from a web site hosting provider such as Microsoft Azure or create a new profile and add your server, credentials, and other settings in the next procedure.
    
    For more information about importing publishing profiles or creating new publishing profiles, see [Creating a Publish Profile](https://msdn.microsoft.com/library/dd465337.aspx#creating_a_profile).
    
3. On the **Publish your add-in** page, choose the **Deploy your web project** link.
    
    The  **Publish** dialog box appears. For more information about using this wizard, see [How to: Deploy a Web Project using On-Click Publishing in Visual Studio](https://msdn.microsoft.com/library/dd465337.aspx).
    

## To package your add-in using Visual Studio 2017

Complete the following steps to package your add-in using Visual Studio 2017.

1. In the **Publish your add-in** page, choose the **Package the add-in** button.
    
    A wizard appears with the **Package the add-in** page.
    
2. In the **Where is your website hosted?** box, enter the URL of the website that will host the content files of your add-in, and then choose **Finish**.
    
    > [!IMPORTANT]
    > [!include[HTTPS guidance](../includes/https-guidance.md)] Azure websites automatically provide an HTTPS endpoint.

    Visual Studio generates the files that you need to publish your add-in and then opens the publish output folder.
    
If you plan to submit your add-in to AppSource, you can choose the **Perform a validation check** button to identify any issues that will prevent your add-in from being accepted. You should address all issues before you submit your add-in to the store.

You can now upload your XML manifest to the appropriate location to [publish your add-in](../publish/publish.md). You can find the XML manifest in `OfficeAppManifests` in the `app.publish` folder. For example:

 `%UserProfile%\Documents\Visual Studio 2017\Projects\MyApp\bin\Debug\app.publish\OfficeAppManifests`


## See also

- [Publish your Office Add-in](../publish/publish.md)
- [Make your solutions available in AppSource and within Office](https://docs.microsoft.com/office/dev/store/submit-to-the-office-store)
    
