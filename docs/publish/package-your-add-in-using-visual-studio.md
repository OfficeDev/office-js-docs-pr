---
title: Package your add-in using Visual Studio to prepare for publishing
description: 
ms.date: 11/20/2017 
---


# Package your add-in using Visual Studio to prepare for publishing

Your Office Add-in package contains an XML [manifest file](../overview/add-in-manifests.md) that you'll use to publish the add-in. You'll have to publish the web application files of your project separately. This article describes how to deploy your web project and package your add-in by using Visual Studio 2015.

## To deploy your web project using Visual Studio 2015

Complete the following steps to deploy your web project using Visual Studio 2015.

1. In  **Solution Explorer**, open the shortcut menu for the add-in project, and then choose  **Publish**.
    
    The  **Publish your add-in** page appears.
    
2. In the  **Current profile** drop-down list, select a profile or choose **New ...** to create a new profile.
    
    > [!NOTE]
    > A publish profile specifies the server you are deploying to, the credentials needed to log on to the server, the databases to deploy, and other deployment options.

    If you choose  **New ...**, the  **Create publishing profile** wizard appears. You can use this wizard to import a publishing profile from a web site hosting provider such as Microsoft Azure or create a new profile and add your server, credentials, and other settings in the next procedure.
    
    For more information about importing publishing profiles or creating new publishing profiles, see [Creating a Publish Profile](http://msdn.microsoft.com/en-us/library/dd465337.aspx#creating_a_profile).
    
3. In the  **Publish your add-in** page, choose the **Deploy your web project** link.
    
    The  **Publish Web** dialog box appears. For more information about using this wizard, see [How to: Deploy a Web Project using On-Click Publishing in Visual Studio](http://msdn.microsoft.com/en-us/library/dd465337.aspx).
    

## To package your add-in using Visual Studio 2015

Complete the following steps to package your add-in using Visual Studio 2015.

1. In the **Publish your add-in** page, choose the **Package the add-in** link.
    
    The **Publish Office and SharePoint Add-ins** wizard appears.
    
2. In the **Where is your website hosted?** dropdown list, select or enter the URL of the website that will host the content files of your add-in, and then choose **Finish**.
    
    You must specify an address that begins with the HTTPS prefix to complete this wizard. Although using an HTTPS endpoint for your website is generally recommended, doing so is not required if you don't plan to publish your add-in to the Office Store. If you want to use an HTTP endpoint for your website, you can open the XML manifest file in a text editor after the package has been created and replace the HTTPS prefix of your website with an HTTP prefix. For more information, see [Why do my apps and add-ins have to be SSL-secured?](https://dev.office.com/officestore/docs/office-store-submission-faq#why-do-my-apps-and-add-ins-have-to-be-ssl-secured).
    
    > [!NOTE]
    > Azure websites automatically provide an HTTPS endpoint.

    Visual Studio generates the files that you need to publish your add-in and then opens the publish output folder. 
    
If you plan to submit your add-in to the Office Store, you can choose the **Perform a validation check** link to identify any issues that will prevent your add-in from being accepted. You should address all issues before you submit your add-in to the store.

You can now upload your XML manifest to the appropriate location to [publish your add-in](../publish/publish.md). You can find the XML manifest in `OfficeAppManifests` in the `app.publish` folder. For example:

 `%UserProfile%\Documents\Visual Studio 2015\Projects\MyApp\bin\Debug\app.publish\OfficeAppManifests`


## Additional resources

- [Publish your Office Add-in](../publish/publish.md)
- [Submit Office and SharePoint Add-ins and Office 365 web apps to the Office Store](https://dev.office.com/officestore/docs/submit-to-the-office-store)
    
