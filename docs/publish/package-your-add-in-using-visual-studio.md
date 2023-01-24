---
title: Publish your add-in using Visual Studio
description: How to deploy your web project and package your add-in by using Visual Studio 2019.
ms.date: 01/23/2023
ms.localizationpriority: medium
---

# Publish your add-in using Visual Studio

Your Office Add-in package contains an XML [manifest file](../develop/add-in-manifests.md) that you'll use to publish the add-in. You'll have to publish the web application files of your project separately. This article describes how to deploy your web project and package your add-in by using Visual Studio 2019.

> [!NOTE]
> For information about publishing an Office Add-in that you created using the Yeoman generator and developed with Visual Studio Code or any other editor, see [Publish an add-in developed with Visual Studio Code](publish-add-in-vs-code.md).

## To deploy your web project using Visual Studio 2019

Complete the following steps to deploy your web project using Visual Studio 2019.

1. From the **Build** tab, choose **Publish [Name of your add-in]**.

2. In the **Pick a publish target** window, choose one of the options to publish to your preferred target. Each publish target requires you to include more information to get started, such as an Azure Virtual Machine or folder location. Once you have specified a publish location and filled in all of the information required, select **Publish**

    > [!NOTE]
    > Picking a publish target specifies the server you are deploying to, the credentials needed to sign in to the server, the databases to deploy, and other deployment options.

3. For more information about deployment steps for each publish target option, see [First look at deployment in Visual Studio](/visualstudio/deployment/deploying-applications-services-and-components?view=vs-2019&preserve-view=true).

## To package and publish your add-in using IIS, FTP, or Web Deploy using Visual Studio 2019

Complete the following steps to package your add-in using Visual Studio 2019.

1. From the **Build** tab, choose **Publish [Name of your add-in]**.
2. In the **Pick a publish target** window, choose **IIS, FTP, etc**, and select **Configure**. Next, select **Publish**.
3. A wizard appears that will help guide you through the process. Ensure the publish method is your preferred method, such as Web Deploy.
4. In the **Destination URL** box, enter the URL of the website that will host the content files of your add-in, and then select **Next**. If you plan to submit your add-in to AppSource, you can choose the **Validate Connection** button to identify any issues that will prevent your add-in from being accepted. You should address all issues before you submit your add-in to the store.
5. Confirm any settings desired including **File Publish Options** and select **Save**.

    > [!IMPORTANT]
    > [!include[HTTPS guidance](../includes/https-guidance.md)] Azure websites automatically provide an HTTPS endpoint.

You can now upload your XML manifest to the appropriate location to [publish your add-in](../publish/publish.md). You can find the XML manifest in `OfficeAppManifests` in the `app.publish` folder. For example:

 `%UserProfile%\Documents\Visual Studio 2019\Projects\MyApp\bin\Debug\app.publish\OfficeAppManifests`

## Deploy updates

[!INCLUDE [deploy-updates-that-require-admin-consent](../includes/deploy-updates-that-require-admin-consent.md)]

## See also

- [Publish your Office Add-in](../publish/publish.md)
- [Make your solutions available in AppSource and within Office](/office/dev/store/submit-to-the-office-store)
