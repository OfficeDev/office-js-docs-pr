---
title: Stage an Outlook add-in on Microsoft Azure with Visual Studio
description: Learn how to deploy an add-in web app to Azure and sideload the add-in for testing in an Office client application.
ms.date: 09/25/2024
ms.localizationpriority: medium
---

# Stage an Outlook add-in on Microsoft Azure with Visual Studio

This article describes how to deploy an Outlook add-in web app to Azure and [sideload the add-in](../testing/test-debug-non-local-server.md) for testing in an Outlook client application.

> [!IMPORTANT]
> The process described in this article applies only to Outlook add-ins. For instructions about staging add-ins for other Office applications on Azure, see [Stage an Office Add-in on Microsoft Azure](host-an-office-add-in-on-microsoft-azure.md).

## Prerequisites

1. Install [Visual Studio 2022](https://www.visualstudio.com/downloads) and choose to include the **Azure development** workload.

    > [!NOTE]
    > If you've previously installed Visual Studio 2022, [use the Visual Studio Installer](/visualstudio/install/modify-visual-studio) to ensure that the **Azure development** workload is installed.

1. Install Office.

    > [!NOTE]
    > If you don't already have Office, you can [register for a free 1-month trial](https://www.microsoft.com/microsoft-365/try).

1. Obtain an Azure subscription.

    > [!NOTE]
    > If don't already have an Azure subscription, you can [get one as part of your Visual Studio subscription](https://azure.microsoft.com/pricing/member-offers/visual-studio-subscriptions/) or [register for a free trial](https://azure.microsoft.com/pricing/free-trial).

## Step 1: Create an Outlook add-in in Visual Studio

1. Start Visual Studio as an administrator.

1. Choose **Create a new project**.

1. Using the search box, enter **add-in**.

1. Choose **Outlook Web Add-in** as the project type, and then choose **Next** to accept the default settings.

Visual Studio creates a basic Outlook add-in that you'll be able to publish as-is, without making any changes to its web project.

## Step 2: Publish your Outlook add-in web app to Azure

1. With your add-in project open in Visual Studio, right-click (or select and hold) the web project and then choose **Publish**.

1. Follow the instructions at [Publish your web app](/azure/app-service/quickstart-dotnetcore?tabs=netframework48&pivots=development-environment-vs#2-publish-your-web-app). Skip the article sections that precede **Publish your web app**, but be sure that the **Visual Studio** button is selected at the top of the page.

   Visual Studio publishes the web project for your Outlook add-in to your Azure web app. When Visual Studio finishes publishing the web project, your browser opens and shows a webpage with the text "Your web app is running and waiting for your content." This is the current default page for the web app.

1. Copy the root URL (for example: `https://YourDomain.azurewebsites.net`); you'll need it when you edit the add-in manifest file later in this article.

## Step 3: Edit and deploy the add-in manifest file

1. In Visual Studio with the sample Outlook add-in open in **Solution Explorer**, expand the solution so that both projects show.

1. Expand the Outlook add-in project (for example OutlookWebAddIn), right-click (or select and hold) the manifest folder, and then choose **Open**. The add-in manifest file opens.

1. In the manifest file, find and replace all instances of "~remoteAppUrl" with the root URL of the add-in web app on Azure. This is the URL that you copied earlier after you published the add-in web app to Azure; for example, `https://YourDomain.azurewebsites.net`.

## Step 4: Sideload the manifest to Outlook

Follow the guidance at [Sideload an Outlook add-in that uses an add-in only manifest](../outlook/sideload-outlook-add-ins-for-testing#sideload-an-add-in-that-uses-an-xml-manifest) to sideload the add-in.

## See also

- [Publish your Office Add-in](../publish/publish.md)
- [Publish your add-in using Visual Studio](../publish/package-your-add-in-using-visual-studio.md)
