---
title: Requirements for running Office Add-ins
description: Learn about the client and server requirements that an end user needs to run Office Add-ins.
ms.date: 05/20/2023
ms.localizationpriority: medium
---

# Requirements for running Office Add-ins

This article describes the software and device requirements for running Office Add-ins.

[!INCLUDE [publish policies note](../includes/note-publish-policies.md)]

For a high-level view of where Office Add-ins are currently supported, see [Office client application and platform availability for Office Add-ins](/javascript/api/requirement-sets).

## Server requirements

To be able to install and run any Office Add-in, you first need to deploy the manifest and webpage files for the UI and code of your add-in to the appropriate server locations.

For all types of add-ins (content, Outlook, and task pane add-ins and add-in commands), you need to deploy your add-in's webpage files to a web server, or web hosting service, such as [Microsoft Azure](../publish/host-an-office-add-in-on-microsoft-azure.md).

[!include[HTTPS guidance](../includes/https-guidance.md)]

> [!TIP]
> When you develop and debug an add-in in Visual Studio, Visual Studio deploys and runs your add-in's webpage files locally with IIS Express, and doesn't require an additional web server.

For content and task pane add-ins, in the supported Office client applications - Excel, PowerPoint, Project, or Word - you also need either an [app catalog](../publish/publish-task-pane-and-content-add-ins-to-an-add-in-catalog.md) on SharePoint to upload the add-in's XML manifest file, or you need to deploy the add-in using [Integrated Apps](/microsoft-365/admin/manage/test-and-deploy-microsoft-365-apps).

To test and run an Outlook add-in, the user's Outlook email account must reside on Exchange 2013 or later, which is available through Microsoft 365, Exchange Online, or through an on-premises installation. The user or administrator installs manifest files for Outlook add-ins on that server.

> [!NOTE]
> POP and IMAP email accounts in Outlook don't support Office Add-ins.

## Client requirements: Windows desktop and tablet

The following software is required for developing an Office Add-in for the supported Office desktop clients or web clients that run on Windows-based desktop, laptop, or tablet devices.

- For Windows x86 and x64 desktops, and tablets such as Surface Pro:
  - The 32- or 64-bit version of Office 2013 or a later version, running on Windows 7 or a later version.
  - Excel 2013, Outlook 2013, PowerPoint 2013, Project Professional 2013, Project 2013 SP1, Word 2013, or a later version of the Office client, if you're testing or running an Office Add-in specifically for one of these Office desktop clients. Office desktop clients can be installed on premises or via Click-to-Run on the client computer.

  If you have a valid Microsoft 365 subscription and you don't have access to the Office client, you can [download and install the latest version of Office](https://support.microsoft.com/office/4414eaaf-0478-48be-9c42-23adc4716658).

- Microsoft Edge must be installed, but doesn't have to be the default browser. To support Office Add-ins, the Office client that acts as host uses webview components that are part of Microsoft Edge.

  > [!NOTE]
  >
  > - Strictly speaking, it's possible to develop add-ins on a machine that has Internet Explorer 11 (IE11) installed, but not Microsoft Edge. However, IE11 is used to run add-ins only on certain older combinations of Windows and Office versions. See [Browsers and webview controls used by Office Add-ins](browsers-used-by-office-web-add-ins.md) for more details. We don't recommend using such old environments as your primary add-in development environment. However, if you're likely to have customers of your add-in that are working in these older combinations, we recommend that you support the Trident webview that's provided by Internet Explorer. For more information, see [Support older Microsoft webviews and Office versions](../develop/support-ie-11.md).
  > - Internet Explorer's Enhanced Security Configuration (ESC) must be turned off for Office Web Add-ins to work. If you are using a Windows Server computer as your client when developing add-ins, note that ESC is turned on by default in Windows Server.

- One of the following as the default browser: Internet Explorer 11, or the latest version of Microsoft Edge, Chrome, Firefox, or Safari (Mac OS).
- An HTML and JavaScript editor such as [Visual Studio Code](https://code.visualstudio.com/), [Visual Studio and the Microsoft Developer Tools](https://www.visualstudio.com/features/office-tools-vs), or non-Microsoft web development tool.

## Client requirements: OS X desktop

Outlook on Mac, which is distributed as part of Microsoft 365, supports Outlook add-ins. Running Outlook add-ins in Outlook on Mac has the same requirements as Outlook on Mac itself: the operating system must be at least OS X v10.10 "Yosemite". Because Outlook on Mac uses WebKit as a layout engine to render the add-in pages, there is no additional browser dependency.

The following are the minimum client versions of Office on Mac that support Office Add-ins.

- Word version 15.18 (160109)
- Excel version 15.19 (160206)
- PowerPoint version 15.24 (160614)

## Client requirements: Browser support for Office web clients and SharePoint

Any browser, except Internet Explorer, that supports ECMAScript 5.1, HTML5, and CSS3, such as Microsoft Edge, Chrome, Firefox, or Safari (Mac OS).

## Client requirements: Non-Windows smartphone and tablet

Specifically for Outlook running on smartphones and non-Windows tablet devices, the following software is required for testing and running Outlook add-ins.

| Office application | Device | Operating system | Exchange account | Mobile browser |
|:-----|:-----|:-----|:-----|:-----|
|Outlook on Android|- Android tablets<br>- Android smartphones|- Android 4.4 KitKat or later|On the latest update of Microsoft 365 Apps for business or Exchange Online|Browser not applicable. Use the native app for Android.<sup>1</sup>|
|Outlook on iOS|- iPad tablets<br>- iPhone smartphones|- iOS 11 or later|On the latest update of Microsoft 365 Apps for business or Exchange Online|Browser not applicable. Use the native app for iOS.<sup>1</sup>|
|Outlook on the web (modern)<sup>2</sup>|- iPad 2 or later<br>- Android tablets |- iOS 5 or later<br>- Android 4.4 KitKat or later|On Microsoft 365, Exchange Online|- Microsoft Edge<br>- Chrome<br>- Firefox<br>- Safari|
|Outlook on the web (classic)|- iPhone 4 or later<br>- iPad 2 or later<br>- iPod Touch 4 or later|- iOS 5 or later|On on-premises Exchange Server 2013 or later<sup>3</sup>|- Safari|

> [!NOTE]
> <sup>1</sup> OWA for Android, OWA for iPad, and OWA for iPhone native apps have been [deprecated](https://support.microsoft.com/office/076ec122-4576-4900-bc26-937f84d25a4b).
>
> <sup>2</sup> Modern Outlook on the web on iPhone and Android smartphones is no longer required or available for testing Outlook add-ins.
>
> <sup>3</sup> Add-ins aren't supported in Outlook on Android, on iOS, and modern mobile web with on-premises Exchange accounts.

> [!TIP]
> You can distinguish between classic and modern Outlook in a web browser by checking your mailbox toolbar.
>
> **modern**
>
> ![Partial screenshot of the modern Outlook toolbar.](../images/outlook-on-the-web-new-toolbar.png)
>
> **classic**
>
> ![Partial screenshot of the classic Outlook toolbar.](../images/outlook-on-the-web-classic-toolbar.png)

## See also

- [Office Add-ins platform overview](../overview/office-add-ins.md)
- [Office client application and platform availability for Office Add-ins](/javascript/api/requirement-sets)
- [Browsers and webview controls used by Office Add-ins](browsers-used-by-office-web-add-ins.md)
