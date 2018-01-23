---
title: Requirements for running Office Add-ins
description: ''
ms.date: 01/23/2018
---



# Requirements for running Office Add-ins


This article describes the software and device requirements for running Office Add-ins.

> [!NOTE]
> If you plan to [publish](../publish/publish.md) your add-in to AppSource and make it available within the Office experience, make sure that you conform to the [AppSource validation policies](https://docs.microsoft.com/en-us/office/dev/store/validation-policies). For example, to pass validation, your add-in must work across all platforms that support the methods that you define (for more information, see [section 4.12](https://docs.microsoft.com/en-us/office/dev/store/validation-policies#4-apps-and-add-ins-behave-predictably) and the [Office Add-in host and availability page](../overview/office-add-in-availability.md)). 

For a high-level view of where Office Add-ins are currently supported, see [Office Add-in host and platform availability](../overview/office-add-in-availability.md).

## Server requirements

To be able to install and run any Office Add-in, you first need to deploy the manifest and webpage files for the UI and code of your add-in to the appropriate server locations.

For all types of add-ins (content, Outlook, and task pane add-ins and add-in commands), you need to deploy your add-in's webpage files to a web server, or web hosting service, such as [Microsoft Azure](../publish/host-an-office-add-in-on-microsoft-azure.md).


> [!NOTE]
> When you develop and debug an add-in in Visual Studio, Visual Studio deploys and runs your add-in's webpage files locally with IIS Express, and doesn't require an additional web server. 

For content and task pane add-ins, in the supported Office host applications - Access web apps, Word, Excel, PowerPoint, or Project - you also need an [add-in catalog](../publish/publish-task-pane-and-content-add-ins-to-an-add-in-catalog.md) on SharePoint to upload the add-in's XML manifest file.

To test and run an Outlook add-in, the user's Outlook email account must reside on Exchange 2013 or later, which is available through Office 365, Exchange Online, or through an on-premises installation. The user or administrator installs manifest files for Outlook add-ins on that server.

> [!NOTE]
> POP and IMAP email accounts in Outlook don't support Office Add-ins.




## Client requirements: Windows desktop and tablet

The following software is required for developing an Office Add-in for the supported Office desktop clients or web clients that run on Windows-based desktop, laptop, or tablet devices:


- For Windows x86 and x64 desktops, and tablets such as Surface Pro:
    - The 32- or 64-bit version of Office 2013 or a later version, running on Windows 7 or a later version.
    - Excel 2013, Outlook 2013, PowerPoint 2013, Project Professional 2013, Project 2013 SP1, Word 2013, or a later version of the Office client, if you are testing or running an Office Add-in specifically for one of these Office desktop clients. Office desktop clients can be installed on premises or via Click-to-Run on the client computer.
    
  If you have a valid Office 365 subscription and you do not have access to Office 2013, you can download it via one the following CDN links:		
    - [Office 2013 for Business (.exe)](https://c2rsetup.officeapps.live.com/c2r/download.aspx?productReleaseID=O365BusinessRetail&platform=X86&language=en-us&version=O15GA&source=O15OLSO365) 
    - [Office 2013 for Home (.exe)](https://c2rsetup.officeapps.live.com/c2r/download.aspx?productReleaseID=O365HomePremRetail&platform=X86&language=en-us&version=O15GA&source=O15OLSO365) 

- Internet Explorer 11 or later, which must be installed but doesn't have to be the default browser. To support Office Add-ins, the Office client that acts as host uses browser components that are part of Internet Explorer 11 or later.
- One of the following as the default browser: Internet Explorer 11 or later, or the latest version of Microsoft Edge, Chrome, Firefox, or Safari (Mac OS).
- An HTML and JavaScript editor such as Notepad, [Visual Studio and the Microsoft Developer Tools](https://www.visualstudio.com/features/office-tools-vs), or a third-party web development tool.

## Client requirements: OS X desktop

Outlook for Mac, which is distributed as part of Office 365, supports Outlook add-ins. Running Outlook add-ins on Outlook for Mac has the same requirements as Outlook for Mac itself: the operating system must be at least OS X v10.10 "Yosemite". Because Outlook for Mac uses WebKit as a layout engine to render the add-in pages, there is no additional browser dependency.

The following are the minimum client versions of Office for Mac that support Office Add-ins:

- Word for Mac version 15.18 (160109) 
- Excel for Mac version 15.19 (160206) 
- PowerPoint for Mac version 15.24 (160614)

## Client requirements: Browser support for Office Online web clients and SharePoint

Any browser that supports ECMAScript 5.1, HTML5, and CSS3, such as Internet Explorer 11 or later, or the latest version of Microsoft Edge, Chrome, Firefox, or Safari (Mac OS).


## Client requirements: non-Windows smartphone and tablet

Specifically for OWA for Devices, and Outlook Web App running in a browser on smartphones and non-Windows tablet devices, the following software is required for testing and running Outlook add-ins.


| Host application | Device | Operating system | Exchange account | Mobile browser |
|:-----|:-----|:-----|:-----|:-----|
|OWA for Android|Android smartphones. Technically, those devices considered as "small" or "normal" by [Android OS](https://developer.android.com/guide/practices/screens_support.html).|Android 4.4 KitKat or later|On the latest update of Office 365 for business or Exchange Online|Native add-in for Android, browser not applicable|
|OWA for iPad|iPad 2 or later|iOS 6 or later|On the latest update of Office 365 for business or Exchange Online|Native add-in for iOS, browser not applicable|
|OWA for iPhone|iPhone 4S or later|iOS 6 or later|On the latest update of Office 365 for business or Exchange Online|Native add-in for iOS, browser not applicable|
|Outlook Web App|iPhone 4 or later, iPad 2 or later, iPod Touch 4 or later|iOS 5 or later|On Office 365, Exchange Online, or on premise on Exchange Server 2013 or later|Safari|


## See also

- [Office Add-ins platform overview](../overview/office-add-ins.md)
- [Office Add-in host and platform availability](../overview/office-add-in-availability.md)
