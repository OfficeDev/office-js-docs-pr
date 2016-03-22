
# Requirements for running Office Add-ins



This article describes the software and device requirements for running Office Add-ins.

## Server requirements

To be able to install and run any Office Add-in, you first need to deploy the manifest and webpage files for the UI and code of your add-in to the appropriate server locations.

For all types of add-ins (content, Outlook, and task pane add-ins and add-in commands), you need to deploy your add-in's webpage files to a web server, or web hosting service, such as [Microsoft Azure](../publish/host-an-office-add-in-on-microsoft-azure.md). 


 >**Note**   When you develop and debug an add-in in Visual Studio, Visual Studio deploys and runs your add-in's webpage files locally with IIS Express, and doesn't require an additional web server. Similarly, when you develop and debug with Napa in the browser, it deploys and runs your add-in's webpage files from storage associated with the account you used to sign into Napa.

For content and task pane add-ins, in the supported Office host applications - Access web apps, Word, Excel, PowerPoint, or Project - you also need a [ network file share](../publish/create-a-network-shared-folder-catalog-for-task-pane-and-content-add-ins.md) or an [add-in catalog](../publish/publish-task-pane-and-content-add-ins-to-an-add-in-catalog.md) on SharePoint to upload the add-in's XML manifest file.

To test and run an Outlook add-in, the user's Outlook email account must reside on Exchange 2013 or later, which is available through Office 365, Exchange Online, or through an on-premises installation. The user or administrator installs manifest files for Outlook add-ins on that server.

 >**Note**   POP and IMAP email accounts in Outlook don't support Office Add-ins.


## Client support summary

This table shows the Office host applications (including desktop, tablet, mobile, and web clients) that can run Office Add-ins, and the types of add-ins supported by each host.


**Supported add-in types**


|**Office application**|**Content add-ins**|**Outlook add-ins**|**Task pane add-ins**|
|:-----|:-----|:-----|:-----|
|Access web apps|![Check symbol](../../images/mod_off15_checkmark.png)|||
|Excel 2013 or later|![Check symbol](../../images/mod_off15_checkmark.png)||![Check symbol](../../images/mod_off15_checkmark.png)|
|Excel Online|![Check symbol](../../images/mod_off15_checkmark.png)||![Check symbol](../../images/mod_off15_checkmark.png)|
|Excel for iPad|![Check symbol](../../images/mod_off15_checkmark.png)||![Check symbol](../../images/mod_off15_checkmark.png)|
|Outlook 2013 or later||![Check symbol](../../images/mod_off15_checkmark.png)||
|Outlook for Mac||![Check symbol](../../images/mod_off15_checkmark.png)||
|Outlook Web App||![Check symbol](../../images/mod_off15_checkmark.png)||
|OWA for Devices||![Check symbol](../../images/mod_off15_checkmark.png)||
|PowerPoint 2013 or later|![Check symbol](../../images/mod_off15_checkmark.png)||![Check symbol](../../images/mod_off15_checkmark.png)|
|PowerPoint Online|![Check symbol](../../images/mod_off15_checkmark.png)||![Check symbol](../../images/mod_off15_checkmark.png)|
|Project 2013 or later|||![Check symbol](../../images/mod_off15_checkmark.png)|
|Word 2013 or later|||![Check symbol](../../images/mod_off15_checkmark.png)|
|Word Online|||![Check symbol](../../images/mod_off15_checkmark.png)|
|Word for iPad|||![Check symbol](../../images/mod_off15_checkmark.png)|

## Client requirements: Windows desktop and tablet

The following software is required for developing an Office Add-in for the supported Office desktop clients or web clients that run on Windows-based desktop, laptop, or tablet devices:


- For Windows x86 and x64 desktops, and tablets such as Surface Pro:
    
      - The 32- or 64-bit version of Office 2013 or a later version, running on Windows 7 or a later version.
    
  - Excel 2013, Outlook 2013, PowerPoint 2013, Project Professional 2013, Project 2013 SP1, Word 2013, or a later version of the Office client, if you are testing or running an Office Add-in specifically for one of these Office desktop clients. Office desktop clients can be installed on premises or via Click-to-Run on the client computer.
    
- Internet Explorer 9 or later, which must be installed but doesn't have to be the default browser. To support Office Add-ins, the Office client that acts as host uses browser components that are part of Internet Explorer 9 or later.
    
- One of the following as the default browser: Internet Explorer 9, Safari 5.0.6, Firefox 5, Chrome 13, or a later version of one of these browsers.
    
- An HTML and JavaScript editor such as Notepad, [Visual Studio and the Microsoft Developer Tools](https://www.visualstudio.com/features/office-tools-vs), or a third-party web development tool.
    

## Client requirements: OS X desktop

Outlook for Mac, which is distributed as part of Office 365, supports Outlook add-ins. Running Outlook add-ins on Outlook for Mac has the same requirements as Outlook for Mac itself: the operating system must be at least OS X v10.10 "Yosemite". Because Outlook for Mac uses WebKit as a layout engine to render the add-in pages, there is no additional browser dependency.


## Client requirements: Browser support for Office Online web clients and SharePoint

Any browser that supports ECMAScript 5.1, HTML5, and CSS3, such as Internet Explorer 9, Chrome 13, Firefox 5, Safari 5.0.6, or a later version of these browsers.


## Client requirements: non-Windows smartphone and tablet

Specifically for OWA for Devices, and Outlook Web App running in a browser on smartphones and non-Windows tablet devices, the following software is required for testing and running Outlook add-ins.



|**Host application**|**Device**|**Operating system**|**Exchange account**|**Mobile browser**|
|:-----|:-----|:-----|:-----|:-----|
|OWA for Android|Android smartphones. Technically, those devices considered as "small" or "normal" by [Android OS](http://developer.android.com/guide/practices/screens_support.mdl).|Android 4.4 KitKat or later|On the latest update of Office 365 for business or Exchange Online|Native add-in for Android, browser not applicable|
|OWA for iPad|iPad 2 or later|iOS 6 or later|On the latest update of Office 365 for business or Exchange Online|Native add-in for iOS, browser not applicable|
|OWA for iPhone|iPhone 4S or later|iOS 6 or later|On the latest update of Office 365 for business or Exchange Online|Native add-in for iOS, browser not applicable|
|Outlook Web App|iPhone 4 or later, iPad 2 or later, iPod Touch 4 or later|iOS 5 or later|On Office 365, Exchange Online, or on premise on Exchange Server 2013 or later|Safari|

## Components of an Office Add-in solution


A typical Office Add-in solution involves the following components:


- A client device running the supported Office client - which can be a desktop, laptop, tablet, or smartphone (for Outlook add-ins on OWA for Devices). 
    
- For Access web apps, Word, Excel, PowerPoint, or Project:
    
      - A database, document, workbook, presentation, or project.
    
  - A task pane or content add-in that the user installed from the public Office Store or from a private SharePoint or file-based add-in catalog.
    
- For Outlook: 
    
      - The user's email account and mailbox, which resides on an Exchange Server.
    
  - An Outlook add-in that the user or Exchange Server administrator installed through the Exchange Admin Center (EAC).
    

 >**Note**  The user's installation of an Office Add-in consists of a pointer to the corresponding XML manifest file, which specifies the URL from which to load the add-in webpage and script at run time.

For all supported Office applications, the implementation of the Office Add-in itself consists of the following server-based components:


- An XML manifest file which resides on a public or private add-in catalog, or the user's Exchange Server.
    
- The add-in HTML, CSS, and JavaScript files, which the developer creates and which reside on a web server.
    
- The JavaScript library files, such as JavaScript API for Office (Office.js) and the Microsoft AJAX Library (MicrosoftAjax.js), which Microsoft provides. The add-in accesses the JavaScript library files from content delivery network (CDN) URLs, as specified in its HTML file.
    
- If you are using external JavaScript libraries from a CDN or using web services, ensure you access those resources using Secure Sockets Layer (SSL), otherwise you will get a browser warning when you run your add-in. To use SSL, add the https URL to your resource in the <SCRIPT> tag in your add-in.
    
When a supported Office application starts, it reads the XML manifests for the add-ins that have been installed for or by the user. Subsequently, when a user starts an Office Add-in in the Office application, the following events occur: 


1. For Access web apps, Word, Excel, PowerPoint, or Project: When a user inserts the Office Add-in, or opens an Access web app, document, workbook, presentation, or project that already contains an add-in, the Office application loads the add-in, making its UI visible in the user interface.
    
    For Outlook: Whenever the current Outlook context satisfies the activation conditions of an add-in, Outlook activates the add-in, making the add-in visible in the Outlook UI for selection.
    
2. For Windows or web-based Office applications: The Office application opens the HTML page in a web browser control (desktop or ARM-specific client) or an  **iframe** (web client). The web browser control uses Internet Explorer 9 or later components and provides security and performance isolation.
    
    For OS X-based Outlook for Mac: Outlook for Mac uses a sandboxed WebKit runtime host process to open the HTML page of an Outlook add-in, to help provide similar level of security and performance protection.
    
3. The correspondng browser control,  **iframe**, or WebKit runtime host process loads the HTML body, and calls the event handler for the  **onload** event.
    
4. The Office Add-ins framework calls the event handler for the [initialize](../../reference/shared/office.initialize.md) event of the [Office](../../reference/shared/office.md) object.
    
5. When the HTML body finishes loading and the add-in finishes initializing, the main function of the add-in can proceed.
    

## Additional resources



- [Office Add-ins platform overview](../../docs/overview/office-add-ins.md)
    
