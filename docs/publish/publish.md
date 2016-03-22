
# Publish your Office Add-in


You can publish your add-ins to the Office Store or upload them to a private shared folder add-in catalog on SharePoint, a shared network folder, or an Exchange server. The options that are available depend on the type of add-in you create. 

For information about publishing to the Office Store, see [Submit add-ins and web apps to the Office Store](http://msdn.microsoft.com/library/ff075782-1303-4517-91cc-b3d730e9b9ae%28Office.15%29.aspx). 


**Options for publishing Office Add-ins**


|**Type**|**Office Store**|**Corporate add-in catalog**|**Shared folder add-in catalog**|**Exchange server**|
|:-----|:-----|:-----|:-----|:-----|
|Task pane add-in|x|x|x||
|Content add-in|x|x|x||
|Outlook add-in|x|||x|
Before you publish your add-in, you'll need to [package it](../publish/package-your-add-in-using-napa-or-visual-studio.md). In addition to making your add-ins available to end users, you'll want to consider how you can broaden your add-in's reach.


## Publishing task pane and content add-ins to an add-in catalog


For task pane and content add-ins, IT departments can deploy and configure private corporate add-in catalogs to provide the same Office-solution catalog experience that the Office Store provides. This new catalog and development platform lets IT use a streamlined method to provision Office and SharePoint Add-ins to managed users from a central location without the need to deploy solutions to each client. You can then use the telemetry tool to monitor add-in usage, verify compatibility, and troubleshoot end user issues. To learn more, see: 


- [Publish task pane and content add-ins to an add-in catalog on SharePoint](../publish/publish-task-pane-and-content-add-ins-to-an-add-in-catalog.md)
    
- [Set up an add-in catalog on SharePoint](../publish/set-up-an-add-in-catalog-on-sharepoint.md)
    
- [Set up an add-in catalog on Office 365](../publish/set-up-an-add-in-catalog-on-office-365.md)
    

## Publishing task pane and content add-ins to a shared network folder


Alternatively, in a corporate setting, IT can deploy task pane and content add-ins created either by in-house or third-party developers to a shared network folder, where the manifest files will be stored and managed. In either case, when developers update their add-ins, they don't have to push updates to end users or IT does not have to redeploy them to corporate users. For information about setting up a shared network folder add-in catalog, see [Create a network shared folder catalog for task pane and content add-ins](../publish/create-a-network-shared-folder-catalog-for-task-pane-and-content-add-ins.md).


## Publishing Outlook add-ins to an Exchange server


Outlook add-ins are published to an Exchange catalog that is available to users of the Exchange server on which it resides. It enables publishing and management of Outlook add-ins, including internally created ones as well as solutions that are acquired from the Office Store and licensed for corporate use. Outlook add-ins are installed into an Exchange catalog by using either the Exchange Admin Center (EAC) or by running remote Windows PowerShell commands (cmdlets). For information about publishing an Outlook add-in, see [Deploy and install Outlook add-ins for testing](../outlook/testing-and-tips.md).


## Add-in experience for end users


Add-ins are easy for end users to acquire, insert, and run. Users have the same experience regardless of whether they access add-ins from any of the following:


- The public Office Store, by using their Microsoft account.
    
- A SharePoint add-in catalog, by using their company ID.
    
- A shared network folder.
    
- An Exchange server.
    
To acquire a new task pane add-in in Excel, for example, users log on to Office with their Microsoft account, open an Excel workbook, and select  **My Add-ins** on the **Insert** tab of the ribbon. The **Office Add-ins** dialog box appears.

In the  **Office Add-ins** dialog box, the user chooses **Find more add-ins at the Office Store**. After users log on to Office.com, using the same Microsoft account, they can download the add-in of their choice and pay for it with a credit card.

In Excel, in the  **Office Add-ins** dialog box, the user chooses **Refresh**, selects the add-in they downloaded, and then chooses  **Insert**.

When they sign in to their account, they have access to their add-ins from any computer, anywhere, including those running Office 365.


## Broaden the reach for your add-in


To ensure that your add-in reaches more end users, make sure that it works across platforms. The Office.js version 1.1 includes support for Office Online, and the Office Store validation process verifies add-in support for Office Online. Before you publish, test your add-in to make sure that it works in Office Online.

To make your add-in available in the Office Store, see [Submit add-ins and web apps to the Office Store](http://msdn.microsoft.com/library/ff075782-1303-4517-91cc-b3d730e9b9ae%28Office.15%29.aspx).

Office Add-ins are also supported on Office for iPad. Make sure to test your add-in on the iPad before you [ submit it to the Seller Dashboard](http://msdn.microsoft.com/library/260ef238-0be4-42d6-ba15-1249a8e2ff12%28Office.15%29.aspx). When you have verified that your add-in works as expected, you can mark your submission as iOS-compatible in the Seller Dashboard. For validation, you will need to provide your Apple developer ID. See also [Debug Office Add-ins on iPad and Mac](../testing/debug-office-add-ins-on-ipad-and-mac.md).

To address user issues with your add-ins, see [Troubleshoot user errors with Office Add-ins](../testing/testing-and-troubleshooting.md). You can also [respond directly to customer reviews in the Office Store](https://msdn.microsoft.com/library/jj635874.aspx).




## Additional resources



- [Office Add-ins](../../docs/overview/office-add-ins.md)
    
- [Submit Office and SharePoint Add-ins and Office 365 web apps to the Office Store](http://msdn.microsoft.com/library/ff075782-1303-4517-91cc-b3d730e9b9ae%28Office.15%29.aspx)
    


