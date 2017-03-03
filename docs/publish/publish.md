
# Deploy and publish your Office Add-in

You can use one of several methods to deploy your Office Add-in for testing or distribution to users: 

- [Sideloading](../testing/create-a-network-shared-folder-catalog-for-task-pane-and-content-add-ins.md) - Use as part of your development process to test your add-in running on Windows, Office Online, iPad, or Mac.
- [Office 365 admin center (preview)](#office-365-admin-center-preview) - Use to distribute your add-in to users in your organization in a cloud or hybrid deployment.
- [Office Store] - Use to distribute your add-in publicly to users.
- [SharePoint catalog](publish-task-pane-and-content-add-ins-to-an-add-in-catalog.md) - Use as part of your development process to test your add-in, or, in an on-premises environment, to distribute your add-in to users in your organization.

The options that are available depend on the Office host that you're targeting and the type of add-in you create.

>**Note:** When you build your add-in, if you plan to [publish](../publish/publish.md) your add-in to the Office Store, make sure that you conform to the [Office Store validation policies](https://msdn.microsoft.com/en-us/library/jj220035.aspx). For example, to pass validation, your add-in must work across all platforms that support the methods that you define (for more information, see [section 4.12](https://msdn.microsoft.com/en-us/library/jj220035.aspx#Anchor_3) and the [Office Add-in host and availability page](https://dev.office.com/add-in-availability)).

**Deployment options for Word, Excel, and PowerPoint add-ins**

| Extension point            | Sideloading | Office 365 admin center (preview) |Office Store  | SharePoint catalog  |
|:----------------|:-----------:|:------------------:|:-------------------------------:|:------------:|
| Content         | X           | X                  | X                               | X            |
| Task pane       | X           | X                  | X                               | X            |
| Command 		  | X           | X                  | X                                |              |

> **Note:** SharePoint catalogs are not supported for Office 2016 for Mac. To deploy Office Add-ins to Mac clients, you must submit them to the [Office Store].    

**Deployment options for Outlook add-ins**

| Extension point     | Sideloading | Exchange server | Office Store |
|:---------|:-----------:|:---------------:|:------------:|
| Mail app | X           | X               | X            |
| Command  | X           | X               | X            |

To broaden the reach of your add-in, make sure that it works across platforms. Office Add-ins are supported on Windows, Mac, Web, iOS and Android. For an overview of which features are supported by each platform, see [Office Add-in host and platform availability].   

For information about licensing your Office Store add-ins, see [License your add-ins](https://msdn.microsoft.com/EN-US/library/office/jj163257.aspx).

For information about how end users acquire, insert, and run add-ins, see [Start using your Office Add-in](https://support.office.com/en-ie/article/Start-using-your-Office-Add-in-82e665c4-6700-4b56-a3f3-ef5441996862?ui=en-US&rs=en-IE&ad=IE).

## Office 365 admin center (preview)

The Office 365 admin center (preview) makes it easy for an admin to deploy Office Add-ins to users or groups within their organization. Add-ins deployed via the admin center are available to users in their Office applications right away, with no client configuration required.

The admin center (preview) currently supports:

- Centralized deployment of add-ins and add-in update to individuals, groups, or an organization.
- Multiple platforms, including Windows, Office Online, and Mac (coming soon).
- Word, Excel, and PowerPoint add-in deployment.
- Deployment to worldwide tenants.
- Deployment of internal add-ins and add-ins provided by ISVs.
- Cloud-hosted Office Add-ins.
- A pre-deployment validation checking service.
- Automatic installation of Word, Excel, and PowerPoint add-ins on launch of the application.
- Add-in URLs hosted within your firewall.
- Deployment of Office Store add-ins (coming soon).

Future investments in add-in deployment scenarios will focus on the Office 365 admin center. We recommend that you use the admin center to deploy add-ins to your organization, if your organization meets the criteria.

### Criteria for admin center deployment 

You can use centralized deployment if your organization meets the following criteria:

- Users are running Office 2016 build 7070 or later.
- Users sign in to Office 2016 with their work or school account.
- Your organization uses the Azure Active Directory (Azure AD) identity service.

Centralized deployment does not support the followng scenarios:

- Add-ins that target Word, Excel, or PowerPoint in Office 2013.
- An on-premises directory service.
- SharePoint Add-in deployment.
- Add-in deployment to Office Online Server.
- Deployment to Mac platforms (iOS, iPad).
- Deployment of COM/VSTO add-ins.

To deploy Office 2013 and SharePoint Add-ins, use a [SharePoint add-in catalog](publish-task-pane-and-content-add-ins-to-an-add-in-catalog.md).

>**Important!** SharePoint add-in catalogs do not support add-in features that are implemented in the [VersionOverrides](../../reference/manifest/versionoverrides.md) node of the add-in manifest, such as [add-in commands](../design/add-in-commands.md). 

To deploy COM/VSTO add-ins, see [Deploying an Office solution](https://msdn.microsoft.com/en-us/library/bb386179.aspx)

## SharePoint catalog

A SharePoint add-in catalog is a special site collection that you can create to host Word, Excel, and PowerPoint add-ins. Add-in commands deployed via a SharePoint catalog default to task pane add-ins.

You can deploy Word, Excel, and PowerPoint add-ins that target Office 2013 or an on-premises environment via a SharePoint catalog. For details, see [Publish task pane and content add-ins to a SharePoint catalog](publish-task-pane-and-content-add-ins-to-an-add-in-catalog.md).

>**Note:** Because SharePoint catalogs don't support new add-in features implemented in the VersionOverrides node of the manifest, we recommend that you use centralized deployment via the admin center preview if possible.

## Outlook add-in deployment

For on-premises and online environments that do not use the Azure AD identity service, you can deploy Outlook add-ins via the Exchange server. 

To assign add-ins to whole tenants, you can upload an add-in via the Exchange admin center directly from the manifest, or from the Office Store. To assign add-ins to individual users, you must use Exchange PowerShell. For details, see [Install or remove Outlook add-ins for your organization](https://technet.microsoft.com/en-us/library/jj943752(v=exchg.150).aspx) on TechNet.

Outlook add-in deployment required Office 365, Exchange Online, Exchange Server 2013 or later, and Outlook 2013 or later.

## Additional resources

- [Office Add-in host and platform availability]
- [Deploy and install Outlook add-ins for testing](../outlook/testing-and-tips.md) 
- [Submit add-ins and web apps to the Office Store][Office Store]
- [Design guidelines for Office Add-ins](../design/add-in-design)
- [Create effective Office Store add-ins](https://msdn.microsoft.com/en-us/library/jj635874.aspx)
- [Troubleshoot user errors with Office Add-ins](../testing/testing-and-troubleshooting.md)

[Office Store]: http://msdn.microsoft.com/library/ff075782-1303-4517-91cc-b3d730e9b9ae%28Office.15%29.aspx
[Office Add-in host and platform availability]: http://dev.office.com/add-in-availability
