
# Deploy and publish your Office Add-in

You can use one of several methods to deploy your Office Add-in for testing or distribution to users.

|**Method**|**Use...**|
|:---------|:------------|
|[Sideloading](../testing/create-a-network-shared-folder-catalog-for-task-pane-and-content-add-ins.md)|As part of your development process, to test your add-in running on Windows, Office Online, iPad, or Mac.|
|[Office 365 admin center (preview)](#office-365-admin-center-preview)|in a cloud or hybrid deployment, to distribute your add-in to users in your organization.|
|[Office Store]|To distribute your add-in publicly to users.|
|[SharePoint catalog](publish-task-pane-and-content-add-ins-to-an-add-in-catalog.md)|In an on-premises environment, to distribute your add-in to users in your organization.|
|[Exchange server](#outlook-add-in-deployment)|In an on-premises or online environment, to distribute Outlook add-ins to users.|

The options that are available depend on the Office host that you're targeting and the type of add-in you create.

>**Note:** If you plan to your add-in to the Office Store, make sure that you conform to the [Office Store validation policies](https://msdn.microsoft.com/en-us/library/jj220035.aspx). For example, to pass validation, your add-in must work across all platforms that support the methods that you define (for more information, see [section 4.12](https://dev.office.com/officestore/docs/validation-policies#4-apps-and-add-ins-behave-predictably) and the [Office Add-in host and availability page](https://dev.office.com/add-in-availability)).

## Deployment options for Word, Excel, and PowerPoint add-ins

| Extension point            | Sideloading | Office 365 admin center (preview) |Office Store| SharePoint catalog*  |
|:----------------|:-----------:|:------------------:|:-------------------------------:|:------------:|
| Content         | X           | X                  | X                               | X|
| Task pane       | X           | X                  | X                               | X|
| Command 		  | X           | X                  | X                               |  |

&#42; SharePoint catalogs do not support Office 2016 for Mac.

## Deployment options for Outlook add-ins

| Extension point     | Sideloading | Exchange server | Office Store |
|:---------|:-----------:|:---------------:|:------------:|
| Mail app | X           | X               | X            |
| Command  | X           | X               | X            |


For information about how end users acquire, insert, and run add-ins, see [Start using your Office Add-in](https://support.office.com/en-ie/article/Start-using-your-Office-Add-in-82e665c4-6700-4b56-a3f3-ef5441996862?ui=en-US&rs=en-IE&ad=IE).

## Office 365 admin center (preview) deployment

The Office 365 admin center makes it easy for an admin to deploy Word, Excel, and PowerPoint add-ins to users or groups within their organization. Add-ins deployed via the admin center are available to users in their Office applications right away, with no client configuration required. You can deploy internal add-ins as well as add-ins provided by ISVs via the admin center.

<!--
The admin center currently supports the following scenarios:

- Centralized deployment of new and updated add-ins to individuals, groups, or an organization.
- Support for multiple platforms, including Windows and Office Online, with Mac coming soon.
- Deployment to English language and worldwide tenants.
- Cloud-hosted add-in deployment.
- Automatic installation on launch of the Office application.
- Add-in URLs hosted within a firewall.
- Deployment of Office Store add-ins (coming soon).

Future investments in add-in deployment scenarios will focus on the Office 365 admin center. We recommend that you use the admin center to deploy add-ins to your organization, if your organization meets the prerequisites.

### Prerequisites for admin center deployment 

You can deploy add-ins via the admin center if your organization meets the following criteria:

- Users are running Office 2016 build 7070 or later.
- Users sign in to Office 2016 with their work or school account.
- Your organization uses the Azure Active Directory (Azure AD) identity service.

The admin center does not support the followng:

- Add-ins that target Word, Excel, or PowerPoint in Office 2013.
- An on-premises directory service.
- SharePoint Add-in deployment.
- Add-in deployment to Office Online Server.
- Deployment of COM/VSTO add-ins.

To deploy SharePoint Add-ins or add-ins that target Office 2013, use a [SharePoint add-in catalog](publish-task-pane-and-content-add-ins-to-an-add-in-catalog.md).

>**Important!** SharePoint add-in catalogs do not support add-in features that are implemented in the [VersionOverrides](../../reference/manifest/versionoverrides.md) node of the add-in manifest, such as [add-in commands](../design/add-in-commands.md). 

To deploy COM/VSTO add-ins, use ClickOnce or Windows Installer. For details, see [Deploying an Office solution](https://msdn.microsoft.com/en-us/library/bb386179.aspx). -->

## SharePoint catalog deployment

A SharePoint add-in catalog is a special site collection that you can create to host Word, Excel, and PowerPoint add-ins. Because SharePoint catalogs don't support new add-in features implemented in the VersionOverrides node of the manifest, including add-in commands, we recommend that you use centralized deployment via the admin center (preview) if possible. Add-in commands deployed via a SharePoint catalog open in a task pane by default.

If you are deploying add-ins in an on-premises environment, use a SharePoint catalog. For details, see [Publish task pane and content add-ins to a SharePoint catalog](publish-task-pane-and-content-add-ins-to-an-add-in-catalog.md).

> **Note:** SharePoint catalogs do not support Office 2016 for Mac. To deploy Office Add-ins to Mac clients, you must submit them to the [Office Store]. 

## Outlook add-in deployment

For on-premises and online environments that do not use the Azure AD identity service, you can deploy Outlook add-ins via the Exchange server. 

Outlook add-in deployment requires:

- Office 365, Exchange Online, or Exchange Server 2013 or later
- Outlook 2013 or later

To assign add-ins to tenants, you use the Exchange admin center to upload a manifest directly, either from a file or a URL, or add an add-in from the Office Store. To assign add-ins to individual users, you must use Exchange PowerShell. For details, see [Install or remove Outlook add-ins for your organization](https://technet.microsoft.com/en-us/library/jj943752(v=exchg.150).aspx) on TechNet.


## Additional resources

- [Deploy and install Outlook add-ins for testing](../outlook/testing-and-tips.md) 
- [Submit add-ins and web apps to the Office Store][Office Store]
- [Design guidelines for Office Add-ins](../design/add-in-design)
- [Create effective Office Store add-ins](https://msdn.microsoft.com/en-us/library/jj635874.aspx)
- [Troubleshoot user errors with Office Add-ins](../testing/testing-and-troubleshooting.md)

[Office Store]: http://msdn.microsoft.com/library/ff075782-1303-4517-91cc-b3d730e9b9ae%28Office.15%29.aspx
[Office Add-in host and platform availability]: http://dev.office.com/add-in-availability
