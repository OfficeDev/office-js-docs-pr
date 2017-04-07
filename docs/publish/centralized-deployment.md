# Use centralized deployment to publish Office Add-ins

The Office 365 admin center makes it easy for an administrator to deploy Word, Excel, and PowerPoint add-ins to users or groups within their organization. Add-ins deployed via the admin center are available to users in their Office applications right away, with no client configuration required. You can deploy internal add-ins as well as add-ins provided by ISVs via centralized deployment.

The admin center currently supports the following scenarios:

- Centralized deployment of new and updated add-ins to individuals, groups, or an organization.
- Deployment to multiple platforms, including Windows and Office Online, with Mac coming soon.
- Deployment to English language and worldwide tenants.
- Cloud-hosted add-in deployment.
- Automatic installation on launch of the Office application.
- Add-in URLs hosted within a firewall.
- Deployment of Office Store add-ins (coming soon).

<!--
The admin center also includes a pre-deployment validation checking service.
-->

Future investments in add-in deployment scenarios will focus on the Office 365 admin center. We recommend that you use the admin center to deploy add-ins to your organization, if your organization meets the prerequisites.

## Prerequisites for centralized deployment 

You can deploy add-ins via the admin center if your organization meets the following criteria:

- Users are running a version of Office 2016 ProPlus:
    - Windows build 16.0.8027 or later
    - Mac build 15.33.170327 or later
- Users sign in to Office 2016 with their work or school account.
- Your organization uses the Azure Active Directory (Azure AD) identity service.
- Users' Exchange mailboxes have [OAuth enabled](https://msdn.microsoft.com/en-us/library/office/dn626019(v=exchg.150).aspx#Anchor_0).

Currently, add-ins for the following Office clients are supported. 

|**Office application**|**Office 2016 for Windows**|**Office Online**|**Office 2016 for Mac**|
|:---------------------|:--------------------------|:--------------|:------------------|
|Word|X|X|X|
|Excel|X|X|X|
|PowerPoint|X|X|X|
|Outlook|Coming soon|Coming soon|Coming soon|

The admin center does not support the followng:

- Add-ins that target Word, Excel, PowerPoint, or Outlook in Office 2013.
- An on-premises directory service.
- SharePoint Add-in deployment.
- Add-in deployment to Office Online Server.
- Deployment of COM/VSTO add-ins.

To deploy SharePoint Add-ins or add-ins that target Office 2013, use a [SharePoint add-in catalog](publish-task-pane-and-content-add-ins-to-an-add-in-catalog.md).

>**Important!** SharePoint add-in catalogs do not support add-in features that are implemented in the [VersionOverrides](../../reference/manifest/versionoverrides.md) node of the add-in manifest, such as [add-in commands](../design/add-in-commands.md). 

To deploy COM/VSTO add-ins, use ClickOnce or Windows Installer. For details, see [Deploying an Office solution](https://msdn.microsoft.com/en-us/library/bb386179.aspx).

<!-- Need URL on SOC site.
For more information about requirements, see [centralized deployment eligibility]().
-->

## Publish an add-in via centralized deployment

To publish an add-in via centralized deployment:

1.	Verify that your organization meets the [prerequisites for centralized deployment](#prerequisites-for-centralized-deployment).
2.	On the Office 365 admin center page, choose **Settings** > **Services & add-ins**.
3.	Choose **Add an Office Add-in** at the top of the page. You have the following options:

    - Add an add-in from the Office Store.
    - Choose **Browse** to locate your manifest (.xml) file.
    - Enter a URL for your manifest in the field provided.

5.	Choose **Next**.
6.	If you're adding an add-in from the Office Store, select the add-in. The add-in is now enabled. 
7.	Choose **Edit** to assign the add-in to users. 
8.	Search for the people or groups to whom you want to deploy the add-in and choose **Add** next to their name.
9.	Choose **Save**, review the add-in settings, and then choose **Close**.


If the add-in supports add-in commands, the commands will appear on the Office application ribbon for all users to whom the add-in is deployed. 

If the add-in does not support add-in commands, users can add it from the **My Add-ins** button by doing the following:

1.	In Word 2016, Excel 2016, or PowerPoint 2016, choose **Insert** > **My Add-ins**.
2.	Choose the **Admin Managed** tab in the add-in window.
3.	Choose the add-in, and then choose **Add**. 

