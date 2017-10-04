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

The admin center also includes a deployment compatibility checking service.

Future investments in add-in deployment scenarios will focus on the Office 365 admin center. We recommend that you use the admin center to deploy add-ins to your organization, if your organization meets the prerequisites.

## Prerequisites for centralized deployment 

For Word, Excel and PowerPoint 
- Your users must be using Office Professional Plus 2016 on the following operating systems:
  * Win32: build 16.0.8067 or later 
  * Mac: build 15.34.17051500 or later 
  
For Outlook 
- 2013 Click to Run version: 15.0.4819.1000 or later 
- 2013 MSI version: 15.0.4937.1000 or later* 
- 2016 Click to Run version: 16.0.7726.5702 or later 
- 2016 MSI version: 16.0.4494.1000 or later* 

*In MSI version of Outlook, admin-installed add-ins will show in the appropriate ribbon in Outlook but will not show the add-in in 'My add-ins' section 

- Users sign in to Office 2016 with their work or school account.
- Your organization uses the Azure Active Directory (Azure AD) identity service.
- Users' Exchange mailboxes have [OAuth enabled](https://msdn.microsoft.com/en-us/library/office/dn626019(v=exchg.150).aspx#Anchor_0).

Currently, add-ins for the following Office clients are supported. 

| Office application    | Office 2016 for Windows   | Office Online | Office 2016 for Mac   |
|:----------------------|:-------------------------:|:-------------:|:---------------------:|
| Word                  | X                         | X             | X                     |
| Excel                 | X                         | X             | X                     |
| PowerPoint            | X                         | X             | X                     |
| Outlook               | X                         | X             | X                     |

The admin center does not support the following:

- Office 2013 (Word, Excel, PowerPoint, or Outlook)
- Office for iPad
- SharePoint Add-ins
- COM/VSTO based Add-ins
- Office Online Server
- An on-premises directory service

To deploy SharePoint Add-ins or add-ins that target Office 2013, use a [SharePoint add-in catalog](publish-task-pane-and-content-add-ins-to-an-add-in-catalog.md).

>**Important** SharePoint add-in catalogs do not support add-in features that are implemented in the [VersionOverrides](../../reference/manifest/versionoverrides.md) node of the add-in manifest, such as [add-in commands](../design/add-in-commands.md). 

To deploy COM/VSTO add-ins, use ClickOnce or Windows Installer. For details, see [Deploying an Office solution](https://msdn.microsoft.com/en-us/library/bb386179.aspx).

For more information about prerequisites and compatibility checking, see [Determine whether centralized deployment of add-ins works for your Office 365 organization](https://support.office.com/en-us/article/Determine-if-Centralized-Deployment-of-add-ins-works-for-your-Office-365-organization-B4527D49-4073-4B43-8274-31B7A3166F92?ui=en-US&rs=en-US&ad=US).

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
8.	Search for the people or groups to whom you want to deploy the add-in, and choose **Add** next to their name.
    >**Note:** For single sign-on (SSO) add-ins, the users and groups assigned will also be shared with add-ins that share the same Azure App ID. Any changes to user assignments will also apply to those add-ins. The related add-ins will be shown on this page.
9.  For SSO add-ins only: This page will display the list of Microsoft Graph permissions that the add-in requires.
10.	Choose **Save**, review the add-in settings, and then choose **Close**. 
    >**Note:** When an administrator chooses **Save**, consent is given for all users. 


If the add-in supports add-in commands, the commands will appear on the Office application ribbon for all users to whom the add-in is deployed. 

If the add-in does not support add-in commands, users can add it from the **My Add-ins** button by doing the following:

1.	In Word 2016, Excel 2016, or PowerPoint 2016, choose **Insert** > **My Add-ins**.
2.	Choose the **Admin Managed** tab in the add-in window.
3.	Choose the add-in, and then choose **Add**. 

