---
title: Deploy and publish your Office Add-in
description: ''
ms.date: 01/23/2018
---

# Deploy and publish your Office Add-in

You can use one of several methods to deploy your Office Add-in for testing or distribution to users.

|**Method**|**Use...**|
|:---------|:------------|
|[Sideloading](../testing/create-a-network-shared-folder-catalog-for-task-pane-and-content-add-ins.md)|As part of your development process, to test your add-in running on Windows, Office Online, iPad, or Mac.|
|[Centralized Deployment](centralized-deployment.md)|In a cloud or hybrid deployment, to distribute your add-in to users in your organization by using the Office 365 admin center.|
|[SharePoint catalog](publish-task-pane-and-content-add-ins-to-an-add-in-catalog.md)|In an on-premises environment, to distribute your add-in to users in your organization.|
|[AppSource](https://docs.microsoft.com/office/dev/store/submit-to-the-office-store)|To distribute your add-in publicly to users.|
|[Exchange server](#outlook-add-in-deployment)|In an on-premises or online environment, to distribute Outlook add-ins to users.|
|[Network share](../testing/create-a-network-shared-folder-catalog-for-task-pane-and-content-add-ins.md)|On a Windows computer on a network where you want to host your add-in, go to the parent folder, or drive letter, of the folder you want to use as your shared folder catalog.|

> [!NOTE]
> If you plan to [publish](../publish/publish.md) your add-in to AppSource and make it available within the Office experience, make sure that you conform to the [AppSource validation policies](https://docs.microsoft.com/office/dev/store/validation-policies). For example, to pass validation, your add-in must work across all platforms that support the methods that you define (for more information, see [section 4.12](https://docs.microsoft.com/office/dev/store/validation-policies#4-apps-and-add-ins-behave-predictably) and the [Office Add-in host and availability page](../overview/office-add-in-availability.md)).

## Deployment options by Office host

The deployment options that are available depend on the Office host that you're targeting and the type of add-in you create.

### Deployment options for Word, Excel, and PowerPoint add-ins

| Extension point | Sideloading | Office 365 admin center |AppSource   | SharePoint catalog\* |
|:----------------|:-----------:|:-----------------------:|:----------:|:--------------------:|
| Content         | X           | X                       | X          | X                    |
| Task pane       | X           | X                       | X          | X                    |
| Command 		  | X           | X                       | X          |                      |

&#42; SharePoint catalogs do not support Office for Mac.

### Deployment options for Outlook add-ins

| Extension point | Sideloading | Exchange server | AppSource    |
|:----------------|:-----------:|:---------------:|:------------:|
| Mail app        | X           | X               | X            |
| Command         | X           | X               | X            |

## Deployment methods

The following sections provide additional information about the deployment methods that are most commonly used to distribute Office Add-ins to users within an organization.

For information about how end users acquire, insert, and run add-ins, see [Start using your Office Add-in](https://support.office.com/en-ie/article/Start-using-your-Office-Add-in-82e665c4-6700-4b56-a3f3-ef5441996862?ui=en-US&rs=en-IE&ad=IE).

### Centralized Deployment via the Office 365 admin center 

The Office 365 admin center makes it easy for an administrator to deploy Office Add-ins to users and groups in their organization. Add-ins deployed via the admin center are available to users in their Office applications right away, with no client configuration required. You can use Centralized Deployment to deploy internal add-ins as well as add-ins provided by ISVs.

For more information, see [Publish Office Add-ins using Centralized Deployment via the Office 365 admin center](centralized-deployment.md).

### SharePoint catalog deployment

A SharePoint add-in catalog is a special site collection that you can create to host Word, Excel, and PowerPoint add-ins. Because SharePoint catalogs don't support new add-in features implemented in the `VersionOverrides` node of the manifest, including add-in commands, we recommend that you use Centralized Deployment via the admin center if possible. Add-in commands deployed via a SharePoint catalog open in a task pane by default.

If you are deploying add-ins in an on-premises environment, use a SharePoint catalog. For details, see [Publish task pane and content add-ins to a SharePoint catalog](publish-task-pane-and-content-add-ins-to-an-add-in-catalog.md).

> [!NOTE]
> SharePoint catalogs do not support Office for Mac. To deploy Office Add-ins to Mac clients, you must submit them to [AppSource](https://docs.microsoft.com/office/dev/store/submit-to-the-office-store). 

### Outlook add-in deployment

For on-premises and online environments that do not use the Azure AD identity service, you can deploy Outlook add-ins via the Exchange server. 

Outlook add-in deployment requires:

- Office 365, Exchange Online, or Exchange Server 2013 or later
- Outlook 2013 or later

To assign add-ins to tenants, you use the Exchange admin center to upload a manifest directly, either from a file or a URL, or add an add-in from AppSource. To assign add-ins to individual users, you must use Exchange PowerShell. For details, see [Install or remove Outlook add-ins for your organization](https://technet.microsoft.com/library/jj943752(v=exchg.150).aspx) on TechNet.

## See also

- [Sideload Outlook add-ins for testing](../testing/create-a-network-shared-folder-catalog-for-task-pane-and-content-add-ins.md)
- [Submit to AppSource][AppSource]
- [Design guidelines for Office Add-ins](../design/add-in-design.md)
- [Create effective AppSource listings](https://docs.microsoft.com/office/dev/store/create-effective-office-store-listings)
- [Troubleshoot user errors with Office Add-ins](../testing/testing-and-troubleshooting.md)

[AppSource]: https://docs.microsoft.com/office/dev/store/submit-to-the-office-store
[Office Add-in host and platform availability]: ../overview/office-add-in-availability
