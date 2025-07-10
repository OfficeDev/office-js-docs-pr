---
title: Deploy and publish Office Add-ins
description: Methods and options to deploy your Office Add-in for testing or distribution to users.
ms.date: 06/23/2025
ms.localizationpriority: high
---

# Deploy and publish Office Add-ins

You can use one of several methods to deploy your Office Add-in for testing or distribution to users. The deployment method can also affect which platforms surface your add-in.

> [!NOTE]
> For information about how end users acquire, insert, and run add-ins, see [Start using your Office Add-in](https://support.microsoft.com/office/82e665c4-6700-4b56-a3f3-ef5441996862).

## Primary publication methods

The following table summarizes the primary publication methods that can be used regardless of which type of manifest the add-in uses. If the add-in uses the add-in only manifest, see also [Additional publication methods for the add-in only manifest](#additional-publication-methods-for-the-add-in-only-manifest).

|Method|Use|
|:---------|:------------|
|[Sideloading](../testing/test-debug-office-add-ins.md#sideload-an-office-add-in-for-testing)|As part of your development process, to test your add-in running on Windows, iPad, Mac, or in a browser. (Not for production add-ins.) |
|[AppSource](#appsource)|To distribute your add-in publicly to users.|
|[Microsoft 365 admin center](/microsoft-365/admin/manage/test-and-deploy-microsoft-365-apps)|In a cloud deployment, to distribute your add-in to users in your organization by using the Microsoft 365 admin center. This is done through [Integrated Apps](/microsoft-365/admin/manage/test-and-deploy-microsoft-365-apps) or [Centralized Deployment](/microsoft-365/admin/manage/centralized-deployment-of-add-ins). |

### Production deployment methods

The following sections provide additional information about the deployment methods that are most commonly used to distribute production Office Add-ins to users.

#### AppSource

You can make your add-in available through [AppSource](https://appsource.microsoft.com/marketplace/apps?product=office), Microsoft's online app store which is accessible through a browser and through the UI of Office applications. Distribution through AppSource gives you the option of including installation of your add-in with the installation of your Windows app or a COM or VSTO add-in. For more information, see [Publish to your Office Add-in to AppSource](publish-office-add-ins-to-appsource.md).

[!INCLUDE [publish policies note](../includes/note-publish-policies.md)]

#### Integrated Apps via the Microsoft 365 admin center

The Microsoft 365 admin center makes it easy for an administrator to deploy Office Add-ins to users and groups in their organization. Add-ins deployed via the admin center are available to users in their Office applications right away, with no client configuration required. You can use Integrated Apps to deploy internal add-ins as well as add-ins provided by independent software vendors (ISVs). Integrated Apps also shows admins add-ins and other apps bundled together by same ISV, giving them exposure to the entire experience across the Microsoft 365 platform.

When you link your Office Add-ins, Teams apps, SharePoint Framework (SPFx) apps, and [other apps](/microsoft-365/admin/manage/test-and-deploy-microsoft-365-apps#what-apps-can-i-deploy-from-integrated-apps) together, you create a single software as a service (SaaS) offering for your customers. For general information about this process, see [How to plan a SaaS offer for the commercial marketplace](/azure/marketplace/plan-saas-offer). For specifics on how to create Integrated Apps, see [Configure Microsoft 365 App integration](/azure/marketplace/create-new-saas-offer#configure-microsoft-365-app-integration).

For more information on the Integrated Apps deployment process, see [Test and deploy Microsoft 365 Apps by partners in the Integrated apps portal](/microsoft-365/admin/manage/test-and-deploy-microsoft-365-apps).

> [!NOTE]
> If your add-in uses the [unified manifest for Microsoft 365](../develop/unified-manifest-overview.md) and is distributed with the Microsoft 365 admin center, it won't be installable by users with certain versions of Office. For more information, see [Office Add-ins with the unified app manifest for Microsoft 365 - Client and platform support](../develop/unified-manifest-overview.md#client-and-platform-support).

> [!IMPORTANT]
> Customers in sovereign or government clouds don't have access to Integrated Apps. They use Centralized Deployment instead. Centralized Deployment is a similar deployment method, but doesn't expose connected add-ins and apps to the admin. For more information, see [Determine if Centralized Deployment of add-ins works for your organization](/microsoft-365/admin/manage/centralized-deployment-of-add-ins).

#### Deploy updates

[!INCLUDE [deploy-updates-that-require-admin-consent](../includes/deploy-updates-that-require-admin-consent.md)]

## Additional publication methods for the add-in only manifest

The following table summarizes publication methods that are available *only when the add-in uses the add-in only manifest*.

|Method|Use|Support limitations|
|:---------|:------------|:------------|
|[Network share](../testing/create-a-network-shared-folder-catalog-for-task-pane-and-content-add-ins.md)|As part of your development process, to test your add-in running on Windows computers other than your development computer after you have published the add-in to a server other than localhost.|<ul><li>Not supported for production add-ins.</li><li>Not supported for Outlook add-ins.</li><li>Not supported for testing on iPad, Mac, or the web.</li></ul>|
|[SharePoint catalog](#sharepoint-app-catalog-deployment)|In an on-premises environment, to distribute your add-in to users in your organization.|<ul><li>Not supported for Outlook add-ins.</li><li>Not supported for Office on Mac.</li><li>Not supported for add-ins with any feature that requires a **\<VersionOverrides\>** element in the add-in only manifest.</li></ul>|
|[Exchange server](#outlook-add-in-exchange-server-deployment)|In an on-premises or online environment, to distribute Outlook add-ins to users.|Only supported for Outlook add-ins.|

### SharePoint app catalog deployment

A SharePoint app catalog is a special SharePoint site collection that you can create to host the manifests (add-in only manifest type) of a Word, Excel, or PowerPoint add-in. If you're deploying add-ins in an on-premises environment and none of the add-in users use a Mac, consider using a SharePoint catalog. For details, see [Publish task pane and content add-ins to a SharePoint catalog](publish-task-pane-and-content-add-ins-to-an-add-in-catalog.md).

Because SharePoint catalogs don't support new add-in features implemented in the `VersionOverrides` node of the manifest, including add-in commands, for these add-ins, we recommend that you use Centralized Deployment via the admin center if possible.

### Outlook add-in Exchange server deployment

For on-premises and online environments that don't use the [Microsoft Entra](/entra/fundamentals/what-is-entra) identity service, you can deploy Outlook add-ins via the Exchange server.

Outlook add-in deployment requires:

- Microsoft 365, Exchange Online, or Exchange Server 2016 or later
- Outlook 2016 or later

To assign and manage add-ins for your tenants and users, use [Exchange PowerShell](/powershell/module/exchange). For more information, see [Add-ins for Outlook in Exchange Server](/exchange/add-ins-for-outlook-2013-help) and [Add-ins for Outlook in Exchange Online](/exchange/clients-and-mobile-in-exchange-online/add-ins-for-outlook/add-ins-for-outlook).

It's important to note that some versions of Outlook clients and Exchange servers may only support certain Mailbox requirement sets. For details about supported requirement sets, see [Requirement sets supported by Exchange servers and Outlook clients](/javascript/api/requirement-sets/outlook/outlook-api-requirement-sets#requirement-sets-supported-by-exchange-servers-and-outlook-clients).

## GoDaddy Microsoft 365 SKUs

[Microsoft 365 subscriptions provided by GoDaddy](https://www.godaddy.com/business/office-365) have limited support for add-ins. The following options are **not** supported.

- Deployment through the Microsoft Admin Center.
- Deployment through Exchange servers.
- Acquiring add-ins from AppSource.

## See also

- [Sideload Outlook add-ins for testing](../outlook/sideload-outlook-add-ins-for-testing.md)
- [Publish to your Office Add-in to AppSource](publish-office-add-ins-to-appsource.md)
- [AppSource](https://appsource.microsoft.com/marketplace/apps?product=office)
- [Design guidelines for Office Add-ins](../design/add-in-design.md)
- [Create effective AppSource listings](/partner-center/marketplace-offers/create-effective-office-store-listings)
- [Troubleshoot user errors with Office Add-ins](../testing/testing-and-troubleshooting.md)
- [What is the Microsoft commercial marketplace?](/azure/marketplace/overview)
- [Microsoft Dev Center app publishing page](https://developer.microsoft.com/microsoft-teams/app-publishing)

