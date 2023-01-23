---
title: Deploy and publish Office Add-ins
description: Methods and options to deploy your Office Add-in for testing or distribution to users.
ms.date: 01/23/2023
ms.localizationpriority: high
---

# Deploy and publish Office Add-ins

You can use one of several methods to deploy your Office Add-in for testing or distribution to users.

|**Method**|**Use...**|
|:---------|:------------|
|[Sideloading](../testing/test-debug-office-add-ins.md#sideload-an-office-add-in-for-testing)|As part of your development process, to test your add-in running on Windows, iPad, Mac, or in a browser. (Not for production add-ins.)|
|[Network share](../testing/create-a-network-shared-folder-catalog-for-task-pane-and-content-add-ins.md)|As part of your development process, to test your add-in running on Windows after you have published the add-in to a server other than localhost. (Not for production add-ins or for testing on iPad, Mac, or the web.)|
|[AppSource][AppSource]|To distribute your add-in publicly to users.|
|[Microsoft 365 admin center](/microsoft-365/admin/manage/test-and-deploy-microsoft-365-apps)|In a cloud deployment, to distribute your add-in to users in your organization by using the Microsoft 365 admin center. This is done through [Integrated Apps](/microsoft-365/admin/manage/test-and-deploy-microsoft-365-apps) or [Centralized Deployment](/microsoft-365/admin/manage/centralized-deployment-of-add-ins). |
|[SharePoint catalog](publish-task-pane-and-content-add-ins-to-an-add-in-catalog.md)|In an on-premises environment, to distribute your add-in to users in your organization.|
|[Exchange server](#outlook-add-in-deployment)|In an on-premises or online environment, to distribute Outlook add-ins to users.|

[!INCLUDE [publish policies note](../includes/note-publish-policies.md)]

## Deployment options by Office application and add-in type

The deployment options that are available depend on the Office application that you're targeting and the type of add-in you create.

### Deployment options for Word, Excel, and PowerPoint add-ins

| Extension point | Sideloading | Network share | AppSource | Microsoft 365 admin center | SharePoint catalog\* |
|:----------------|:-----------:|:-------------:|:---------:|:--------------------------:|:--------------------:|
| Content         | Supported   | Supported     | Supported | Supported                  | Supported            |
| Task pane       | Supported   | Supported     | Supported | Supported                  | Supported            |
| Command         | Supported   | Supported     | Supported | Supported                  | Not available        |

&#42; SharePoint catalogs do not support Office on Mac.

### Deployment options for Outlook add-ins

| Extension point | Sideloading | AppSource | Exchange server |
|:----------------|:-----------:|:---------:|:---------------:|
| Mail app        | Supported   | Supported | Supported       |
| Command         | Supported   | Supported | Supported       |

## Production deployment methods

The following sections provide additional information about the deployment methods that are most commonly used to distribute production Office Add-ins to users within an organization.

> [!IMPORTANT]
> If you deploy updates to an already deployed add-in, some changes in the manifest require admin consent. Users are unable to use the updated add-in until consent is granted. The following changes in the manifest will require the admin to consent to them.
> - Changes to requested [permissions](/javascript/api/manifest/permissions).
> - Adding new [scopes](/javascript/api/manifest/scopes).
> - Adding new [Outlook events](../outlook/autolaunch.md).

For information about how end users acquire, insert, and run add-ins, see [Start using your Office Add-in](https://support.microsoft.com/office/82e665c4-6700-4b56-a3f3-ef5441996862).

### Integrated Apps via the Microsoft 365 admin center

The Microsoft 365 admin center makes it easy for an administrator to deploy Office Add-ins to users and groups in their organization. Add-ins deployed via the admin center are available to users in their Office applications right away, with no client configuration required. You can use Integrated Apps to deploy internal add-ins as well as add-ins provided by ISVs. Integrated Apps also shows admins add-ins and other apps bundled together by same ISV, giving them exposure to the entire experience across the Microsoft 365 platform.

When you link your Office Add-ins, Teams apps, SPFx apps, and [other apps](/microsoft-365/admin/manage/test-and-deploy-microsoft-365-apps#what-apps-can-i-deploy-from-integrated-apps) together, you create a single software as a service (SaaS) offering for your customers. For general information about this process, see [How to plan a SaaS offer for the commercial marketplace](/azure/marketplace/plan-saas-offer). For specifics on how to create Integrated Apps, see [Configure Microsoft 365 App integration](/azure/marketplace/create-new-saas-offer#configure-microsoft-365-app-integration).

For more information on the Integrated Apps deployment process, see [Test and deploy Microsoft 365 Apps by partners in the Integrated apps portal](/microsoft-365/admin/manage/test-and-deploy-microsoft-365-apps).

> [!IMPORTANT]
> Customers in sovereign or government clouds don't have access to Integrated Apps. They will use Centralized Deployment instead. Centralized Deployment is a similar deploy method, but doesn't expose connected add-ins and apps to the admin. For more information, see [Determine if Centralized Deployment of add-ins works for your organization](/microsoft-365/admin/manage/centralized-deployment-of-add-ins).

### SharePoint app catalog deployment

A SharePoint app catalog is a special site collection that you can create to host Word, Excel, and PowerPoint add-ins. Because SharePoint catalogs don't support new add-in features implemented in the `VersionOverrides` node of the manifest, including add-in commands, we recommend that you use Centralized Deployment via the admin center if possible. Add-in commands deployed via a SharePoint catalog open in a task pane by default.

If you are deploying add-ins in an on-premises environment, use a SharePoint catalog. For details, see [Publish task pane and content add-ins to a SharePoint catalog](publish-task-pane-and-content-add-ins-to-an-add-in-catalog.md).

> [!NOTE]
> SharePoint catalogs do not support Office on Mac. To deploy Office Add-ins to Mac clients, you must submit them to [AppSource](/office/dev/store/submit-to-the-office-store).

### Outlook add-in deployment

For on-premises and online environments that do not use the Azure AD identity service, you can deploy Outlook add-ins via the Exchange server.

Outlook add-in deployment requires:

- Microsoft 365, Exchange Online, or Exchange Server 2013 or later
- Outlook 2013 or later

To assign add-ins to tenants, use the Exchange admin center to upload a manifest directly, either from a file or a URL, or add an add-in from AppSource. To assign add-ins to individual users, you must use Exchange PowerShell. For details, see [Add-ins for Outlook in Exchange Server](/exchange/add-ins-for-outlook-2013-help).

## See also

- [Sideload Outlook add-ins for testing](../testing/create-a-network-shared-folder-catalog-for-task-pane-and-content-add-ins.md)
- [Submit to AppSource][AppSource]
- [AppSource](https://appsource.microsoft.com/marketplace/apps?src=office&page=1)
- [Design guidelines for Office Add-ins](../design/add-in-design.md)
- [Create effective AppSource listings](/office/dev/store/create-effective-office-store-listings)
- [Troubleshoot user errors with Office Add-ins](../testing/testing-and-troubleshooting.md)
- [What is the Microsoft commercial marketplace?](/azure/marketplace/overview)
- [Microsoft Dev Center app publishing page](https://developer.microsoft.com/microsoft-teams/app-publishing)

[AppSource]: /office/dev/store/submit-to-appsource-via-partner-center
