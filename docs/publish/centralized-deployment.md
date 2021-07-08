---
title: Publish Office Add-ins using Centralized Deployment via the Microsoft 365 admin center
description: 'Learn how to use Centralized Deployment to deploy internal add-ins as well as add-ins provided by ISVs.'
ms.date: 03/22/2021
localization_priority: Normal
---


# Publish Office Add-ins using Centralized Deployment via the Microsoft 365 admin center

The Microsoft 365 admin center makes it easy for an administrator to deploy Office Add-ins to users and groups within their organization. Add-ins deployed via the admin center are available to users in their Office applications right away, with no client configuration required. You can use Centralized Deployment to deploy internal add-ins as well as add-ins provided by ISVs.

The Microsoft 365 admin center currently supports the following scenarios.

- Centralized Deployment of new and updated add-ins to individuals, groups, or an organization.
- Deployment to multiple client platforms, including Windows, Mac, and the web. For Outlook, deployment to iOS and Android is also supported. (However, while user installation of Excel, Outlook, Word, and PowerPoint add-ins on iPad is supported, Centralized Deployment to iPad is **not** supported.)
- Deployment to English language and worldwide tenants.
- Deployment of cloud-hosted add-ins.
- Deployment of add-ins that are hosted within a firewall.
- Deployment of AppSource add-ins.
- Automatic installation of an add-in for users when they launch the Office application.
- Automatic removal of an add-in for users if the admin turns off or deletes the add-in, or if users are removed from Azure Active Directory or from a group to which the add-in has been deployed.

Centralized Deployment is the recommended way for a Microsoft 365 admin to deploy Office Add-ins within an organization, provided that the organization meets all requirements for using Centralized Deployment. For information about how to determine if your organization can use Centralized Deployment, see [Determine if Centralized Deployment of add-ins works for your Microsoft 365 organization](/office365/admin/manage/centralized-deployment-of-add-ins).

> [!NOTE]
> In an on-premises environment with no connection to Microsoft 365, or to deploy SharePoint add-ins or Office Add-ins that target Office 2013, use a [SharePoint app catalog](publish-task-pane-and-content-add-ins-to-an-add-in-catalog.md). To deploy COM/VSTO add-ins, use ClickOnce or Windows Installer, as described in [Deploying an Office solution](/visualstudio/vsto/deploying-an-office-solution).

## Recommended approach for deploying Office Add-ins

Consider deploying Office Add-ins in a phased approach to help ensure that the deployment goes smoothly. We recommend the following plan.

1. Deploy the add-in to a small set of business stakeholders and members of the IT department. If the deployment is successful, move on to step 2.

1. Deploy the add-in to a larger set of individuals within the business who will be using the add-in. If the deployment is successful, move on to step 3.

1. Deploy the add-in to the full set of individuals who will be using the add-in.

Depending on the size of the target audience, you may want to add steps to or remove steps from this procedure.

## Publish an Office Add-in via Centralized Deployment

Before you begin, confirm that your organization meets all requirements for using Centralized Deployment, as described in [Determine if Centralized Deployment of add-ins works for your Microsoft 365 organization](/microsoft-365/admin/manage/centralized-deployment-of-add-ins).

If your organization meets all requirements, complete the following steps to publish an Office Add-in via Centralized Deployment.

1. Sign in to Microsoft 365 with your work or education account.
1. Select the app launcher icon in the upper-left and choose **Admin**.
1. In the navigation menu, select **Show more**, then choose **Settings** > **Integrated apps**.
1. At the top of the page, choose **Add-ins**.
1. If you see a message on the top of the page announcing the new Microsoft 365 admin center, choose the message to go to the Admin Center Preview (see [About the Microsoft 365 admin center](/microsoft-365/admin/admin-overview/about-the-admin-center)).
1. Choose **Deploy Add-In** at the top of the page.
1. Choose **Next** after reviewing the requirements.
1. Choose one of the following options on the **Centralized Deployment** page.

    - **I want to add an Add-In from the Office Store.**
    - **I have the manifest file (.xml) on this device.** For this option, choose **Browse** to locate the manifest file (.xml) that you want to use.
    - **I have a URL for the manifest file.** For this option, type the manifest's URL in the field provided.

    ![New Add-In dialog in Microsoft 365 admin center.](../images/new-add-in.png)

1. If you selected the option to add an add-in from the Office Store, select the add-in. You can view available add-ins via categories of **Suggested for you**, **Rating**, or **Name**. You may only add free add-ins from Office Store. Adding paid add-ins isn't currently supported.

    > [!NOTE]
    > With the Office Store option, updates and enhancements to the add-in are automatically available to users without your intervention.

    ![Select an add-In dialog in Microsoft 365 admin center.](../images/select-an-add-in.png)

1. Choose **Continue** after reviewing the add-in details, Privacy Policy, and License Terms.

    ![Selected add-in page in Microsoft 365 admin center.](../images/selected-add-in-admin-center.png)

1. On the **Assign Users** page, choose **Everyone**, **Specific Users/Groups**, or **Only me**. Use the search box to find the users and groups to whom you want to deploy the add-in. For Outlook add-ins, you can also choose the deployment method **Fixed**, **Available**, or **Optional**.

    ![Manage who has access and deployment method in Microsoft 365 admin center.](../images/manage-users-deployment-admin-center.png)

    > [!NOTE]
    > Add-ins that utilize [single sign-on (SSO)](../develop/sso-in-office-add-ins.md) will prompt the admin to consent to the scopes listed in the add-in manifest.  If the same backing service is used across multiple add-ins (the same Azure App ID is used with SSO in different add-ins), the scopes for each add-in will be prompted for consent with each deployment. This page will also display the list of permissions that the add-in requires.

1. When finished, choose **Deploy**. This process may take up to three minutes. Then, finish the walkthrough by pressing **Next**. You now see your add-in along with other Office apps.

    > [!NOTE]
    > When an administrator chooses **Deploy**, consent is given for all users.

    ![List of apps in Microsoft 365 admin center.](../images/citations.png)

> [!TIP]
> When you deploy a new add-in to users and/or groups in your organization, consider sending them an email that describes when and how to use the add-in, and includes links to relevant Help content, FAQs, or other support resources.

## Considerations when granting access to an add-in

Admins can assign an add-in to everyone in the organization or to specific users and/or groups within the organization. The following list describes the implications of each option.

- **Everyone**: As the name implies, this option assigns the add-in to every user in the tenant. Use this option sparingly and only for add-ins that are truly universal to your organization.

- **Users**: If you assign an add-in to individual users, you'll need to update the Central Deployment settings for the add-in each time you want to assign it additional users. Likewise, you'll need to update the Central Deployment settings for the add-in each time you want to remove a user's access to the add-in.

- **Groups**: If you assign an add-in to a group, users who are added to the group will automatically be assigned the add-in. Likewise, when a user is removed from a group, the user automatically loses access to the add-in. In either case, no additional action is required from the Microsoft 365 admin.

In general, for ease of maintenance, we recommend assigning add-ins by using groups whenever possible. However, in situations where you want to restrict add-in access to a very small number of users, it may be more practical to assign the add-in to specific users.

## Add-in states

The following table describes the different states of an add-in.

|State|How the state occurs|Impact|
|-----|--------------------|------|
|**Active**|Admin uploaded the add-in and assigned it to users and/or groups.|Users and/or groups assigned to the add-in see it in the relevant Office clients.|
|**Turned off**|Admin turned off the add-in.|Users and/or groups assigned to the add-in no longer have access to it. If the add-in state is changed from **Turned off** to **Active**, the users and groups will regain access to it.|
|**Deleted**|Admin deleted the add-in.|Users and/or groups assigned the add-in no longer have access to it.|

## Updating Office Add-ins that are published via Centralized Deployment

After an Office Add-in has been published via Centralized Deployment, any changes made to the add-in's web application will automatically be available to all users after those changes are implemented in the web application. Changes made to an add-in's [XML manifest file](../develop/add-in-manifests.md) to, for example, update the add-in's icon, text, or add-in commands, happen as follows:

- **Line-of-business add-in**: If an admin explicitly uploaded a manifest file (either from their device or by pointing to a URL) when implementing Centralized Deployment via the Microsoft 365 admin center, the admin must upload a new manifest file that contains the desired changes. After the updated manifest file has been uploaded, the next time the relevant Office applications start, the add-in will update.

  > [!NOTE]
  > An admin doesn't need to remove a LOB add-in to make an update. In the Add-ins section, the admin can simply choose the LOB add-in and invoke this functionality by pressing the **Update add-in** button present in the bottom right corner.
  >
  > ![Screenshot shows the Update add-in dialog in Microsoft 365 admin center.](../images/update-add-in-admin-center.png)

- **Office Store add-in**: If an admin selected an add-in from the Office Store when implementing Centralized Deployment via the Microsoft 365 admin center, and the add-in updates in the Office Store, the add-in will update later via Centralized Deployment. It can take up to 24 hours for the Store add-in updates to flow for all end users. After this duration, the next time the relevant Office applications restart for these users, the add-in will update. Users can also trigger a Manual Refresh to get the latest Store add-in version by selecting **Insert Tab** > **Add-ins** > **Admin Managed Tab** > **Hit Refresh**.

## End user experience with add-ins

After an add-in has been published via Centralized Deployment, end users may start using it on any platform that the add-in supports.

If the add-in supports add-in commands, the commands will appear on the Office application ribbon for all users to whom the add-in is deployed. In the following example, the command **Search Citation** appears in the ribbon for the **Citations** add-in.

![Screenshot shows a section of the Office app ribbon with the Search Citation command highlighted in the Citations add-in.](../images/search-citation.png)

If the add-in does not support add-in commands, users can add it to their Office application by doing the following:

1. In Word 2016 or later, Excel 2016 or later, or PowerPoint 2016 or later, choose **Insert** > **My Add-ins**.
1. Choose the **Admin Managed** tab in the add-in window.
1. Choose the add-in, and then choose **Add**.

    ![Screenshot shows the Admin Managed tab of the Office Add-ins page of an Office application. The Citations add-in is shown on the tab.](../images/office-add-ins-admin-managed.png)

However, for Outlook 2016 or later, users can do the following:

1. In Outlook, choose **Home** > **Store**.
1. Choose the **Admin-managed** item under the add-in tab.
1. Choose the add-in, and then choose **Add**.

    ![Screenshot shows the Admin-managed area of the Store page of the Outlook application.](../images/outlook-add-ins-admin-managed.png)

## See also

- [Determine if Centralized Deployment of add-ins works for your Microsoft 365 organization](/office365/admin/manage/centralized-deployment-of-add-ins)
