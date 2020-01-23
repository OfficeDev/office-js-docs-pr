---
title: Customize the SSO-enabled add-in that you created with the Yeoman generator
description: Learn about customizing the SSO-enabled add-in that you created with the Yeoman generator.
ms.date: 01/24/2020
ms.prod: non-product-specific
localization_priority: Normal
---

# Customize the SSO-enabled add-in you that created with the Yeoman generator

> [!IMPORTANT]
> This article builds upon the SSO-enabled add-in that's created by completing the [single sign-on (SSO) quick start](sso-quickstart.md). Please complete the quick start before reading this article.

The [SSO quick start](sso-quickstart.md) creates an SSO-enabled add-in that gets the signed-in user's profile information and writes it to the document. In this article, you'll walk through the process of customizing the SSO-enabled add-in that you created with the quick start, to add new functionality that requires different permissions.

## Prerequisites

* An Office Add-in that you created by following the instructions in the [SSO quick start](sso-quickstart.md).

* [Node.js](https://nodejs.org) (the latest [LTS](https://nodejs.org/about/releases) version)

[!include[additional prerequisites](../includes/sso-tutorial-prereqs.md)]

* At least a few files and folders stored on OneDrive for Business in your Office 365 subscription.

## Review contents of the project

Let's begin with a quick review of the add-in project that you've created with the Yeoman generator.

[!include[project structure for an SSO-enabled add-in created with the Yeoman generator](../includes/sso-yeoman-project-structure.md)]

## Add new functionality 

The add-in that you created with the SSO quick start uses Microsoft Graph to get the signed-in user's profile information and writes that information to the document. Let's add new functionality to get the first 10 file and folder names from the signed-in user's OneDrive for Business and write that information to the document. Enabling this new functionality requires updating code within the add-in project and also updating app permissions in Azure.

### Update the code

...TO DO...



### Update app permissions in Azure

Before the add-in can successfully implement the new functionality we've added to read the user's files from OneDrive, the app must be granted the appropriate permissions. Complete the following steps to grant the app the **Files.Read.All** permission.

1. Navigate to the [Azure portal](https://ms.portal.azure.com/#home) and sign in using your Office 365 administrator credentials. 

2. Navigate to the **App registrations** page, by doing either of the following:
    - Choose the **App registrations** tile on the Azure home page.
    - Use the search box on the home page to find and choose **App registrations**.

3. On the **App registrations** page, choose the app that you created during the quick start. 
    > [!TIP]
    > The **Display name** of the app will match the add-in name that you specified when you created the project with the Yeoman generator.

4. From the app overview page, choose **API permissions** under the **Manage** heading on the left side of the page.

5. Select **Add a permission**.

6. On the panel that opens choose **Microsoft Graph** and then choose **Delegated permissions**.

7. On the **Request API permissions** panel, under **Files**, select **Files.Read.All** and then select the **Add permissions** button at the bottom of the panel.





## Try it out

...TO DO...

1. In the root folder of the project, run the following command to build the project, start the local web server, and sideload your add-in in the previously selected Office client application.

    > [!NOTE]
    > Office Add-ins should use HTTPS, not HTTP, even when you are developing. If you are prompted to install a certificate after you run the following command, accept the prompt to install the certificate that the Yeoman generator provides.

    ```command&nbsp;line
    npm start
    ```

2. In the Office client application that opens when you run the previous command (i.e., Excel, Word or PowerPoint), make sure that you're signed in with a user that's a member of the same Office 365 organization as the Office 365 administrator account that you used to connect to Azure while configuring SSO in step 3 of the [previous section](#configure-sso). Doing so establishes the appropriate conditions for SSO to succeed. 

3. In the Office client application, choose the **Home** tab, and then choose the **Show Taskpane** button in the ribbon to open the add-in task pane. The following image shows this button in Excel.

    ![Excel add-in button](../images/excel-quickstart-addin-3b.png)

4. At the bottom of the task pane, choose the **Get My User Profile Information** button to initiate the SSO process. 

    > [!NOTE] 
    > If you're not already signed in to Office at this point, you'll be prompted to sign in. As described previously, you should sign in with a user that's a member of the same Office 365 organization as the Office 365 administrator account that you used to connect to Azure while configuring SSO in step 3 of the [previous section](#configure-sso), if you want SSO to succeed.

5. If a dialog window appears to request permissions on behalf of the add-in, this means that SSO is not supported for your scenario and the add-in has instead fallen back to an alternate method of user authentication. This may occur when the tenant administrator hasn't granted consent for the add-in to access Microsoft Graph, or when the user isn't signed into Office with a valid Microsoft Account or Office 365 ("Work or School") account. Choose the **Accept** button in the dialog window to continue.

    ![Permissions request dialog](../images/sso-permissions-request.png)

    > [!NOTE]
    > After a user accepts this permissions request, they won't be prompted again in the future.

6. The add-in retrieves profile information for the signed-in user and writes it to the document. The following image shows an example of profile information written to an Excel worksheet.

    ![User profile information in Excel worksheet](../images/sso-user-profile-info-excel.png)

## Next steps

...TO DO...

Congratulations, you've successfully created a task pane add-in that uses SSO when possible, and uses an alternate method of user authentication when SSO is not supported. To learn more about SSO configuration steps that the Yeoman generator completed automatically, and the code that facilitates the SSO process, see the [Create a Node.js Office Add-in that uses single sign-on](../develop/create-sso-office-add-ins-nodejs.md) tutorial.

## See also

- ...TO DO...
- [Enable single sign-on for Office Add-ins](../develop/sso-in-office-add-ins.md)
- [Create a Node.js Office Add-in that uses single sign-on](../develop/create-sso-office-add-ins-nodejs.md)
- [Troubleshoot error messages for single sign-on (SSO)](../develop/troubleshoot-sso-in-office-add-ins.md)