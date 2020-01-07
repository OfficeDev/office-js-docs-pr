---
title: Use the Yeoman generator to create an Office Add-in that uses SSO (preview)
description: Use the Yeoman generator to build a Node.js Office Add-in that uses single sign-on (preview).
ms.date: 01/08/2020
ms.prod: non-product-specific
localization_priority: Priority
---

# Use the Yeoman generator to create an Office Add-in that uses single sign-on (preview)

In this article, you'll walk through the process of using the Yeoman generator to create an Office Add-in for Excel, Word, or PowerPoint that uses single sign-on (SSO). 

> [!TIP]
> Before you attempt to complete this quick start, review [Enable single sign-on for Office Add-ins](../develop/sso-in-office-add-ins.md) to learn basic concepts about SSO in Office Add-ins. 
 
The Yeoman generator simplifies the process of creating an SSO add-in, by automating the steps required to configure SSO within Azure and generating the code that's necessary for an add-in to use SSO. For a detailed walkthrough that describes how to manually complete the steps that the Yeoman generator automates, see the [Create a Node.js Office Add-in that uses single sign-on](../develop/create-sso-office-add-ins-nodejs.md) tutorial.

## Prerequisites

- [Node.js](https://nodejs.org) (version 10.15.0 or later)

- The latest version of [Yeoman](https://github.com/yeoman/yo) and the [Yeoman generator for Office Add-ins](https://github.com/OfficeDev/generator-office). To install these tools globally, run the following command via the command prompt:

    ```command&nbsp;line
    npm install -g yo generator-office
    ```

    [!include[note to update Yeoman generator](../includes/note-yeoman-generator-update.md)]

- Office 365 (the subscription version of Office) account which you can get by joining the [Office 365 Developer Program](https://aka.ms/devprogramsignup) that includes a free 1 year subscription to Office 365. You should use the latest monthly version and build from the Insiders channel but you need to be an Office Insider to get this version. For more information, see [Be an Office Insider](https://products.office.com/office-insider?tab=tab-1). Please note that when a build graduates to the production semi-annual channel, support for preview features, including SSO, is turned off for that build.

- A Microsoft Azure Tenant. This add-in requires Azure Active Directory (AD). Azure AD provides identity services that applications use for authentication and authorization. A trial subscription can be acquired at [Microsoft Azure](https://account.windowsazure.com/SignUp).

## Create the add-in project

> [!TIP]
> The Yeoman generator can create an SSO-enabled Office Add-in for Excel, Word, or PowerPoint, and can be created with script type of JavaScript or TypeScript. The following instructions specify `JavaScript` and `Excel`, but you should choose the script type and Office client application that best suits your scenario.

[!include[Yeoman generator create project guidance](../includes/yo-office-command-guidance.md)]

- **Choose a project type:** `Office Add-in Task Pane project supporting single sign-on`
- **Choose a script type:** `Javascript`
- **What do you want to name your add-in?** `My SSO Office Add-in`
- **Which Office client application would you like to support?** `Excel`

![A screenshot of the prompts and answers for the Yeoman generator](../images/yo-office-sso-excel.png)

After you complete the wizard, the generator creates the project and installs supporting Node components.

[!include[Yeoman generator next steps](../includes/yo-office-next-steps.md)]

## Explore the project

The add-in project that you've created with the Yeoman generator contains sample code for an SSO-enabled task pane add-in. 

- The **./ENV** file in the root directory of the project defines constants that are used by the add-in project.
- The **./manifest.xml** file in the root directory of the project defines the settings and capabilities of the add-in.
- The **./src/helpers/documentHelper.js** file ...
- The **./src/helpers/fallbackauthdialog.html** file ...
- The **./src/helpers/fallbackauthdialog.js** file ...
- The **./src/helpers/fallbackauthhelper.js** file ...
- The **./src/helpers/ssoauthhelper.js** file ...
- The **./src/taskpane/taskpane.html** file contains the HTML markup for the task pane.
- The **./src/taskpane/taskpane.css** file contains the CSS that's applied to content in the task pane.
- The **./src/taskpane/taskpane.js** file contains the Office JavaScript API code that facilitates interaction between the task pane and the Office host application.

## Try it out

1. Navigate to the root folder of the project.

    ```command&nbsp;line
    cd "My SSO Office Add-in"
    ```

2. Run the following command to configure SSO for the add-in.

    ```command&nbsp;line
    npm run configure-sso
    ```

3. A web browser window will open and prompt you to enter your Azure credentials. Enter your Azure credentials when prompted. These credentials will be used to register a new application in Azure and configure the settings required by SSO.  

4. After you enter your Azure credentials, close the browser window and return to the command prompt. As the SSO configuration process continues, you'll see status messages being written to the console. As described in the console messages, files within the add-in project that the Yeoman generator created are automatically updated with data that's required by the SSO process.

5. When the SSO configuration process completes, run the following command to start the local web server and sideload your add-in.

    > [!NOTE]
    > Office Add-ins should use HTTPS, not HTTP, even when you are developing. If you are prompted to install a certificate after you run the following command, accept the prompt to install the certificate that the Yeoman generator provides.

    ```command&nbsp;line
    npm start
    ```

6. In the Office client application that you chose when creating the add-in project with the Yeoman generator (i.e., Excel, Word or PowerPoint), choose the **Home** tab, and then choose the **Show Taskpane** button in the ribbon to open the add-in task pane. The following image shows this button in Excel.

    ![Excel add-in button](../images/excel-quickstart-addin-3b.png)

7. At the bottom of the task pane, choose the **Get My User Profile Information** button to initiate the SSO process.

8. If you're not alredy signed in to Office, you'll be prompted to sign in. Sign in with the credentials that you want the add-in to use for SSO.

9. A dialog window appears to inform you about the permissions that the add-in is requesting. Because this add-in will write the signed-in user's profile information to the document, it needs permsissions to sign you in and read your profile and to maintain access to data you've given it access to. Choose the **Accept** button to grant those permissions to the add-in.

    ![Permissions request dialog](../images/sso-permissions-request.png)

10. The add-in retrieves profile information for the signed-in user and writes it to the document. The following image shows an example of profile information written to an Excel worksheet.

    ![User profile information in Excel worksheet](../images/sso-user-profile-info-excel.png)

## Next steps

Congratulations, you've successfully created a task pane add-in that uses SSO. To learn more about SSO configuration steps that the Yeoman generator completed automatically, and the code that facilitates the SSO process, see the [Create a Node.js Office Add-in that uses single sign-on](../develop/create-sso-office-add-ins-nodejs.md) tutorial.

## See also

- [Enable single sign-on for Office Add-ins](../develop/sso-in-office-add-ins.md)
- [Create a Node.js Office Add-in that uses single sign-on](../develop/create-sso-office-add-ins-nodejs.md)
- [Troubleshoot error messages for single sign-on (SSO)](../develop/troubleshoot-sso-in-office-add-ins.md)