---
title: Single sign-on (SSO) quick start
description: Use the Yeoman generator to build a Node.js Office Add-in that uses single sign-on.
ms.date: 05/16/2025
ms.service: microsoft-365
ms.localizationpriority: high
---

# Single sign-on (SSO) quick start

In this article, you'll use the Yeoman generator for Office Add-ins to create an Office Add-in for Excel, Outlook, Word, or PowerPoint that uses single sign-on (SSO).

> [!NOTE]
> The SSO template provided by the Yeoman generator for Office Add-ins only runs on localhost and cannot be deployed. If you're building a new Office Add-in with SSO for production purposes, follow the instructions in [Create a Node.js Office Add-in that uses single sign-on](../develop/create-sso-office-add-ins-nodejs.md).

## Prerequisites

- [Node.js](https://nodejs.org) (the latest [LTS](https://nodejs.org/en/about/previous-releases) version).

- The latest version of [Yeoman](https://github.com/yeoman/yo) and the [Yeoman generator for Office Add-ins](../develop/yeoman-generator-overview.md). To install these tools globally, run the following command via the command prompt.

    ```command&nbsp;line
    npm install -g yo generator-office
    ```

    [!include[note to update Yeoman generator](../includes/note-yeoman-generator-update.md)]

- If you're using a Mac and don't have the Azure CLI installed on your machine, you must install [Homebrew](https://brew.sh/). The SSO configuration script that you'll run during this quick start will use Homebrew to install the Azure CLI, and will then use the Azure CLI to configure SSO within Azure.

## Create the add-in project

> [!TIP]
> The Yeoman generator can create an SSO-enabled Office Add-in for Excel, Outlook, Word, or PowerPoint with script type of JavaScript or TypeScript. The following instructions specify `JavaScript` and `Excel`, but you should choose the script type and Office client application that best suits your scenario.

[!include[Yeoman generator create project guidance](../includes/yo-office-command-guidance.md)]

- **Choose a project type:** `Office Add-in Task Pane project supporting single sign-on (localhost)`
- **Choose a script type:** `JavaScript`
- **What do you want to name your add-in?** `My Office Add-in`
- **Which Office client application would you like to support?** Choose `Excel`, `Outlook`, `Word`, or `Powerpoint`.

:::image type="content" source="../images/yo-office-sso-excel.png" alt-text="Prompts and answers for the Yeoman generator in a command line interface.":::

After you complete the wizard, the generator creates the project and installs supporting Node components.

## Explore the project

The add-in project that you've created with the Yeoman generator contains code for an SSO-enabled task pane add-in.

[!include[project structure for an SSO-enabled add-in created with the Yeoman generator](../includes/sso-yeoman-project-structure.md)]

## Configure SSO

Now that your add-in project is created and contains the code that's necessary to facilitate the SSO process, complete the following steps to configure SSO for your add-in.

1. Go to the root folder of the project.

    ```command&nbsp;line
    cd "My Office Add-in"
    ```

1. Run the following command to configure SSO for the add-in.

    ```command&nbsp;line
    npm run configure-sso
    ```

    > [!WARNING]
    > This command will fail if your tenant is configured to require two-factor authentication. In this scenario, you'll need to manually complete the Azure app registration and SSO configuration steps by following all the steps in the [Create a Node.js Office Add-in that uses single sign-on](../develop/create-sso-office-add-ins-nodejs.md) tutorial.

1. A web browser window will open and prompt you to sign in to Azure. Sign in to Azure using your Microsoft 365 administrator credentials. These credentials will be used to register a new application in Azure and configure the settings required by SSO.

    > [!NOTE]
    > If you sign in to Azure using non-administrator credentials during this step, the `configure-sso` script won't be able to provide administrator consent for the add-in to users within your organization. SSO will therefore not be available to users of the add-in and they'll be prompted to sign-in.

1. After you enter your credentials, close the browser window and return to the command prompt. As the SSO configuration process continues, you'll see status messages being written to the console. As described in the console messages, files within the add-in project that the Yeoman generator created are automatically updated with data that's required by the SSO process.

## Test your add-in

If you've created an Excel, Word, or PowerPoint add-in, complete the steps in the following section to try it. If you've created an Outlook add-in, complete the steps in the [Outlook](#outlook) section instead.

### Excel, Word, and PowerPoint

Complete the following steps to test an Excel, Word, or PowerPoint add-in.

1. When the SSO configuration process completes, run the following command to build the project, start the local web server, and sideload your add-in in the previously selected Office client application.

    [!INCLUDE [alert use https](../includes/alert-use-https.md)]

    ```command&nbsp;line
    npm start
    ```

1. When Excel, Word, or PowerPoint opens when you run the previous command, make sure you're signed in with a user account that's a member of the same Microsoft 365 organization as the Microsoft 365 administrator account that you used to connect to Azure while configuring SSO in step 3 of the [previous section](#configure-sso). Doing so establishes the appropriate conditions for SSO to succeed.

1. In the Office client application, choose the **Home** tab, and then choose **Show Taskpane** to open the add-in task pane.

    :::image type="content" source="../images/excel-quickstart-addin-3b.png" alt-text="Excel add-in button.":::

1. At the bottom of the task pane, choose the **Get My User Profile Information** button to initiate the SSO process.

1. If a dialog window appears to request permissions on behalf of the add-in, this means that SSO is not supported for your scenario and the add-in has instead fallen back to an alternate method of user authentication. This may occur when the tenant administrator hasn't granted consent for the add-in to access Microsoft Graph, or when the user isn't signed in to Office with a valid Microsoft account or Microsoft 365 Education or Work account. Choose **Accept** to continue.

    ![The permissions requested dialog with Accept button highlighted.](../images/sso-permissions-request.png)

    > [!NOTE]
    > After a user accepts this permissions request, they won't be prompted again in the future.

1. The add-in retrieves profile information for the signed-in user and writes it to the document. The following image shows an example of profile information written to an Excel worksheet.

    ![The user profile information in Excel worksheet.](../images/sso-user-profile-info-excel.png)

1. [!include[Instructions to stop web server and uninstall dev add-in](../includes/stop-uninstall-dev-add-in.md)]

### Outlook

Complete the following steps to try out an Outlook add-in.

1. When the SSO configuration process completes, run the following command to build the project and start the local web server.

    [!INCLUDE [alert use https](../includes/alert-use-https.md)]

    ```command&nbsp;line
    npm start
    ```

1. Outlook will start and sideload the add-in. Make sure that you're signed in to Outlook with a user that's a member of the same Microsoft 365 organization as the Microsoft 365 administrator account that you used to connect to Azure while configuring SSO in step 3 of the [previous section](#configure-sso). Doing so establishes the appropriate conditions for SSO to succeed.

1. In Outlook, compose a new message.

1. In the message compose window, choose the **Show Taskpane** button to open the add-in task pane.

    ![The highlighted add-in ribbon button in Outlook compose message window.](../images/outlook-sso-ribbon-button.png)

1. At the bottom of the task pane, choose the **Get My User Profile Information** button to initiate the SSO process.

1. If a dialog window appears to request permissions on behalf of the add-in, this means that SSO is not supported for your scenario and the add-in has instead fallen back to an alternate method of user authentication. This may occur when the tenant administrator hasn't granted consent for the add-in to access Microsoft Graph, or when the user isn't signed in to Office with a valid Microsoft account or Microsoft 365 Education or Work account. Choose **Accept** to continue.

    ![The permissions requested dialog with Accept button highlighted.](../images/sso-permissions-request.png)

    > [!NOTE]
    > After a user accepts this permissions request, they won't be prompted again in the future.

1. The add-in retrieves profile information for the signed-in user and writes it to the body of the email message.

    ![The user profile information in Outlook compose message window.](../images/sso-user-profile-info-outlook.png)

1. [!include[Instructions to stop web server and uninstall dev add-in](../includes/stop-uninstall-outlook-dev-add-in.md)]

## Next steps

Congratulations, you've successfully created a task pane add-in that uses SSO when possible, and uses an alternate method of user authentication when SSO is not supported. To learn about customizing your add-in to add new functionality that requires different permissions, see [Customize your Node.js SSO-enabled add-in](sso-quickstart-customize.md).

[!include[The common troubleshooting section for all Yo Office quick starts](../includes/quickstart-troubleshooting-yo.md)]

## See also

- [Enable single sign-on for Office Add-ins](../develop/sso-in-office-add-ins.md)
- [Customize your Node.js SSO-enabled add-in](sso-quickstart-customize.md)
- [Create a Node.js Office Add-in that uses single sign-on](../develop/create-sso-office-add-ins-nodejs.md)
- [Troubleshoot error messages for single sign-on (SSO)](../develop/troubleshoot-sso-in-office-add-ins.md)
- [Using Visual Studio Code to publish](../publish/publish-add-in-vs-code.md#using-visual-studio-code-to-publish)