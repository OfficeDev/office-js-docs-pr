---
title: Use the Yeoman generator to create an Office Add-in that uses SSO
description: Use the Yeoman generator to build a Node.js Office Add-in that uses single sign-on (preview).
ms.date: 01/08/2020
ms.prod: non-product-specific
localization_priority: Priority
---

# Use the Yeoman generator to create an Office Add-in that uses SSO

In this article, you'll walk through the process of using the Yeoman generator to create an Office Add-in for Excel, Word, or PowerPoint that uses single sign-on (SSO). 

> [!TIP]
> Before you attempt to complete this quick start, review [Enable single sign-on for Office Add-ins](../develop/sso-in-office-add-ins.md) to learn basic concepts about SSO in Office Add-ins. 
 
The Yeoman generator simplifies the process of creating an SSO add-in, by automating the steps required to configure SSO within Azure and generating the code that's necessary for an add-in to use SSO. For detailed information about how to manually complete the steps that the Yeoman generator automates, see the [Create a Node.js Office Add-in that uses single sign-on](../develop/create-sso-office-add-ins-nodejs.md) tutorial.

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

- The **./manifest.xml** file in the root directory of the project defines the settings and capabilities of the add-in.
- The **./ENV** file in the root directory of the project defines constants that are used by the add-in project.
- The **./src/helpers/documentHelper.js** file ...[@Ricky and @Courtney] - please provide a brief description of this file in your PR feedback...
- The **./src/helpers/fallbackauthdialog.html** file ...[@Ricky and @Courtney] - please provide a brief description of this file in your PR feedback...
- The **./src/helpers/fallbackauthdialog.js** file ...[@Ricky and @Courtney] - please provide a brief description of this file in your PR feedback...
- The **./src/helpers/fallbackauthhelper.js** file ...[@Ricky and @Courtney] - please provide a brief description of this file in your PR feedback...
- The **./src/helpers/ssoauthhelper.js** file ...[@Ricky and @Courtney] - please provide a brief description of this file in your PR feedback...
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

3. A web browser window will open and prompt you to enter your Azure credentials. Enter your Azure credentials when prompted. These credentials will be used to register a new application in Azure and configure the settings required by SSO. Additionally, files within the add-in project that the Yeoman generator created will automatically be updated with the required data (i.e., application ID and port). 

4. Return to the command prompt and run the following command to start the local web server and sideload your add-in.

    > [!NOTE]
    > Office Add-ins should use HTTPS, not HTTP, even when you are developing. If you are prompted to install a certificate after you run the following command, accept the prompt to install the certificate that the Yeoman generator provides.

    ```command&nbsp;line
    npm start
    ```

5. 


## Next steps

Congratulations, you've successfully created a OneNote task pane add-in! Next, learn more about the [core concepts of building OneNote add-ins](../onenote/onenote-add-ins-programming-overview.md).

## See also

* [Office Add-ins platform overview](../overview/office-add-ins.md)
* [Building Office Add-ins](../overview/office-add-ins-fundamentals.md)
* [Develop Office Add-ins](../develop/develop-overview.md)
- [OneNote JavaScript API programming overview](../onenote/onenote-add-ins-programming-overview.md)
- [OneNote JavaScript API reference](/office/dev/add-ins/reference/overview/onenote-add-ins-javascript-reference)
- [Rubric Grader sample](https://github.com/OfficeDev/OneNote-Add-in-Rubric-Grader)

