---
title: Use the Yeoman generator to create an Office Add-in that uses SSO
description: Use the Yeoman generator to build a Node.js Office Add-in that uses single sign-on (preview).
ms.date: 01/07/2020
ms.prod: non-product-specific
localization_priority: Priority
---

# Use the Yeoman generator to create an Office Add-in that uses SSO

In this article, you'll walk through the process of using the Yeoman generator to create an Office Add-in that uses single sign-on (SSO). 
TODO: finish intro

> [!NOTE]
> > TODO: finish note | add corresponding note to detailed walkthrough
> To learn more about the details...see [detailed walkthrough].
> To learn basic concepts about SSO in Office Add-ins...see [overview article].

## Prerequisites

TODO: update prereqs -- copy from https://docs.microsoft.com/office/dev/add-ins/develop/create-sso-office-add-ins-nodejs ?

[!include[Yeoman generator prerequisites](../includes/quickstart-yo-prerequisites.md)]

## Create the add-in project

[!include[Yeoman generator create project guidance](../includes/yo-office-command-guidance.md)]

- **Choose a project type:** `Office Add-in Task Pane project`
- **Choose a script type:** `Javascript`
- **What do you want to name your add-in?** `My Office Add-in`
- **Which Office client application would you like to support?** `OneNote`

![A screenshot of the prompts and answers for the Yeoman generator](../images/yo-office-onenote.png)

After you complete the wizard, the generator creates the project and installs supporting Node components.

[!include[Yeoman generator next steps](../includes/yo-office-next-steps.md)]

## Explore the project

The add-in project that you've created with the Yeoman generator contains sample code for a very basic task pane add-in. 

- The **./manifest.xml** file in the root directory of the project defines the settings and capabilities of the add-in.
- The **./src/taskpane/taskpane.html** file contains the HTML markup for the task pane.
- The **./src/taskpane/taskpane.css** file contains the CSS that's applied to content in the task pane.
- The **./src/taskpane/taskpane.js** file contains the Office JavaScript API code that facilitates interaction between the task pane and the Office host application.

## Update the code

In your code editor, open the file **./src/taskpane/taskpane.js** and add the following code within the **run** function. This code uses the OneNote JavaScript API to set the page title and add an outline to the body of the page.

```js
try {
    await OneNote.run(async context => {

        // Get the current page.
        var page = context.application.getActivePage();

        // Queue a command to set the page title.
        page.title = "Hello World";

        // Queue a command to add an outline to the page.
        var html = "<p><ol><li>Item #1</li><li>Item #2</li></ol></p>";
        page.addOutline(40, 90, html);

        // Run the queued commands, and return a promise to indicate task completion.
        return context.sync();
    });
} catch (error) {
    console.log("Error: " + error);
}
```

## Try it out

1. Navigate to the root folder of the project.

    ```command&nbsp;line
    cd "My Office Add-in"
    ```

2. Start the local web server and sideload your add-in.

    > [!NOTE]
    > Office Add-ins should use HTTPS, not HTTP, even when you are developing. If you are prompted to install a certificate after you run one of the following commands, accept the prompt to install the certificate that the Yeoman generator provides.

    > [!TIP]
    > If you're testing your add-in on Mac, run the following command before proceeding. When you run this command, the local web server starts.
    >
    > ```command&nbsp;line
    > npm run dev-server
    > ```

    Run the following command in the root directory of your project. When you run this command, the local web server will start (if it's not already running).

    ```command&nbsp;line
    npm run start:web
    ```

3. In [OneNote on the web](https://www.onenote.com/notebooks), open a notebook and create a new page.

4. Choose **Insert > Office Add-ins** to open the Office Add-ins dialog.

    - If you're signed in with your consumer account, select the **MY ADD-INS** tab, and then choose **Upload My Add-in**.

    - If you're signed in with your work or school account, select the **MY ORGANIZATION** tab, and then select **Upload My Add-in**. 

    The following image shows the **MY ADD-INS** tab for consumer notebooks.

    <img alt="The Office Add-ins dialog showing the MY ADD-INS tab" src="../images/onenote-office-add-ins-dialog.png" width="500">

5. In the Upload Add-in dialog, browse to **manifest.xml** in your project folder, and then choose **Upload**. 

6. From the **Home** tab, choose the **Show Taskpane** button in the ribbon. The add-in task pane opens in an iFrame next to the OneNote page.

7. At the bottom of the task pane, choose the **Run** link to set the page title and add an outline to the body of the page.

    ![The OneNote add-in built from this walkthrough](../images/onenote-first-add-in-4.png)

## Next steps

Congratulations, you've successfully created a OneNote task pane add-in! Next, learn more about the [core concepts of building OneNote add-ins](../onenote/onenote-add-ins-programming-overview.md).

## See also

* [Office Add-ins platform overview](../overview/office-add-ins.md)
* [Building Office Add-ins](../overview/office-add-ins-fundamentals.md)
* [Develop Office Add-ins](../develop/develop-overview.md)
- [OneNote JavaScript API programming overview](../onenote/onenote-add-ins-programming-overview.md)
- [OneNote JavaScript API reference](/office/dev/add-ins/reference/overview/onenote-add-ins-javascript-reference)
- [Rubric Grader sample](https://github.com/OfficeDev/OneNote-Add-in-Rubric-Grader)

