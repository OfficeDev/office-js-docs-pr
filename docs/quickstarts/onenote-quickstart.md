---
title: Build your first OneNote task pane add-in
description: 
ms.date: 04/23/2019
ms.prod: onenote
localization_priority: Priority
---

# Build your first OneNote task pane add-in

In this article, you'll walk through the process of building a OneNote task pane add-in.

## Prerequisites

[!include[Yeoman generator prerequisites](../includes/quickstart-yo-prerequisites.md)]

## Create the add-in project

1. Use the Yeoman generator to create a OneNote add-in project. Run the following command and then answer the prompts as follows:

    ```command&nbsp;line
    yo office
    ```

    - **Choose a project type:** `Office Add-in Task Pane project`
    - **Choose a script type:** `Javascript`
    - **What do you want to name your add-in?:** `My Office Add-in`
    - **Which Office client application would you like to support?:** `OneNote`

    ![A screenshot of the prompts and answers for the Yeoman generator](../images/yo-office-onenote.png)
    
    After you complete the wizard, the generator will create the project and install supporting Node components.
	
2. Navigate to the root folder of the project.

    ```command&nbsp;line
    cd "My Office Add-in"
    ```

## Explore the project

The add-in project that you've created with the Yeoman generator contains sample code for a very basic task pane add-in. 

- The **./manifest.xml** file in the root directory of the project defines the settings and capabilities of the add-in.
- The **./src/taskpane/taskpane.html** file contains the HTML markup for the task pane.
- The **./src/taskpane/taskpane.css** file contains the CSS styles that are used by **taskpane.html**.
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

1. Start the local web server by running the following command:

    ```
    npm run start:web
    ```

    > [!NOTE]
    > Office Add-ins should use HTTPS, not HTTP, even when you are developing. If you are prompted to install a certificate after you run `npm run start:web`, accept the prompt to install the certificate that the Yeoman generator provides. 
    
2. In [OneNote Online](https://www.onenote.com/notebooks), open a notebook and create a new page.

3. Choose **Insert > Office Add-ins** to open the Office Add-ins dialog.

    - If you're signed in with your consumer account, select the **MY ADD-INS** tab, and then choose **Upload My Add-in**.

    - If you're signed in with your work or school account, select the **MY ORGANIZATION** tab, and then select **Upload My Add-in**. 

    The following image shows the **MY ADD-INS** tab for consumer notebooks.

    <img alt="The Office Add-ins dialog showing the MY ADD-INS tab" src="../images/onenote-office-add-ins-dialog.png" width="500">

3. In the Upload Add-in dialog, browse to **manifest.xml** in your project folder, and then choose **Upload**. 

4. From the **Home** tab, choose the **Show Taskpane** button in the ribbon. The add-in task pane opens in an iFrame next to the OneNote page.

5. At the bottom of the task pane, choose the **Run** link to set the page title and add an outline to the body of the page.

    ![The OneNote add-in built from this walkthrough](../images/onenote-first-add-in-4.png)

## Next steps

Congratulations, you've successfully created a OneNote task pane add-in! Next, learn more about the core concepts of building OneNote add-ins.

> [!div class="nextstepaction"]
> [OneNote JavaScript API programming overview](../onenote/onenote-add-ins-programming-overview.md)

## See also

- [OneNote JavaScript API programming overview](../onenote/onenote-add-ins-programming-overview.md)
- [OneNote JavaScript API reference](/office/dev/add-ins/reference/overview/onenote-add-ins-javascript-reference)
- [Rubric Grader sample](https://github.com/OfficeDev/OneNote-Add-in-Rubric-Grader)
- [Office Add-ins platform overview](../overview/office-add-ins.md)

