---
title: Build your first Project task pane add-in
description: 
ms.date: 01/16/2020
ms.prod: project
localization_priority: Normal
---

# Build your first Project task pane add-in

In this article, you'll walk through the process of building a Project task pane add-in.

## Prerequisites

[!include[Yeoman generator prerequisites](../includes/quickstart-yo-prerequisites.md)]

- Project 2016 or later on Windows

## Create the add-in

[!include[Yeoman generator create project guidance](../includes/yo-office-command-guidance.md)]

- **Choose a project type:** `Office Add-in Task Pane project`
- **Choose a script type:** `Javascript`
- **What do you want to name your add-in?** `My Office Add-in`
- **Which Office client application would you like to support?** `Project`

![A screenshot of the prompts and answers for the Yeoman generator](../images/yo-office-project.png)

After you complete the wizard, the generator creates the project and installs supporting Node components.

[!include[Yeoman generator next steps](../includes/yo-office-next-steps.md)]

## Explore the project

The add-in project that you've created with the Yeoman generator contains sample code for a very basic task pane add-in. 

- The **./manifest.xml** file in the root directory of the project defines the settings and capabilities of the add-in.
- The **./src/taskpane/taskpane.html** file contains the HTML markup for the task pane.
- The **./src/taskpane/taskpane.css** file contains the CSS that's applied to content in the task pane.
- The **./src/taskpane/taskpane.js** file contains the Office JavaScript API code that facilitates interaction between the task pane and the Office host application.

## Update the code

In your code editor, open the file **./src/taskpane/taskpane.js** and add the following code within the **run** function. This code uses the Office JavaScript API to set the `Name` field and `Notes` field of the selected task.

```js
var taskGuid;

// Get the GUID of the selected task
Office.context.document.getSelectedTaskAsync(
    function (result) {
        if (result.status === Office.AsyncResultStatus.Succeeded) {
            taskGuid = result.value;

            // Set the specified fields for the selected task.
            var targetFields = [Office.ProjectTaskFields.Name, Office.ProjectTaskFields.Notes];
            var fieldValues = ['New task name', 'Notes for the task.'];

            // Set the field value. If the call is successful, set the next field.
            for (var i = 0; i < targetFields.length; i++) {
                Office.context.document.setTaskFieldAsync(
                    taskGuid,
                    targetFields[i],
                    fieldValues[i],
                    function (result) {
                        if (result.status === Office.AsyncResultStatus.Succeeded) {
                            i++;
                        }
                        else {
                            var err = result.error;
                            console.log(err.name + ' ' + err.code + ' ' + err.message);
                        }
                    }
                );
            }
        } else {
            var err = result.error;
            console.log(err.name + ' ' + err.code + ' ' + err.message);
        }
    }
);
```

## Try it out

1. Navigate to the root folder of the project.

    ```command&nbsp;line
    cd "My Office Add-in"
    ```

2. Start the local web server.

    > [!NOTE]
    > Office Add-ins should use HTTPS, not HTTP, even when you are developing. If you are prompted to install a certificate after you run the following command, accept the prompt to install the certificate that the Yeoman generator provides.

    Run the following command in the root directory of your project. When you run this command, the local web server will start.

    ```command&nbsp;line
    npm start
    ```

3. In Project, create a simple project plan.

4. Load your add-in in Project by following the instructions in [Sideload Office Add-ins on Windows](../testing/create-a-network-shared-folder-catalog-for-task-pane-and-content-add-ins.md).

5. Select a single task within the project.

6. At the bottom of the task pane, choose the **Run** link to rename the selected task and add notes to the selected task.

    ![Screenshot of the Project application with the task pane add-in loaded](../images/project-quickstart-addin-1.png)

## Next steps

Congratulations, you've successfully created a Project task pane add-in! Next, learn more about the capabilities of a Project add-in and explore common scenarios.

> [!div class="nextstepaction"]
> [Project add-ins](../project/project-add-ins.md)

## See also

- [Building Office Add-ins](../overview/office-add-ins-fundamentals.md)
- [Core concepts for Office Add-ins](../overview/core-concepts-office-add-ins.md)
- [Develop Office Add-ins](../develop/develop-overview.md)
