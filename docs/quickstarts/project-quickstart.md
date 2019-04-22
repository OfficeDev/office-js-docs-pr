---
title: Build your first Project task pane add-in
description: 
ms.date: 04/23/2019
ms.prod: project
localization_priority: Priority
---

# Build your first Project task pane add-in

In this article, you'll walk through the process of building a Project task pane add-in.

## Prerequisites

[!include[Yeoman generator prerequisites](../includes/quickstart-yo-prerequisites.md)]

- Project 2016 or later for Windows 

## Create the add-in

1. Use the Yeoman generator to create a Project add-in project. Run the following command and then answer the prompts as follows:

    ```bash
    yo office
    ```

    - **Choose a project type:** `Office Add-in Task Pane project`
    - **Choose a script type:** `Javascript`
    - **What do you want to name your add-in?:** `My Office Add-in`
    - **Which Office client application would you like to support?:** `Project`

    ![A screenshot of the prompts and answers for the Yeoman generator](../images/yo-office-project.png)
    
    After you complete the wizard, the generator will create the project and install supporting Node components.
	
2. Navigate to the root folder of the project.

    ```bash
    cd "My Office Add-in"
    ```

## Explore the project

The add-in project that you've created with the Yeoman generator contains sample code for a very basic task pane add-in. 

- The **./manifest.xml** file in the root directory of the project defines the settings and capabilities of the add-in.
- The **./src/taskpane/taskpane.html** file contains the HTML markup for the task pane.
- The **./src/taskpane/taskpane.css** file contains the CSS styles that are used by **taskpane.html**.
- The **./src/taskpane/taskpane.js** file contains the Office JavaScript API code that facilitates interaction between the task pane and the Office host application.

## Update the code

In your code editor, open the **./src/taskpane/taskpane.js** file and add the following code within the **run** function. This code uses the Office JavaScript API to set the `Name` field and `Notes` field of the selected task.

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

1. Start the local web server by running the following command:

    ```
    npm start
    ```

    > [!NOTE]
    > Office Web Add-ins should use HTTPS, not HTTP, even when you are developing. If you are prompted to install a certificate after you run `npm run start`, accept the prompt to install the certificate that the Yeoman generator provides. 

2. In Project, create a simple project plan.

3. Load your add-in in Project by following the instructions in [Sideload Office Add-ins on Windows](../testing/create-a-network-shared-folder-catalog-for-task-pane-and-content-add-ins.md).

4. Select a single task within the project.

5. Scroll to the bottom of the task pane and choose the **Run** link to rename the selected task and add notes to the selected task.

    ![Screenshot of the Project application with the task pane add-in loaded](../images/project-quickstart-addin-1.png)

## Next steps

Congratulations, you've successfully created a Project task pane add-in! Next, learn more about the capabilities of a Project add-in and explore common scenarios.

> [!div class="nextstepaction"]
> [Project add-ins](../project/project-add-ins.md)

