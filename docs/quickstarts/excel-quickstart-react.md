---
title: Use React to build an Excel task pane add-in
description: Learn how to build a simple Excel task pane add-in by using the Office JS API and React.
ms.date: 07/07/2021
ms.prod: excel
localization_priority: Priority
---

# Use React to build an Excel task pane add-in

In this article, you'll walk through the process of building an Excel task pane add-in using React and the Excel JavaScript API.

## Prerequisites

[!include[Set up requirements](../includes/set-up-dev-environment-beforehand.md)]
[!include[Yeoman generator prerequisites](../includes/quickstart-yo-prerequisites.md)]

## Create the add-in project

[!include[Yeoman generator create project guidance](../includes/yo-office-command-guidance.md)]

- **Choose a project type:** `Office Add-in Task Pane project using React framework`
- **Choose a script type:** `TypeScript`
- **What do you want to name your add-in?** `My Office Add-in`
- **Which Office client application would you like to support?** `Excel`

![Screenshot of the Yeoman Office Add-in generator command line interface, with project type set to the React framework.](../images/yo-office-excel-react-2.png)

After you complete the wizard, the generator creates the project and installs supporting Node components.

[!include[Yeoman generator next steps](../includes/yo-office-next-steps.md)]

## Explore the project

The add-in project that you've created with the Yeoman generator contains sample code for a basic task pane add-in. If you'd like to explore the key components of your add-in project, open the project in your code editor and review the files listed below. When you're ready to try out your add-in, proceed to the next section.

- The **manifest.xml** file in the root directory of the project defines the settings and capabilities of the add-in. To learn more about the **manifest.xml** file, see [Office Add-ins XML manifest](../develop/add-in-manifests.md).
- The **./src/taskpane/taskpane.html** file defines the HTML framework of the task pane, and the files within the **./src/taskpane/components** folder define the various parts of the task pane UI.
- The **./src/taskpane/taskpane.css** file contains the CSS that's applied to content in the task pane.
- The **./src/taskpane/components/App.tsx** file contains the Office JavaScript API code that facilitates interaction between the task pane and Excel.

## Try it out

1. Navigate to the root folder of the project.

    ```command&nbsp;line
    cd "My Office Add-in"
    ```

1. [!include[Start server section](../includes/quickstart-yo-start-server-excel.md)] 

1. In Excel, choose the **Home** tab, and then choose the **Show Taskpane** button in the ribbon to open the add-in task pane.

    ![Screenshot of the Excel Home menu, with the Show Taskpane button highlighted.](../images/excel-quickstart-addin-3b.png)

1. Select any range of cells in the worksheet.

1. At the bottom of the task pane, choose the **Run** link to set the color of the selected range to yellow.

    ![Screenshot of Excel, with the add-in task pane open, and the Run button highlighted in the add-in task pane.](../images/excel-quickstart-addin-3c.png)

## Next steps

Congratulations, you've successfully created an Excel task pane add-in using React! Next, learn more about the capabilities of an Excel add-in and build a more complex add-in by following along with the Excel add-in tutorial.

> [!div class="nextstepaction"]
> [Excel add-in tutorial](../tutorials/excel-tutorial.md)

## See also

- [Excel add-in tutorial](../tutorials/excel-tutorial.md)
- [Excel JavaScript object model in Office Add-ins](../excel/excel-add-ins-core-concepts.md)
- [Excel add-in code samples](https://developer.microsoft.com/office/gallery/?filterBy=Samples,Excel)
- [Excel JavaScript API reference](../reference/overview/excel-add-ins-reference-overview.md)