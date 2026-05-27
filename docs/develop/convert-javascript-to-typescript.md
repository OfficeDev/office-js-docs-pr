---
title: Convert an Office Add-in from JavaScript to TypeScript in Visual Studio
description: Convert a Visual Studio Office Add-in project from JavaScript to TypeScript by adding TypeScript build settings and updating project files.
ms.topic: how-to
ms.date: 05/26/2026
ms.localizationpriority: medium
ai-usage: ai-assisted
---

# Convert an Office Add-in from JavaScript to TypeScript in Visual Studio

If your Office Add-in project in Visual Studio starts in JavaScript, you can migrate it to TypeScript without rebuilding the project. This article shows the process for an Excel add-in. You can use the same steps for other Office Add-in project types in Visual Studio.

## What you'll do

- Add the `Microsoft.TypeScript.MSBuild` package so TypeScript files transpile during build.
- Add `tsconfig.json` and `package.json` settings for TypeScript and type definitions.
- Rename JavaScript files from `.js` to `.ts` and update code that needs TypeScript changes.
- Run the converted add-in in Excel.

## Prerequisites

- [Visual Studio 2022 or later](https://visualstudio.microsoft.com/downloads/) with the **Office/SharePoint development** workload installed

    > [!TIP]
    > If you've previously installed Visual Studio, [use the Visual Studio Installer](/visualstudio/install/modify-visual-studio) to ensure that the **Office/SharePoint development** workload is installed. If this workload is not yet installed, use the Visual Studio Installer to [install it](/visualstudio/install/modify-visual-studio#modify-workloads).

- Excel 2016 or later.

## Create the add-in project

> [!NOTE]
> [Skip this section](#convert-the-add-in-project-to-typescript) if you already have an existing project.

1. In Visual Studio, choose **Create a new project**. If the Visual Studio development environment is already open, you can create a new project by choosing **File** > **New** > **Project** on the menu bar.

1. Using the search box, enter **add-in**. Select **Excel Web Add-in**, and then select **Next**.

1. Name your project and select **Create**.

1. In the **Create Office Add-in** dialog, select **Add new functionalities to Excel**, and then select **Finish** to create the project.

1. Visual Studio creates a solution and its two projects appear in **Solution Explorer**. The **Home.html** file opens in Visual Studio.

## Convert the add-in project to TypeScript

### Add a NuGet package

1. Open the NuGet Package Manager by selecting **Tools** > **NuGet Package Manager** > **Manage NuGet Packages for Solution**.
1. Select the **Browse** tab. Search for and select **Microsoft.TypeScript.MSBuild**. Install this package to the ASP.NET web project, or update it if it's already installed. The ASP.NET web project has your project name with the text `Web` appended to the end. This will ensure the project will transpile to JavaScript when the build runs.

> [!NOTE]
> In your TypeScript project, you can have a mix of TypeScript and JavaScript files and your project will compile. This is because TypeScript is a typed superset of JavaScript that compiles JavaScript.

### Create a TypeScript config file

1. In **Solution Explorer**, right-click (or select and hold) the ASP.NET web project and choose **Add** > **New Item**. The ASP.NET web project has your project name with the text `Web` appended to the end.
1. In the **Add New Item** dialog, search for and select **TypeScript JSON Configuration File**. Select **Add** to create a **tsconfig.json** file.
1. Update the **tsconfig.json** file to also have an `include` section as shown in the following JSON.

    ```json
    {
      "compilerOptions": {
        "noImplicitAny": false,
        "noEmitOnError": true,
        "removeComments": false,
        "sourceMap": true,
        "target": "es5",
        "lib": [ 
          "es2015",
          "dom"
        ]
      },
      "exclude": [
        "node_modules",
        "wwwroot"
      ],
      "include": [
        "scripts/**/*",
        "**/*"
      ]
    }
    ```

1. Save the file. For more information on **tsconfig.json** settings, see [What is a tsconfig.json?](https://www.typescriptlang.org/docs/handbook/tsconfig-json.html).

### Create an npm configuration file

1. In **Solution Explorer**, right-click (or select and hold) the ASP.NET web project and choose **Add** > **New Item**. The ASP.NET web project has your project name with the text `Web` appended to the end.
1. In the **Add New Item** dialog, search for **npm Configuration File**. Select **Add** to create a **package.json** file.
1. Update the **package.json** file to have the `@types/jquery` package in the `devDependencies` section, as shown in the following JSON.

    ```json
    {
      "version": "1.0.0",
      "name": "asp.net",
      "private": true,
      "devDependencies": {
        "@types/jquery": "^3.5.30"
      }
    }
    ```

1. Save the file.
1. Open npm project properties by going to **Tools** > **Options**, then **Projects and Solutions** > **Web Package Management** > **Package Restore**. Set both **Restore On Project Open** and **Restore On Save** to `True`. Select **OK** to save the settings.

### Update the JavaScript files

Change your JavaScript files (**.js**) to TypeScript files (**.ts**). Then, make the necessary changes for them to compile. This section walks through the default files in a new project.

1. Find the **Home.js** file and rename it to **Home.ts**.

1. Find the **./Functions/FunctionFile.js** file and rename it to **FunctionFile.ts**.

1. Find the **./Scripts/MessageBanner.js** file and rename it to **MessageBanner.ts**.

1. In **./Scripts/MessageBanner.ts**, find the line `_onResize(null);` and replace it with the following code.

    ```typescript
    _onResize();
    ```

The JavaScript files generated by Visual Studio do not contain any TypeScript syntax. You should consider updating them. For example, the following code shows how to update the parameters to `showNotification` to include the string types.

```typescript
function showNotification(header: string, content: string) {
  $("#notification-header").text(header);
  $("#notification-body").text(content);
  messageBanner.showBanner();
  messageBanner.toggleExpansion();
}
```

## Run the converted add-in project

1. In Visual Studio, press <kbd>F5</kbd> or select the **Start** button to launch Excel with the **Show Taskpane** add-in button displayed on the ribbon. The add-in is hosted locally on IIS.

1. In Excel, select the **Home** tab, and then select the **Show Taskpane** button on the ribbon to open the add-in task pane.

1. In the worksheet, select the nine cells that contain numbers.

1. In the task pane, select **Highlight** to highlight the cell in the selected range that contains the highest value.

## See also

- [Develop Office Add-ins with Visual Studio](develop-add-ins-visual-studio.md)
- [Build an Excel task pane add-in with Visual Studio](../quickstarts/excel-quickstart-vs.md)
- [Publish your add-in using Visual Studio](../publish/package-your-add-in-using-visual-studio.md)
- [Office Add-in samples on GitHub](https://github.com/OfficeDev/Office-Add-in-samples)
