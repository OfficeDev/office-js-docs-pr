---
title: Convert an Office Add-in project in Visual Studio to TypeScript
description: Learn how to convert an Office Add-in project in Visual Studio to use TypeScript.
ms.topic: how-to
ms.date: 05/12/2025
ms.localizationpriority: medium
---

# Convert an Office Add-in project in Visual Studio to TypeScript

You can use the Office Add-in template in Visual Studio to create an add-in that uses JavaScript, and then convert that add-in project to TypeScript. This article describes this conversion process for an Excel add-in. You can use the same process to convert other types of Office Add-in projects from JavaScript to TypeScript in Visual Studio.

## Prerequisites

- [Visual Studio 2022 or later](https://www.visualstudio.com/vs/) with the **Office/SharePoint development** workload installed

    > [!TIP]
    > If you've previously installed Visual Studio, [use the Visual Studio Installer](/visualstudio/install/modify-visual-studio) to ensure that the **Office/SharePoint development** workload is installed. If this workload is not yet installed, use the Visual Studio Installer to [install it](/visualstudio/install/modify-visual-studio#modify-workloads).

- Excel 2016 or later.

## Create the add-in project

> [!NOTE]
> [Skip this section](#convert-the-add-in-project-to-typescript) if you already have an existing project.

1. In Visual Studio, choose **Create a new project**. If the Visual Studio development environment is already open, you can create a new project by choosing **File** > **New** > **Project** on the menu bar.

1. Using the search box, enter **add-in**. Choose **Excel Web Add-in**, then select **Next**.

1. Name your project and select **Create**.

1. In the **Create Office Add-in** dialog window, choose **Add new functionalities to Excel**, and then choose **Finish** to create the project.

1. Visual Studio creates a solution and its two projects appear in **Solution Explorer**. The **Home.html** file opens in Visual Studio.

## Convert the add-in project to TypeScript

### Add Nuget package

1. Open the Nuget package manager by choosing **Tools** > **Nuget Package Manager** > **Manage Nuget Packages for Solution**
1. Select the **Browse** tab. Search for and select **Microsoft.TypeScript.MSBuild**. Install this package to the ASP.NET web project, or update it if it's already installed. The ASP.NET web project has your project name with the text `Web` appended to the end. This will ensure the project will transpile to JavaScript when the build runs.

> [!NOTE]
> In your TypeScript project, you can have a mix of TypeScript and JavaScript files and your project will compile. This is because TypeScript is a typed superset of JavaScript that compiles JavaScript.

### Create a TypeScript config file

1. In **Solution Explorer**, right-click (or select and hold) the ASP.NET web project and choose **Add** > **New Item**. The ASP.NET web project has your project name with the text `Web` appended to the end.
1. In the **Add New Item** dialog, search for and select **TypeScript JSON configuration File**. Select **Add** to create a **tsconfig.json** file.
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

1. Save the file. For more information on **tsconfig.json** settings, see [What is a tsconfig.json?](https://www.typescriptlang.org/docs/handbook/tsconfig-json.html)

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
1. Open the npm project properties by going to **Tools** > **Options**, then **Projects and Solutions** > **Web Package Management** > **Package Restore**. Set both **Restore On Project Open** and **Restore On Save** to "True". Select **OK** to save the settings.

### Update the JavaScript files

Change your JavaScript files (**.js**) to TypeScript files (**.ts**). Then, make the necessary changes for them to compile. This section walks through the default files in a new project.

1. Find the **Home.js** file and rename it to **Home.ts**.

1. Find the **./Functions/FunctionFile.js** file and rename it to **FunctionFile.ts**.

1. Find the **./Scripts/MessageBanner.js** file and rename it to **MessageBanner.ts**.

1. In **./Scripts/MessageBanner.ts**, find the line `_onResize(null);` and replace it with the following:

    ```TypeScript
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

1. In Visual Studio, press <kbd>F5</kbd> or choose the **Start** button to launch Excel with the **Show Taskpane** add-in button displayed on the ribbon. The add-in will be hosted locally on IIS.

1. In Excel, choose the **Home** tab, and then choose the **Show Taskpane** button on the ribbon to open the add-in task pane.

1. In the worksheet, select the nine cells that contain numbers.

1. Press the **Highlight** button on the task pane to highlight the cell in the selected range that contains the highest value.

## See also

- [Promise implementation discussion on StackOverflow](https://stackoverflow.com/questions/44461312/office-addins-file-in-its-typescript-version-doesnt-work)
- [Office Add-in samples on GitHub](https://github.com/OfficeDev/Office-Add-in-samples)
