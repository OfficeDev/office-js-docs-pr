---
title: Convert an add-in to use the unified manifest for Microsoft 365
description: Learn the various methods for converting an add-in with an XML manifest to the unified manifest for Microsoft 365 and sideload the add-in.
ms.topic: how-to
ms.date: 07/27/2023
ms.localizationpriority: medium
---

# Convert an add-in to use the unified manifest for Microsoft 365 (preview)

To add Teams capabilities to an add-in that uses the XML manifest, or to just future proof the add-in, you need to convert it to use the unified manifest for Microsoft 365.

There are three basic tasks to converting an add-in project from the XML manifest to the unified manifest.

- Ensure that you have 64x64 pixel and 128x128 pixel images files to serve as icons for the add-in.
- Convert the XML manifest itself to the JSON format of the unified manifest.
- Package the new manifest and main icon and high resolution icon image files into a zip file for sideloading or deployment.

> [!NOTE]
> 
> - The unified manifest is a preview feature for Office Add-ins and is currently supported only for Outlook on Windows.
> - Add-ins that use the unified manifest can be sideloaded only on Office build 16.0.16501.10000 or later.
> - Projects created in Visual Studio, as distinct from Visual Studio Code, can't be converted at this time.
> - If you [created the project with Teams Toolkit](teams-toolkit-overview.md) or with the "unified manifest" option in the [Office Yeoman Generator](yeoman-generator-overview.md), it already uses the unified manifest.

## Ensure that you have the two image files

It's highly likely that you have the two image files already in your add-in. They're specified in the [IconUrl](/javascript/api/manifest/iconurl) and [HighResolutionIconUrl](/javascript/api/manifest/highresolutioniconurl) elements in the XML manifest. Because these are both optional elements, your add-in project mignt not have one or both. These are required for an add-in that uses the unified manifest, so you need to create them if the project doesn't already have them. For guidance, see [Create an icon for your add-in](/office/dev/store/create-effective-office-store-listings#create-an-icon-for-your-add-in). 

When you've added the files to the project, add **\<IconUrl\>** and **\<HighResolutionIconUrl\>** (in that order) to the XML manifest just below the **\<Description\>** element. The following is an example.

```xml
<OfficeApp xmlns="http://schemas.microsoft.com/office/appforoffice/1.1" xmlns:xsi="http://www.w3.org/2001/XMLSchema-instance" xsi:type="MailApp">
  <Id>01234567-89ab-cdef-0123-4567-89abcdef0123</Id>
  <Version>1.0</Version>
  <ProviderName>Contoso</ProviderName>
  <DefaultLocale>en-us</DefaultLocale>
  <DisplayName DefaultValue="Great Add-in"/>
  <Description DefaultValue="A great add-in."/>
  <IconUrl DefaultValue="https://localhost:3000/assets/icon-64.png" />
  <HighResolutionIconUrl DefaultValue="https://localhost:300/assets/icon-128.png" />

  <!-- Other markup omitted -->
```

## Conversion tools and options

There are several ways to carry out the remaining tasks, depending on the IDE and other tools you want to use for your project, and on the tool you used to create the project. 

- [Convert the project with Teams Toolkit](#convert-the-project-with-teams-toolkit)
- [Projects created with the Office Yeoman Generator (aka "Yo Office")](#projects-created-with-the-office-yeoman-generator-aka-yo-office)
- [NodeJS and npm projects that aren't created with Yeoman Generator](#nodejs-and-npm-projects-that-arent-created-with-yeoman-generator)

### Convert the project with Teams Toolkit

The easiest way to convert is to use Teams Toolkit.

#### Prerequisites

- Install [Visual Studio Code](https://code.visualstudio.com/)
- Install [Teams Toolkit](/microsoftteams/platform/toolkit/install-teams-toolkit?tabs=vscode#install-teams-toolkit-for-visual-studio-code)

#### Import the add-in project to Teams Toolkit

1. Open Visual Studio Code and select the Teams Toolkit icon on the **Activity Bar**.

    :::image type="content" source="../images/teams-toolkit-icon.png" alt-text="Teams Toolkit icon.":::

1. Select **Create a new app**.
1. In the **New Project** drop down, select **Outlook Add-in**.

    :::image type="content" source="../images/teams-toolkit-create-outlook-add-in.png" alt-text="The four options in New Project drop down. The fourth option is called 'Outlook add-in'.":::

1. In the **App Features Using an Outlook Add-in** drop down, select **Import an Existing Outlook Add-in**.

    :::image type="content" source="../images/teams-toolkit-create-outlook-task-pane-capability.png" alt-text="The two options in the App Features Using an Outlook Add-in drop down. The second option is called 'Import an Existing Outlook add-in'.":::

1. In the **Existing add-in project folder** drop down, browse to the root folder of the add-in project.
1. In the **Select import project manifest file** drop down, browse to the XML manifest file.
1. In the **Workspace folder** dialog, select the folder where you want to put the converted project.
1. In the **Application name** dialog, give a name to the project (with no spaces). Teams Toolkit creates the project with your source files and scaffolding. It then opens the project *in a second Visual Studio Code window*. Close the original Visual Studio Code window.

#### Sideload the add-in in Visual Studio Code

You can sideload the add-in using the Teams Toolkit or in a command prompt, bash shell, or terminal.

##### Sideload with the Teams Toolkit

1. First, *make sure Outlook desktop is closed.*
1. In Visual Studio Code, open the Teams Toolkit.
1. In the **ACCOUNTS** section, verify that you're signed into Microsoft 365.
1. Select **View** | **Run** in Visual Studio Code. In the **RUN AND DEBUG** drop down menu, select the option, **Outlook Desktop (Edge Chromium)**, and then press **F5**. The project builds and a Node dev-server window opens. This process may take a couple of minutes and then Outlook desktop opens.
1. You can now work with your add-in. Be sure you're working in the **Inbox** of *your Microsoft 365 account identity*.

##### Sideload with a system prompt, bash shell, or terminal

1. First, *make sure Outlook desktop is closed.*
1. Open a system prompt, bash shell, or the Visual Studio Code **TERMINAL**, and navigate to the root of the project.
1. Run the command `npm run start:desktop`. The project builds and a Node dev-server window opens. This process may take a couple of minutes then Outlook desktop opens.
1. You can now work with your add-in.

### Projects created with the Office Yeoman Generator (aka "Yo Office")

If the project was created with the Office Yeoman Generator (using any option except the "unified manifest" option) and you don't want to use the Teams Toolkit, convert it using the following steps.

1. In the root of the project, open a command prompt or bash shell and run the following command. This converts the manifest and updates the package.json to specify current tooling packages. The new unified manifest is in the root of the project and the old XML manifest is in a backup.zip file. For details about this command, see [Office-Addin-Project](https://www.npmjs.com/package/office-addin-project).

    ```command&nbsp;line
    npx office-addin-project convert -m <relative-path-to-XML-manifest>
    ``` 
    
1. Run `npm install`.
1. To sideload the add-in, run `npm run start:desktop`. This command puts the unified manifest and the two image files into a zip file and sideloads it to the Office application. It also starts server in a separate NodeJS window to host the add-in files on localhost.

### NodeJS and npm projects that aren't created with Yeoman Generator

If you don't want to use the Teams Toolkit and your project wasn't created with the Office Yeoman generator, use the office-addin-manifest-converter tool.

In the root of the project, open a command prompt or bash shell and run the following command. This command puts the unified manifest in a subfolder with the same name as the filename stem of the original XML manifest. For example, if the manifest is named MyManifest.xml, the unified manifest is created at .\MyManifest\MyManifest.json. For more details about this command, see [Office-Addin-Manifest-Converter](https://www.npmjs.com/package/office-addin-manifest-converter).

```command&nbsp;line
npx office-addin-manifest-converter convert -m <relative-path-to-XML-manifest>
```

Once you have the unified manifest created, there are two ways to create the zip file and sideload it. They are described in the next two subsections.

#### Sideload with the Office-Addin-Debugging tool

1. To sideload the add-in, run the following command. This command puts the unified manifest and two default icon image files into a zip file and sideloads it to the Office application. It also starts a server in a separate NodeJS window to host the add-in files on localhost. Note that you pass the path to the *unified manifest* that you created in the previous step. For more details about this command, see [Office-Addin-Debugging](https://www.npmjs.com/package/office-addin-debugging).

    ```command&nbsp;line
    npx office-addin-debugging start <relative-path-to-unified-manifest> desktop
    ``` 

1. When you use office-addin-debugging to start an add-in, *always stop the session with the following command*. Closing the server window doesn't reliably stop the server and closing the Office application doesn't reliably cause Office to unacquire the add-in.

    ```command&nbsp;line
    npx office-addin-debugging stop <relative-path-to-unified-manifest>
    ``` 

#### Sideload with the Teams Toolkit CLI (command-line interface) 

1. Manually create the zip package using the following steps.

    1. Open the unified manifest and scroll to the "icons" property. Note the relative path of the two image files.
    1. Use any zip utility to create a zip file that contains the unified manifest and the two image files. *The image files must have the same relative path in the zip file as they do in the project.* For example, if the relative path is "assets/icon-64.png" and "assets/icon-128.png", then you must include the "assets" folder with the two files in the zip package. 
    1. If the folder contains other files, such as image files used in the Office ribbon, remove these from the zip package. It should have only the two image files specified in the "icons" property (in addition to the manifest in the root of the zip package). 

1. In the root of the project, open a command prompt or bash shell and run the following commands.

    ```command&nbsp;line
    npm install -g @microsoft/teamsfx-cli

    teamsfx m365 sideloading --file-path <relative-path-to-zip-file>
    ``` 


