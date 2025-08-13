---
title: Sideload Office Add-ins that use the unified manifest for Microsoft 365
description: Test your Office Add-in on Windows by sideloading.
ms.date: 08/13/2025
ms.localizationpriority: medium
---

# Sideload Office Add-ins that use the unified manifest for Microsoft 365

The process of sideloading an add-in that uses the [Unified manifest for Microsoft 365](../develop/json-manifest-overview.md) varies depending on the tool you want to use and on how the add-in project was created.

> [!NOTE]
> An add-in that uses the unified manifest can be sideloaded on Office on Windows, Version 2304 (Build 16320.20000) or later. Sideloading on Windows has the effect of sideloading to Office on the web too. Currently, it can't be sideloaded on Mac or iPad. If you work on a Mac, you can test the add-in by having your Microsoft 365 administrator deploy the add-in through the [integrated apps portal](/microsoft-365/admin/manage/test-and-deploy-microsoft-365-apps) in the Microsoft 365 admin center.

## Sideload add-ins created with the Yeoman generator for Office Add-ins (Yo Office)

Use the process described in [Sideload with a system prompt, bash shell, or terminal](#sideload-with-a-system-prompt-bash-shell-or-terminal).

## Sideload with Microsoft 365 Agents Toolkit

1. First, *make sure Office desktop application that you want to sideload into is closed.*
1. In Visual Studio Code, open Agents Toolkit.
1. Required for Outlook only: in the **ACCOUNTS** section, verify that you're signed into Microsoft 365.
1. Select **View** | **Run** in Visual Studio Code. In the **RUN AND DEBUG** dropdown menu, select one of these options as appropriate for your add-in.

    - **Excel Desktop (Edge Chromium)**
    - **Outlook Desktop (Edge Chromium)**
    - **PowerPoint Desktop (Edge Chromium)**
    - **Word Desktop (Edge Chromium)**

1. Press <kbd>F5</kbd>. The project builds and a Node dev-server window opens. This process may take a couple of minutes and then the desktop version of the Office application that you selected opens. You can now work with your add-in. For an Outlook add-in, be sure you're working in the **Inbox** of *your Microsoft 365 account identity*.
1. To stop debugging and uninstall the add-in, select **Run** | **Stop Debugging** in Visual Studio Code. Closing the server window doesn't reliably stop the server and closing the Office application doesn't reliably cause Office to unacquire the add-in.

   > [!NOTE]
   > If the preceding step seems to have no effect, uninstall the add-in by opening a **TERMINAL** in Visual Studio Code, and then complete the uninstall step &#8212; the *last* step &#8212; of the section [Sideload with a system prompt, bash shell, or terminal](#sideload-with-a-system-prompt-bash-shell-or-terminal).

## Sideload with a system prompt, bash shell, or terminal

1. First, *make sure the Office desktop application that you want to sideload into is closed.*
1. Open a system prompt, bash shell, or the Visual Studio Code **TERMINAL**, and navigate to the root of the project.
1. The command to sideload the add-in depends on when the project was created. If the `"scripts"` section of the project's package.json file has a "start:desktop" script, then run `npm run start:desktop`; otherwise, run `npm run start`. The project builds and a Node dev-server window opens. This process may take a couple of minutes then the Office host application (Excel, Outlook, PowerPoint, or Word) desktop opens.
1. On some versions of Office, the add-in may not fully activate. For example, the add-in's buttons may not appear on the ribbon. If  this happens, select the **Add-ins** button on the **Home** ribbon. On the flyout that opens, select the add-in. This completes the installation.
1. You can now work with your add-in.
1. When you're done working with your add-in, make sure to run the command `npm run stop`. Closing the server window doesn't reliably stop the server and closing the Office application doesn't reliably cause Office to unacquire the add-in.

## Sideload other NodeJS and npm projects

There are two tools you can use to sideload.

### Sideload with the Office-Addin-Debugging tool

1. To sideload the add-in, run the following command. This command puts the unified manifest and the two icon image files that are referenced in the manifest's `"icons"` property into a zip file and sideloads it to the Office application. It also starts a server in a separate NodeJS window to host the add-in files on localhost. For more details about this command, see [Office-Addin-Debugging](https://www.npmjs.com/package/office-addin-debugging).

    ```command&nbsp;line
    npx office-addin-debugging start <relative-path-to-unified-manifest> desktop
    ```

1. When you use office-addin-debugging to start an add-in, *always stop the session with the following command*. Closing the server window doesn't reliably stop the server and closing the Office application doesn't reliably cause Office to unacquire the add-in.

    ```command&nbsp;line
    npx office-addin-debugging stop <relative-path-to-unified-manifest>
    ```

### Sideload with Microsoft 365 Agents Toolkit CLI (command-line interface)

1. Create a zip package. See [Manually create the add-in package file](#manually-create-the-add-in-package-file).

1. In the root of the project, open a command prompt or bash shell and run the following command to install the Agents Toolkit CLI.

    ```command&nbsp;line
    npm install -g @microsoft/m365agentstoolkit-cli
    ```

1. Run the following command to sideload the add-in.

    ```command&nbsp;line
    atk install --file-path <relative-path-to-zip-file>
    ```

    > [!IMPORTANT]
    > This command returns some information about the add-in including an autogenerated title ID as shown in the following example.
    >
    > :::image type="content" source="../images/atk-cli-install.png" alt-text="The command 'atk install --file-path manifests/contoso/contoso.zip' and the system response including the user's account name, the title id GUID and the app id GUID.":::
    >
    > You'll need this title ID to end the sideloading and debugging session. It is recorded on Windows computers in the following Registry key:
    >
    > **HKEY_CURRENT_USER\SOFTWARE\Microsoft\Office\16.0\Wef\Developer\OutlookSideloadManifestPath\TitleId**
    >
    > The string "Outlook" is in the key name for historical reasons, but it applies to any add-in installed with the Agents Toolkit CLI.
    >
    > Only the most recent add-in installed with the CLI is recorded. If you sideload an add-in with the CLI before you have uninstalled an earlier add-in you installed with the CLI, then there is no record of the earlier add-in's title ID in the Registry. So, we recommend that you also save it in a text file in the root of the project and name the file **TitleID.txt** on both Mac and Windows computers.

1. When you use the Agents Toolkit CLI to start an add-in, *always stop the session with the following command*. Closing the server window doesn't reliably stop the server and closing the Office application doesn't reliably cause Office to unacquire the add-in. Replace "{title ID}" with the title ID of the add-in including the "U_" prefix; for example, `U_90d141c6-cf4f-40ee-b714-9df9ea593f39`.

    ```command&nbsp;line
    atk uninstall --mode title-id --title-id {title ID} --interactive false
    ```

    > [!IMPORTANT]
    > The [documentation for the `uninstall` command](/microsoftteams/platform/toolkit/teams-toolkit-cli?pivots=version-three#teamsapp-uninstall) describes a way to use the add-in's manifest ID instead of the title ID. Due to a bug in an API that the CLI calls, this option doesn't currently work. You must use the `uninstall` command given above and you must include the `--interactive false` option.

## Sideload through the Teams app store

Add-ins that use the unified manifest can be manually sideloaded through the Teams app store, even if they have no Teams-related functionality. The steps are as follows.

1. Create an app package manually if it hasn't already been created by a tool. See [Manually create the add-in package file](#manually-create-the-add-in-package-file).
1. Close all Office applications, and then clear the Office cache following the instructions at [Manually clear the cache](../testing/clear-cache.md#manually-clear-the-cache).
1. Open Teams and select **Apps** from the app bar, then select **Manage your apps** at the bottom of the **Apps** pane.
1. Select **Upload an app** in the **Apps** dialog, and then in the dialog that opens, select **Upload a custom app**.
1. In the **Open** dialog, navigate to, and select, the app package.
1. Select **Add** in the dialog that opens.
1. When you're prompted that the app was added, *don't* open it in Teams. Instead, close Teams.
1. The next task is to start a local web server that hosts your project's HTML and JavaScript files. How you do this depends on several factors including the folder structure of your project, the tools you use, such as a bundler, task manager, server application, and how you have configured those tools. The following instruction applies only to projects that meet the following conditions.

    - There's a **webpack.config.js** file in the root of the project that is similar to the ones in add-in projects that are created with the [Yeoman Generator for Office Add-ins](../develop/yeoman-generator-overview.md) or [Microsoft 365 Agent Toolkit](../develop/agents-toolkit-overview.md).
    - There's a **package.json** file in the root of the project similar to the ones created by the same two tools and the file has a "scripts" section with the following script in it.

       ```json
      "dev-server": "webpack serve --mode development"
      ```

1. In a command prompt or Visual Studio Code **TERMINAL** in the root of the project, run `npm run dev-server` to start the server on localhost.
1. Open the Office application that the add-in targets. Wait until the add-in has loaded. This may take as much as two minutes. Depending on your version of Office, ribbon buttons and other artifacts may appear automatically. In some versions, you need to manually activate the add-in: Select the **Add-ins** button on the **Home** ribbon, and then in the flyout that opens, select your add-in. It will have the name specified in the [`"name.short"`](/microsoft-365/extensibility/schema/root-name) property of the manifest.

> [!IMPORTANT]
> When you want to end a testing session and make changes to the add-in that you sideloaded through the Teams app store, be sure to remove the add-in completely with the following steps. 
>
> 1. Close the Office application.
> 1. Shut down the server. See the documentation for your server application for how to do this. For the webpack dev-server application, shutting it down depends on whether the server is running in the same window in which you ran `npm run dev-server` or a different window. If it's the same window, give the terminal focus and press <kbd>Ctrl</kbd>+<kbd>C</kbd>. Choose "Y" in response to the prompt to end the process. If it's in a different window, then in the window where you ran `npm run dev-server`, run `npm run stop`.
> 1. Clear the Office cache following the instructions at [Manually clear the cache](../testing/clear-cache.md#manually-clear-the-cache).
> 1. Open Teams and select **Apps** from the app bar, then select **Manage your apps** at the bottom of the **Apps** pane.
> 1. Find your add-in in the list of apps. It will have the name specified in the `"name.short"` property of the manifest.
> 1. Select the add-in from the list of apps to expand its row.
> 1. Select the trash can icon and then select **Remove** in the prompt.
>
> Make your changes and then sideload the add-in again.

## Manually create the add-in package file

When the unified manifest is used, the unit of installation and sideloading is a zip-formatted package file. This file is usually created for you by the tools you use to create and test your add-in, but there are scenarios in which you create it manually. To do so, use any zip utility to create a zip file that contains the following files.

- The unified manifest, which goes in the root of the zip file.
- The two image files referenced in the `"icons"` property of the manifest.
- Any localization files that are referenced in the `"localizationInfo"` property of the manifest.
- Any declarative agent files that are referenced in the `"copilotAgents"` property.
- Any second-level supplementary files. For example, declarative agent configuration files sometimes reference second-level supplementary files, such as plugin configuration files. These should be included too.

> [!IMPORTANT]
> *All of these files must have the same relative path in the zip file as specified in the manifest.* For example, if the path of the two image files is **assets/icon-64.png** and **assets/icon-128.png**, then you must include an **assets** folder with the two files in the zip package. Second-level files, such as plugin configuration files for declarative agents, must have the same relative path in the zip file as they do in the first-level file that references them. For example, if the relative path of a declarative agent file specified in the manifest is **agents/myAgent.json**, then you must include an **agents** folder in the zip package and put the **myAgent.json** file in it. If the declarative agent file, in turn, gives the relative path of **plugins/myPlugin.json** for a plugin configuration file, then you must include a **plugins** subfolder under the **agents** folder and put the **myPlugin.json** file in it.

To maximize compatibility with Microsoft 365 development tools, we recommend that you keep the files that will be included in the package in a folder called **appPackage** in the root of your project, and that you put the package file in a subfolder named **build** in the **appPackage** folder.

The following are examples of the recommended structure. The structure inside the **\build\appPackage.zip** file must mirror the structure of the **appPackage** folder, except for the **build** folder itself.

```console
\appPackage
    \assets
        color.png
        outline.png
    \build
        appPackage.zip
    manifest.json
```

```console
\appPackage
    \agents
        myAgent.json
        \plugins
            myPlugin.json
    \assets
        color.png
        outline.png
    \build
        appPackage.zip
    \languages
        fr-FR.json
        es-MX.json
    manifest.json
```
