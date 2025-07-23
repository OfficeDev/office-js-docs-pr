---
title: Sideload Office Add-ins that use the unified manifest for Microsoft 365
description: Test your Office Add-in on Windows by sideloading.
ms.date: 05/19/2025
ms.localizationpriority: medium
---

# Sideload Office Add-ins that use the unified manifest for Microsoft 365

The process of sideloading an add-in that uses the [Unified manifest for Microsoft 365](../develop/json-manifest-overview.md) varies depending on the tool you want to use and on how the add-in project was created. 

> [!NOTE]
> An add-in that uses the unified manifest can be sideloaded on Office on Windows, Version 2304 (Build 16320.20000) or later. Currently, it can't be sideloaded on the web, Mac, or iPad.

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
1. For an Excel, PowerPoint, or Word add-in, there is an additional step: select the **Add-ins** button on the **Home** ribbon. On the flyout that opens, select the add-in. This completes the installation.
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

1. Manually create a zip package using the following steps.

    1. Open the unified manifest and scroll to the `"icons"` property. Note the relative path of the two image files.
    1. Use any zip utility to create a zip file that contains the unified manifest and the two image files. *The image files must have the same relative path in the zip file as they do in the project.* For example, if the relative path is "assets/icon-64.png" and "assets/icon-128.png", then you must include the `"assets"` folder with the two files in the zip package.
    1. If the folder contains other files, such as image files used in the Office ribbon, remove these from the zip package. It should have only the two image files specified in the `"icons"` property (in addition to the manifest in the root of the zip package).

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
