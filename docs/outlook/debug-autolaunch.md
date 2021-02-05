---
title: Debug your event-based Outlook add-in (preview)
description: Learn how to debug your Outlook add-in that implements event-based activation.
ms.topic: article
ms.date: 02/05/2021
localization_priority: Normal
---

# Debug your event-based Outlook add-in (preview)

This article is meant to provide debugging guidance as you implement [event-based activation](autolaunch.md) in your add-in. The event-based activation feature is currently in preview.

> [!IMPORTANT]
> This debugging capability is only supported for preview in Outlook on the web and Windows with a Microsoft 365 subscription. See [How to preview debugging for the event-based activation feature](#how-to-preview-debugging-for-the-event-based-activation-feature) in this article for more details.

In this article, we'll discuss the key steps to enable debugging. Note that this flow assumes that you created your add-in by following the steps of the [walkthrough built on Yeoman generator for Office Add-ins](autolaunch.md).

- [Mark the add-in for debugging](#mark-the-add-in-for-debugging)
- [Configure Visual Studio Code](#configure-visual-studio-code)
- [Attach VS Code](#attach-vs-code)
- [Debug](#debug)

## How to preview debugging for the event-based activation feature

We invite you to try out the debugging capability for the event-based activation feature! Let us know your scenarios and how we can improve by giving us feedback through GitHub (see the **Feedback** section at the end of this page).

To preview this capability:

- For Outlook on the web:
  - [Configure targeted release on your Microsoft 365 tenant](/microsoft-365/admin/manage/release-options-in-office-365?view=o365-worldwide&preserve-view=true#set-up-the-release-option-in-the-admin-center).
  - Reference the **beta** library on the CDN (https://appsforoffice.microsoft.com/lib/beta/hosted/office.js). The [type definition file](https://appsforoffice.microsoft.com/lib/beta/hosted/office.d.ts) for TypeScript compilation and IntelliSense is found at the CDN and [DefinitelyTyped](https://raw.githubusercontent.com/DefinitelyTyped/DefinitelyTyped/master/types/office-js-preview/index.d.ts). You can install these types with `npm install --save-dev @types/office-js-preview`.
- For Outlook on Windows: The minimum required build is 16.0.13729.20000. Join the [Office Insider program](https://insider.office.com) for access to Office beta builds.

## Mark your add-in for debugging

1. In a command line window, navigate to the root of your add-in folder then run the following command.

    ```command&nbsp;line
    npm start
    ```

    In addition to building the code and starting the local server, this command should set a registry key for this add-in to 1: `HKEY_CURRENT_USER\SOFTWARE\Microsoft\Office\16.0\Wef\Developer\[Add-in Id]\UseDirectDebugger`. `[Add-in Id]` is the Id in the add-in manifest.

1. Open Outlook desktop (or reopen if it's already open).
1. Compose a new message or appointment. You should see the following dialog. Do *not* interact with the dialog yet.

    ![Screenshot of Debug Event-based handler dialog](../images/outlook-win-autolaunch-debug-dialog.png)

## Configure Visual Studio Code

1. Back in the command line window, open VS Code.

    ```command&nbsp;line
    code .
    ```

1. Open the file **./.vscode/launch.json** and add the following excerpt to your list of configurations. Save your changes.

    ```json
    {
      "name": "Direct Debugging",
      "type": "node",
      "request": "attach",
      "port": 9229,
      "protocol": "inspector",
      "timeout": 600000,
      "trace": true
    }
    ```

## Attach VS Code

1. To find the add-in's **bundle.js**, open the `%LOCALAPPDATA%\Microsoft\Office\16.0\Wef` folder in Windows Explorer then search for your add-in's Id (found in the manifest). Open the folder prefixed with this GUID then copy the full path. In VS Code, open **bundle.js** from that folder.

    `%LOCALAPPDATA%\Microsoft\Office\16.0\Wef\{[Outlook profile GUID]}\[Outlook profile code]\Javascript\[Add-in Id_...]\bundle.js`

1. Place breakpoints in bundle.js where you want the debugger to stop.
1. In the **DEBUG** dropdown, select the name **Direct Debugging** then select the **Run** button.

    ![Screenshot of selecting Direct Debugging from configuration options in the VS Code Debug dropdown](../images/outlook-win-autolaunch-debug-vsc.png)

## Debug

1. After confirming that the debugger is attached, return to Outlook then choose the **OK** button in the **Debug Event-based handler** dialog.

1. Your breakpoints in VS Code should now be hit, enabling you to debug your event-based activation code.

## Stop debugging

- To stop debugging for the rest of the current Outlook desktop session, choose the **Cancel** button in the **Debug Event-based handler** dialog. To re-enable debugging, reopen Outlook desktop.
- To stop debugging for subsequent Outlook sessions and prevent the **Debug Event-based handler** dialog from popping up, delete the registry key or set its value to 0: `HKEY_CURRENT_USER\SOFTWARE\Microsoft\Office\16.0\Wef\Developer\[Add-in Id]`

## See also

[Configure your Outlook add-in for event-based activation](autolaunch.md)
