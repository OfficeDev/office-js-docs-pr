---
title: Debug your event-based Outlook add-in (preview)
description: Learn how to debug your Outlook add-in that implements event-based activation.
ms.topic: article
ms.date: 05/14/2021
localization_priority: Normal
---

# Debug your event-based Outlook add-in (preview)

This article provides debugging guidance as you implement [event-based activation](autolaunch.md) in your add-in. The event-based activation feature is currently in preview.

> [!IMPORTANT]
> This debugging capability is only supported for preview in Outlook on Windows with a Microsoft 365 subscription. For more information, see the [Preview debugging for the event-based activation feature](#preview-debugging-for-the-event-based-activation-feature) section in this article.

In this article, we discuss the key stages to enable debugging.

- [Mark the add-in for debugging](#mark-your-add-in-for-debugging)
- [Configure Visual Studio Code](#configure-visual-studio-code)
- [Attach Visual Studio Code](#attach-visual-studio-code)
- [Debug](#debug)

You have several options for creating your add-in project. Depending on the option you're using, the steps may vary. Where this is the case, if you used the Yeoman generator for Office Add-ins to create your add-in project (for example, by doing the [event-based activation walkthrough](autolaunch.md)), then follow the **yo office** steps, otherwise follow the **Other** steps. Visual Studio Code should be at least version 1.56.1.

## Preview debugging for the event-based activation feature

We invite you to try out the debugging capability for the event-based activation feature! Let us know your scenarios and how we can improve by giving us feedback through GitHub (see the **Feedback** section at the end of this page).

To preview this capability for Outlook on Windows, the minimum required build is 16.0.13729.20000. For access to Office beta builds, join the [Office Insider program](https://insider.office.com).

## Mark your add-in for debugging

1. Set the registry key `HKEY_CURRENT_USER\SOFTWARE\Microsoft\Office\16.0\Wef\Developer\[Add-in ID]\UseDirectDebugger`. `[Add-in ID]` is the **Id** in the add-in manifest.

    **yo office**: In a command line window, navigate to the root of your add-in folder then run the following command.

    ```command&nbsp;line
    npm start
    ```

    In addition to building the code and starting the local server, this command should set the `UseDirectDebugger` registry key for this add-in to `1`.

    **Other**: Add the `UseDirectDebugger` registry key under `HKEY_CURRENT_USER\SOFTWARE\Microsoft\Office\16.0\WEF\Developer\[Add-in ID]\`. Replace `[Add-in ID]` with the **Id** from the add-in manifest. Set the registry key to `1`.

    [!include[Developer registry key](../includes/developer-registry-key.md)]

1. Start Outlook desktop (or restart Outlook if it's already open).
1. Compose a new message or appointment. You should see the following dialog. Do *not* interact with the dialog yet.

    ![Screenshot of Debug Event-based handler dialog.](../images/outlook-win-autolaunch-debug-dialog.png)

## Configure Visual Studio Code

### yo office

1. Back in the command line window, open Visual Studio Code.

    ```command&nbsp;line
    code .
    ```

1. In Visual Studio Code, open the file **./.vscode/launch.json** and add the following excerpt to your list of configurations. Save your changes.

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

### Other

1. Create a new folder called **Debugging** (perhaps in your **Desktop** folder).
1. Open Visual Studio Code.
1. Go to **File** > **Open Folder**, navigate to the folder you just created, then choose **Select Folder**.
1. On the Activity Bar, select the **Debug** item (Ctrl+Shift+D).

    ![Screenshot of Debug icon on the Activity Bar.](../images/vs-code-debug.png)

1. Select the **create a launch.json file** link.

    ![Screenshot of link to create a launch.json file in Visual Studio Code.](../images/vs-code-create-launch.json.png)

1. In the **Select Environment** dropdown, select **Edge: Launch** to create a launch.json file.
1. Add the following excerpt to your list of configurations. Save your changes.

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

## Attach Visual Studio Code

1. To find the add-in's **bundle.js**, open the following folder in Windows Explorer and search for your add-in's **Id** (found in the manifest).

    ```text
    %LOCALAPPDATA%\Microsoft\Office\16.0\Wef
    ```

    Open the folder prefixed with this ID and copy its full path. In Visual Studio Code, open **bundle.js** from that folder. The pattern of the file path should be as follows:

    `%LOCALAPPDATA%\Microsoft\Office\16.0\Wef\{[Outlook profile GUID]}\[encoding]\Javascript\[Add-in ID]_[Add-in Version]_[locale]\bundle.js`

1. Place breakpoints in bundle.js where you want the debugger to stop.
1. In the **DEBUG** dropdown, select the name **Direct Debugging**, then select **Run**.

    ![Screenshot of selecting Direct Debugging from configuration options in the Visual Studio Code Debug dropdown.](../images/outlook-win-autolaunch-debug-vsc.png)

## Debug

1. After confirming that the debugger is attached, return to Outlook, and in the **Debug Event-based handler** dialog, choose **OK** .

1. You can now hit your breakpoints in Visual Studio Code, enabling you to debug your event-based activation code.

## Stop debugging

To stop debugging for the rest of the current Outlook desktop session, in the **Debug Event-based handler** dialog, choose **Cancel**. To re-enable debugging, restart Outlook desktop.

To prevent the **Debug Event-based handler** dialog from popping up and stop debugging for subsequent Outlook sessions, delete the associated registry key or set its value to `0`: `HKEY_CURRENT_USER\SOFTWARE\Microsoft\Office\16.0\Wef\Developer\[Add-in ID]\UseDirectDebugger`.

## See also

- [Configure your Outlook add-in for event-based activation](autolaunch.md)
- [Debug your add-in with runtime logging](../testing/runtime-logging.md#runtime-logging-on-windows)
