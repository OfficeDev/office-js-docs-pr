---
title: Uninstall add-ins under development
description: Learn how to prevent incomplete uninstallation of add-ins you are developing and how to remove incompletely uninstalled add-ins under development.
ms.topic: troubleshooting-problem-resolution
ms.date: 05/19/2025
ms.localizationpriority: medium
---

# Uninstall add-ins under development

Incompletely removed add-ins can leave artifacts on your computer, such as custom ribbon buttons or registry entries, during development. In this article, we call these "ghost add-ins".

Outlook add-ins also might add these artifacts to other computers when you sign into Outlook on them with the same ID as you used to develop the add-in.  

   > [!IMPORTANT]
   > When you sign into Outlook, it downloads from Exchange, and sideloads, all the Outlook add-in manifests that are associated with your ID, *including add-ins that you are developing on a different computer using the same ID*. For example, any custom ribbon buttons defined in the manifest will appear for the add-in.
   >
   > If the URLs in the manifest point to a non-localhost server and that server is running and accessible to the non-development computer, then Outlook caches the add-in's files in the local file system and the add-in usually runs normally on the computer. Otherwise, the add-in doesn't function, but visible parts of it, such as custom ribbon buttons appear. They have the labels defined in the manifest. The add-in's button icons also appear if they were ever cached locally on the non-development computer and the cache was never cleared. Icon files aren't stored with Exchange, so if they were never cached on the non-development computer (or the cache has been cleared), then the buttons have default icons.
   >
   > Until the add-in's registration is removed from Exchange, the add-in will continue to appear. See [Remove a ghost add-in](#remove-a-ghost-add-in) for information about removing the registration in Exchange.

This article provides some guidance to minimize the chance of these problems and to resolve them if they occur.

## Prevent the problems

When an add-in is sideloaded, several things happen:

- A web server, usually on localhost, is started to serve the add-in's files (such as the HTML, CSS, and JavaScript files).
- These same files are cached on your development computer.
- The add-in is registered with the development computer. The registration is done with Registry entries on a Windows computer or with certain files saved to the file system on a Mac.
- Most tools for sideloading add-ins automatically open the Office application that the add-in targets. The tools also populate the application with any custom ribbon buttons or context menu items that are defined in the add-in's manifest.
- For an Outlook add-in, the add-in's manifest is registered with the Exchange service.

### Use your tool's uninstall facility

To prevent ghost add-ins, end every testing, debugging, and sideloading session by using the uninstall (also called unacquire) option that is provided by the tool that you used to start the session. Doing this reverses the effects of sideloading, as stated earlier in this article.

The following list identifies, for each tool, how to uninstall but doesn't describe the procedures or syntax in detail. *Be sure to use the links to get complete instructions.*

> [!NOTE]
> Some of these tools don't close the Office application that opened automatically. In that case, close the application manually immediately after ending the session.

- **Yeoman generator for Office Add-ins (Yo Office)**: Use the `npm stop` script at the same command line where you started the session with `npm start`. For more information, see the various articles in the **Get started** and **Quick starts** sections and [Remove a sideloaded add-in](sideload-office-add-ins-for-testing.md).
- **Microsoft 365 Agents Toolkit for Visual Studio Code**: Select **Run** | **Stop Debugging** in Visual Studio Code. For more information, see the last step of [Create an Outlook Add-in project](../tutorials\outlook-tutorial.md#create-an-outlook-add-in-project) which also applies to non-Outlook add-ins.
- **Office Add-in Development Kit for Visual Studio Code**: With the Office Add-in Development Kit extension open, select **Stop Previewing Your Office Add-in**. For more information, see [Stop testing your add-in](../develop/development-kit-overview.md?tabs=vscode#stop-testing-your-office-add-in).
- **office-addin-debugging tool**: Use the `office-addin-debugging stop` command at the same command line where you started the session with `office-addin-debugging start`. For more information, see [Sideload with the Office-Addin-Debugging tool](../testing\sideload-add-in-with-unified-manifest.md#sideload-with-the-office-addin-debugging-tool).
- **Microsoft 365 Agents Toolkit CLI**: Use the `atk uninstall` command at the same command line where you started the session with `atk install`. For more information, see [Sideload with Microsoft 365 Agents Toolkit CLI](../testing\sideload-add-in-with-unified-manifest.md#sideload-with-microsoft-365-agents-toolkit-cli-command-line-interface).
- **Visual Studio**: Select **DEBUG** | **Stop debugging** in the menu, or press <kbd>Shift</kbd>+<kbd>F5</kbd>, or click the square red "stop" button on the debugging bar. Alternatively, closing the Office application also stops the session and uninstalls the add-in. For more information, see [First look at the Visual Studio debugger](/visualstudio/debugger/debugger-feature-tour).

## Remove a ghost add-in

To remove a ghost add-in, you need to remove the artifacts that were created when it was last sideloaded, remove its local registration, and for Outlook add-ins remove its registration in Exchange.

> [!TIP]
> There's a fast way to remove a ghost add-in on Windows computers if the add-in was installed with the Teams Toolkit CLI. Try this first, and if it works, you can skip the remainder of this section.
>
> 1. Obtain the add-in's title ID from the Registry key **HKEY_CURRENT_USER\SOFTWARE\Microsoft\Office\16.0\Wef\Developer\OutlookSideloadManifestPath\TitleId**. (The string "Outlook" is in the key name for historical reasons, but it applies to any add-in installed with the Agents Toolkit CLI.)
> 1. Run the following command in a command prompt, bash shell, or terminal. Replace "{title ID}" with the title ID of the add-in including the "U_" prefix; for example, `U_90d141c6-cf4f-40ee-b714-9df9ea593f39`.
>
>    ```command&nbsp;line
>    atk uninstall --mode title-id --title-id {title ID} --interactive false
>    ```

The process for removing the add-in varies depending on whether the add-in is for Outlook or some other Office application.

> [!NOTE]
> In the [unified manifest for Microsoft 365](../develop/unified-manifest-overview.md), an add-in can be configured to support Outlook and one or more other Office applications; that is, there's more than one member of the [`"extensions.requirements.scopes"`](/microsoft-365/extensibility/schema/requirements-extension-element#scopes) array in the manifest and one of the members is `"mail"` (or the `"extensions.requirements.scopes"` property isn't present). Treat an add-in that is configured in this way as an Outlook add-in.

If the ghost add-in is not an Outlook add-in, skip to the section [Remove the add-in artifacts](#remove-the-add-in-artifacts).

### Remove the Exchange registration of a ghost Outlook add-in

1. Log into Outlook with the same ID you used when you sideloaded the add-in.
1. Open PowerShell as an Administrator.
1. Run the following commands. Answer "Yes" to all confirmation prompts.

   ```powershell
   Install-Module -Name ExchangeOnlineManagement -RequiredVersion 3.4.0
   Set-ExecutionPolicy RemoteSigned
   Connect-ExchangeOnline
   ```

      > [!NOTE]
      > If the `Connect-ExchangeOnline` command returns the error "ActiveX control '8856f961-340a-11d0-a96b-00c04fd705a2' cannot be instantiated because the current thread is not in a single-threaded apartment", just run the command a second time. This is a well-known bug.

1. Run the following command. Answer "Yes" to all confirmation prompts.

   ```powershell
   Get-App | Format-Table -Auto DisplayName,AppId
   ```

   A list of the add-ins installed on Outlook displays. These will include built-in Microsoft add-ins and add-ins you have installed. Any ghost Outlook add-ins will also be listed.

1. Find the ghost add-in in the list. If it was created with Yo Office or another Microsoft tool, it probably has the name "Contoso Task Pane Add-in".
1. Copy the App ID (a GUID) of the add-in. You need it for later steps.
1. Run the command `Remove-App -Identity {{The GUID OF YOUR ADD-IN HERE}}` (e.g., `Remove-App -Identity 26ead0cb-10dd-4ba2-86c6-4db111876652`). This command removes the add-in from Exchange.

   > [!WARNING]
   > The removal of the registration needs to propagate to all Exchange servers. Wait at least three hours before continuing with the next step.

1. Continue with the section [Remove the add-in artifacts](#remove-the-add-in-artifacts).

### Remove the add-in artifacts

> [!IMPORTANT]
> Carry out this procedure on all devices on which you have had the add-in sideloaded.

1. Log out from all Office applications and then close them all, including Outlook.
1. [Clear the Office cache](clear-cache.md). If the ghost add-in supports Outlook, use [Clear the cache in Outlook manually](clear-cache.md#manually-clear-the-cache-in-outlook).
1. Continue with the section [Remove the local registration](#remove-the-local-registration).

### Remove the local registration

> [!IMPORTANT]
> Carry out this procedure on all computers on which you have had the add-in sideloaded.

1. Delete the local registration of the ghost add-in. The process varies depending on the operating system.

   #### [Windows](#tab/windows)

   1. Open the **Registry Editor**.
   1. Navigate to **Computer\HKEY_CURRENT_USER\Software\Microsoft\Office\16.0\WEF\Developer**. This key lists the add-ins that are currently sideloaded, or were sideloaded in the past and weren't fully uninstalled. The **Data** value for each entry is the path to the add-in's manifest. The **Name** value varies depending on which version of which tool was used to create and sideload the add-in. If Visual Studio was used, the name is typically is also the path to the manifest. For other tools, the name is typically the add-in's ID. When an Office application launches, it reloads all add-ins listed in this key (that support the Office application). Reloading may have no practical or discernable effect if the add-in's artifacts have been deleted from the cache, or the manifest no longer exists at the path, or the add-in's files aren't being served by a server.

      Find the entry for the ghost add-in and delete it. If it is an Outlook add-in, then you have the ID from [removing the Exchange registration](#remove-the-exchange-registration-of-a-ghost-outlook-add-in). You can also use the path in the **Data** column to find the manifest to help identify the add-in the entry refers to and read the ID from the manifest. If any manifests listed in the **Data** column no longer exist at the specified path, then delete the entries for those manifests.

      :::image type="content" source="../images/addinRegistrationWindowsManifestPath.png" alt-text="The Windows registry for the key named Computer\HKEY_CURRENT_USER\Software\Microsoft\Office\16.0\WEF\Developer." border="false":::

   1. Expand the **... Developer** node in the registry tree. Look for a subkey whose name is the same ghost add-in's ID. If it is there, delete it.

      :::image type="content" source="../images/addinRegistrationWindowsDeveloperSubkeys.png" alt-text="The Windows registry for the key named Computer\HKEY_CURRENT_USER\Software\Microsoft\Office\16.0\WEF\Developer expanded to show subkeys." border="false":::

   1. Navigate to **Computer\HKEY_USERS\\{SID}\Software\Microsoft\Office\16.0\WEF\Developer**, where **{SID}** is the [SID](/windows-server/identity/ad-ds/manage/understand-security-identifiers) of the user you were signed in with when you sideloaded the add-in, and repeat the preceding two steps.
   1. Navigate to **Computer\HKEY_CURRENT_USER\Software\Microsoft\Office\16.0\Common\CustomUIValidationCache**. In the **Name** column, find all the entries that begin with the add-in's ID (a GUID) and delete them. Then navigate to **Computer\HKEY_USERS\\{SID}\Software\Microsoft\Office\16.0\Common\CustomUIValidationCache**, where **{SID}** is the SID of the user you were signed in with when you sideloaded the add-in, and repeat the process.

      :::image type="content" source="../images/addinRegistrationWindows.png" alt-text="The Windows registry for the key named Computer\HKEY_USERS\SID\Software\Microsoft\Office\16.0\Common\CustomUIValidationCache**, where SID is the SID of a user." border="false":::

   #### [Mac](#tab/mac)

   For non-Outlook add-ins, the local registration on a Mac is removed when you clear the cache. See [Remove the add-in artifacts](#remove-the-add-in-artifacts).

   For ghost Outlook add-ins, remove the local registration on a Mac by using the **Add-Ins for Outlook** dialog in Outlook. Follow the guidance at [Remove a sideloaded Outlook add-in](../outlook/sideload-outlook-add-ins-for-testing.md#remove-a-sideloaded-add-in).

   ---

2. If you are removing an Outlook add-in, continue with the section [Test for removal of Outlook add-ins](#test-for-removal-of-outlook-add-ins).

### Test for removal of Outlook add-ins

Open Outlook with the same identity you used when you created the add-in. If artifacts from the add-in (such as custom ribbon buttons) reappear after a few minutes or if event handlers from the add-in seem to be active, then the removal of the add-in's registration from Exchange hasn't propagated to all Exchange servers. Wait at least three hours and then repeat the procedures in the sections [Remove the add-in artifacts](#remove-the-add-in-artifacts) and [Remove the local registration](#remove-the-local-registration) on the computer where you observed the artifacts.

## See also

- [Troubleshoot development errors with Office Add-ins](troubleshoot-development-errors.md)
- [Clear the Office cache](clear-cache.md)
- The PowerShell reference for [Install-Module](/powershell/module/powershellget/install-module), [Set-ExecutionPolicy](/powershell/module/microsoft.powershell.security/set-executionpolicy), [Connect-ExchangeOnline](/powershell/exchange/connect-to-exchange-online-powershell), and [Get-App](/powershell/module/exchange/get-app).
