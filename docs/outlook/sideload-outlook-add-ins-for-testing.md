---
title: Sideload Outlook add-ins for testing
description: Use sideloading to install an Outlook add-in for testing without having to first put it in an add-in catalog.
ms.date: 08/01/2024
ms.topic: how-to
ms.localizationpriority: medium
---

# Sideload Outlook add-ins for testing

Sideload your Outlook add-in for testing without having to first put it in an add-in catalog.

> [!IMPORTANT]
> If your Outlook add-in supports mobile, sideload the manifest using the instructions in this article for your Outlook client on the web, on Windows, or on Mac, then follow the guidance in [Testing your add-ins on mobile](outlook-mobile-addins.md#testing-your-add-ins-on-mobile).

## Sideload an add-in that uses the unified app manifest for Microsoft 365

The process to sideload an add-in that uses the unified app manifest for Microsoft 365 varies depending on the tool you used to create your add-in project.

- For add-in projects created using the [Teams Toolkit](/microsoftteams/platform/toolkit/install-teams-toolkit?tabs=vscode), use one of the following options.
  - [Sideload in Visual Studio Code with Teams Toolkit](../develop/convert-xml-to-json-manifest.md#sideload-with-the-teams-toolkit)
  - [Sideload in Visual Studio Code with a system prompt, bash shell, or terminal](../develop/convert-xml-to-json-manifest.md#sideload-with-a-system-prompt-bash-shell-or-terminal)
- For add-in projects created using the [Yeoman generator for Office Add-ins](../develop/yeoman-generator-overview.md), in a command prompt, navigate to the root of your project. Then, run `npm run start:desktop`. The project builds and a Node dev-server window opens. This process may take a couple of minutes before it opens Outlook desktop.
- For NodeJS and npm projects that weren't created with the Yeoman generator, use one of the following options.
  - [Sideload with the Office-Addin-Debugging tool](../develop/convert-xml-to-json-manifest.md#sideload-with-the-office-addin-debugging-tool)
  - [Sideload with the Teams Toolkit CLI (command-line interface)](../develop/convert-xml-to-json-manifest.md#sideload-with-the-teams-toolkit-cli-command-line-interface)

## Sideload an add-in that uses an XML manifest

An Outlook add-in that uses an XML manifest can be sideloaded automatically through the command line or manually through the **Add-Ins for Outlook** dialog.

### Sideload automatically

If you created your Outlook add-in using the [Yeoman generator for Office Add-ins](../develop/yeoman-generator-overview.md), sideloading is best done through the command line. This takes advantage of our tooling and allows you to sideload across all of your supported devices in one command.

> [!NOTE]
> If you're developing on macOS, you must manually sideload your add-in after running `npm start`. For guidance, see the [Sideload manually](#sideload-manually) section of this article.

1. Open a command prompt and navigate to the root directory of your Yeoman generated add-in project. Run the command `npm start`.

1. Your Outlook add-in will automatically sideload to Outlook on your desktop computer. You'll see a dialog appear, stating there is an attempt to sideload the add-in, listing the name and the location of the manifest file. Select **OK**, which will register the manifest.

    > [!IMPORTANT]
    > If the manifest contains an error or the path to the manifest is invalid, you'll receive an error message.

1. If your manifest contains no errors and the path is valid, your add-in will now be sideloaded and available on both your desktop and in Outlook on the web. It will also be installed across all your supported devices.

### Sideload manually

Though we strongly recommend sideloading automatically through the command line as covered in the previous section, you can also manually sideload an Outlook add-in. Add-ins that use the XML manifest are manually sideloaded through the **Add-Ins for Outlook** dialog. The following steps apply to Outlook on the web, on Windows ([new](https://support.microsoft.com/office/656bb8d9-5a60-49b2-a98b-ba7822bc7627) and classic), and on Mac.

1. In your preferred browser, go to <https://aka.ms/olksideload>. Outlook on the web opens, then the **Add-Ins for Outlook** dialog appears after a few seconds.

    > [!TIP]
    >
    > - If your mailbox toolbar looks like the following image, you're using the classic version of Outlook on the web.
    >
    >   ![Header of the classic Outlook on the web toolbar with "Office 365 | Outlook" as its title.](../images/outlook-on-the-web-classic-toolbar.png)
    >
    >   To access the **Add-Ins for Outlook** dialog, open [Outlook on the web](https://outlook.office365.com). Then, select the gear icon from the top-right section of the toolbar and choose **Manage add-ins**.
    >
    >   ![The Manage add-ins option is selected in classic Outlook on the web.](../images/outlook-sideload-web-manage-integrations.png)
    >
    >   Note that your organization may include its own logo in the mailbox toolbar, so you might see something slightly different from what is shown in the preceding images.
    >
    > - In classic Outlook on Windows, you can also access the **Add-Ins for Outlook** dialog by selecting **File** > **Info** > **Manage Add-ins**. This opens Outlook on the web in your preferred browser, then loads the dialog.
    >
    > - In Outlook on Mac, starting in Version 16.85 (24051214), the **Get Add-ins** button no longer opens the **Add-Ins for Outlook** dialog. Instead, it opens [AppSource](https://appsource.microsoft.com/marketplace/apps?product=office%3Boutlook&page=1&src=office) in your default browser. Earlier versions can still access the **Add-Ins for Outlook** dialog through the **Get Add-ins** button. If you don't see **Get Add-ins** in your version of Outlook, select the ellipsis button (`...`) from the ribbon, then select **Get Add-ins**.
    >
    >   ![The Get Add-ins option is selected from the ellipsis button in Outlook on Mac.](../images/outlook-sideload-new-mac.png)

1. In the **Add-Ins for Outlook** dialog box, select **My add-ins**.

    ![The My add-ins option selected in the Add-Ins for Outlook dialog.](../images/outlook-sideload-my-add-ins-owa.png)

1. Locate the **Custom Addins** section at the bottom of the dialog box. Select the **Add a custom add-in** link, and then select **Add from File**.

    ![The Add from File option is selected in the Custom Addins section.](../images/outlook-sideload-custom-add-in.png)

    [!INCLUDE [outlook-sideloading-url](../includes/outlook-sideloading-url.md)]

1. Locate the manifest file for your custom add-in and install it. Accept all prompts during the installation.

    > [!NOTE]
    > In classic Outlook on Windows, it may take up to 24 hours for your manually sideloaded add-in to appear in the client. This is due to caching.

## Locate a sideloaded add-in

To learn how to access a sideloaded add-in in your Outlook client, see [Use add-ins in Outlook](https://support.microsoft.com/office/1ee261f9-49bf-4ba6-b3e2-2ba7bcab64c8).

## Remove a sideloaded add-in

On all versions of Outlook, the key to removing a sideloaded add-in is the **Add-Ins for Outlook** dialog, which lists your installed add-ins. To access the dialog on your Outlook client, use the steps listed for [manual sideloading](#sideload-manually) in the previous section of this article.

To remove a sideloaded add-in from Outlook, in the **Add-Ins for Outlook** dialog, navigate to the **Custom Addins** section. Choose the ellipsis (`...`) for the add-in, then choose **Remove**.

## See also

- [Add-ins for Outlook on mobile devices](outlook-mobile-addins.md)
