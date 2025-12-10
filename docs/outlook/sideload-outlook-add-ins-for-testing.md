---
title: Sideload Outlook add-ins for testing
description: Use sideloading to install an Outlook add-in for testing without having to first put it in an add-in catalog.
ms.date: 05/19/2025
ms.topic: how-to
ms.localizationpriority: medium
---

# Sideload Outlook add-ins for testing

Sideload your Outlook add-in for testing in Outlook on the web, on Windows ([new](https://support.microsoft.com/office/656bb8d9-5a60-49b2-a98b-ba7822bc7627) or classic), on Mac, or on mobile devices without having to first put it in an add-in catalog.

The sideloading process differs depending on the type of manifest your add-in uses. Select the tab for the applicable manifest. For more information on Office Add-in manifests, see [Office Add-ins manifest](../develop/add-in-manifests.md).

> [!IMPORTANT]
> If your Outlook add-in supports mobile, sideload the manifest using the instructions in this article for your Outlook client on the web, on Windows, or on Mac, then follow the guidance in [Testing your add-ins on mobile](outlook-mobile-addins.md#testing-your-add-ins-on-mobile).

# [Unified manifest for Microsoft 365](#tab/jsonmanifest)

The process to sideload an add-in that uses the unified app manifest for Microsoft 365 varies depending on the tool you used to create your add-in project. For more information, see [Sideload Office Add-ins that use the unified manifest for Microsoft 365](../testing/sideload-add-in-with-unified-manifest.md).

# [Add-in only manifest](#tab/xmlmanifest)

An Outlook add-in that uses an add-in only manifest can be sideloaded automatically through the command line or manually through the **Add-Ins for Outlook** dialog.

### Sideload automatically

If you created your Outlook add-in using the [Yeoman generator for Office Add-ins](../develop/yeoman-generator-overview.md), sideloading is best done through the command line. This takes advantage of our tooling and allows you to sideload across all of your supported devices with one command.

1. Open a command prompt and navigate to the root directory of your Yeoman generated add-in project. Run the command `npm start`.

    > [!NOTE]
    >
    > - If you're developing on macOS, you must manually sideload your add-in after running `npm start`. For guidance, see the [Sideload manually](#sideload-manually) section of this article.
    >
    > - When you first use Yeoman generator to develop an Office Add-in, your default browser opens a window where you'll be prompted to sign in to your Microsoft 365 account. If a sign-in window doesn't appear and you encounter a sideloading or login timeout error, run `atk auth login m365` before running `npm start` again.

1. A dialog appears stating an attempt to sideload the add-in. It lists the name and location of the manifest file. Select **OK** to register the manifest.

    > [!IMPORTANT]
    > If the manifest contains an error or the path to the manifest is invalid, you'll receive an error message.

1. If your manifest contains no errors and the path is valid, your add-in will now be sideloaded and available on the Outlook desktop client and on the web. It will also be installed across all your supported devices.

### Sideload manually

Though we strongly recommend sideloading automatically through the command line as covered in the previous section, you can also manually sideload an Outlook add-in. Add-ins that use the add-in only manifest are manually sideloaded through the **Add-Ins for Outlook** dialog.

The following steps apply to Outlook on the web, on Windows (new and classic), and on Mac.

1. In your preferred browser, go to <https://aka.ms/olksideload>. Outlook on the web opens, then the **Add-Ins for Outlook** dialog appears after a few seconds.

    > [!TIP]
    >
    > - In classic Outlook on Windows, you can also access the **Add-Ins for Outlook** dialog by selecting **File** > **Info** > **Manage Add-ins**. This opens Outlook on the web in your preferred browser, then loads the dialog.
    >
    > - To access the **Add-Ins for Outlook** dialog in the classic version of Outlook on the web or in Outlook on Mac prior to Version 16.85 (24051214), see [Access the add-ins dialog in classic Outlook on the web or earlier versions of Outlook on Mac](#access-the-add-ins-dialog-in-classic-outlook-on-the-web-or-earlier-versions-of-outlook-on-mac).

1. In the **Add-Ins for Outlook** dialog box, select **My add-ins**.

    :::image type="content" source="../images/outlook-sideload-my-add-ins-owa.png" alt-text="The My add-ins option selected in the Add-Ins for Outlook dialog.":::

1. Locate the **Custom Addins** section at the bottom of the dialog box. Select the **Add a custom add-in** link, and then select **Add from File**.

    :::image type="content" source="../images/outlook-sideload-custom-add-in.png" alt-text="The Add from File option is selected in the Custom Addins section.":::

    [!INCLUDE [outlook-sideloading-url](../includes/outlook-sideloading-url.md)]

1. Locate the manifest file for your custom add-in and install it. Accept all prompts during the installation.

    > [!NOTE]
    > In classic Outlook on Windows, it may take up to 24 hours for your manually sideloaded add-in to appear in the client. This is due to caching.

1. The add-in is sideloaded to Outlook. Although the add-in was sideloaded through the **Add-Ins for Outlook** dialog in Outlook on the web, the add-in should also appear in supported Outlook desktop clients.

## Access the add-ins dialog in classic Outlook on the web or earlier versions of Outlook on Mac

### Classic Outlook on the web

In Outlook on the web, if your mailbox toolbar looks like the following image, you're using the classic version of Outlook on the web.

:::image type="content" source="../images/outlook-on-the-web-classic-toolbar.png" alt-text="Header of the classic Outlook on the web toolbar with 'Office 365 | Outlook' as its title.":::

To access the **Add-Ins for Outlook** dialog, open [Outlook on the web](https://outlook.office365.com). Then, select the gear icon from the top-right section of the toolbar and choose **Manage add-ins**.

 :::image type="content" source="../images/outlook-sideload-web-manage-integrations.png" alt-text="The Manage add-ins option is selected in classic Outlook on the web.":::

Note that your organization may include its own logo in the mailbox toolbar, so you might see something slightly different from what is shown in the preceding images.

### Earlier versions of Outlook on Mac

In Outlook on Mac, starting in Version 16.85 (24051214), the **Get Add-ins** button no longer opens the **Add-Ins for Outlook** dialog. Instead, it opens [Microsoft Marketplace](https://marketplace.microsoft.com/marketplace/apps?product=office%3Boutlook&page=1&src=office) in your default browser. Earlier versions can still access the **Add-Ins for Outlook** dialog through the **Get Add-ins** button. If you don't see **Get Add-ins** in your version of Outlook, select the ellipsis button (`...`) from the ribbon, then select **Get Add-ins**.

:::image type="content" source="../images/outlook-sideload-new-mac.png" alt-text="The Get Add-ins option is selected from the ellipsis button in Outlook on Mac.":::

## Remove a sideloaded add-in

If you ran the `npm start` command and your add-in was automatically sideloaded, then run the command `npm stop` when you're ready to stop the dev server and uninstall your add-in. If you ran `npm run start`, then run the command `npm run stop` instead.

Otherwise, on all versions of Outlook, the key to removing a sideloaded add-in is the **Add-Ins for Outlook** dialog, which lists your installed add-ins. To access the dialog on your Outlook client, use the steps listed for [manual sideloading](#sideload-manually) in the previous section of this article.

To manually remove a sideloaded add-in from Outlook, in the **Add-Ins for Outlook** dialog, navigate to the **Custom Addins** section. Choose the ellipsis (`...`) for the add-in, then choose **Remove**.

---

## Locate a sideloaded add-in

To learn how to access a sideloaded add-in in your Outlook client, see [Use add-ins in Outlook](https://support.microsoft.com/office/1ee261f9-49bf-4ba6-b3e2-2ba7bcab64c8).

## See also

- [Add-ins for Outlook on mobile devices](outlook-mobile-addins.md)
