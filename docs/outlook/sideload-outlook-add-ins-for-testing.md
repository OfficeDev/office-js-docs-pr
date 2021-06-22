---
title: Sideload Outlook add-ins for testing
description: Use sideloading to install an Outlook add-in for testing without having to first put it in an add-in catalog.
ms.date: 05/13/2021
localization_priority: Normal
---

# Sideload Outlook add-ins for testing

You can use sideloading to install an Outlook add-in for testing without having to first put it in an add-in catalog.

## Sideload automatically

If you created your Outlook add-in using [the Yeoman generator for Office Add-ins](https://github.com/OfficeDev/generator-office), sideloading is best done through the command line. This will take advantage of our tooling and sideload across all of your supported devices in one command.

1. Using the command line, navigate to the root directory of your Yeoman generated add-in project. Run the command `npm start`.

1. Your Outlook add-in will automatically sideload to Outlook on your desktop computer. You'll see a dialog appear, stating there is an attempt to sideload the add-in, listing the name and the location of the manifest file. Select **OK**, which will register the manifest.

    > [!IMPORTANT]
    > If the manifest contains an error or the path to the manifest is invalid, you'll receive an error message.

1. If your manifest contains no errors and the path is valid, your add-in will now be sideloaded and available on both your desktop and in Outlook on the web. It will also be installed across all your supported devices.

## Sideload manually

Though we strongly recommend sideloading automatically through the command line as covered in the previous section, you can also manually sideload an Outlook add-in based on the Outlook client.

### Outlook on the web

The process for sideloading an add-in in Outlook on the web depends upon whether you are using the new or classic version.

- If your mailbox toolbar looks like the following image, see [Sideload an add-in in the new Outlook on the web](#new-outlook-on-the-web).

    ![Partial screenshot of the new Outlook on the web toolbar.](../images/outlook-on-the-web-new-toolbar.png)

- If your mailbox toolbar looks like the following image, see [Sideload an add-in in classic Outlook on the web](#classic-outlook-on-the-web).

    ![Partial screenshot of the classic Outlook on the web toolbar.](../images/outlook-on-the-web-classic-toolbar.png)

> [!NOTE]
> If your organization has included its logo in the mailbox toolbar, you might see something slightly different than shown in the preceding images.

### New Outlook on the web

1. Go to [Outlook on the web](https://outlook.office.com).

1. Create a new message.

1. Choose **...** from the bottom of the new message and then select **Get Add-ins** from the menu that appears.

    ![Message compose window in the new Outlook on the web with Get Add-ins option highlighted.](../images/outlook-on-the-web-new-get-add-ins.png)

1. In the **Add-Ins for Outlook** dialog box, select **My add-ins**.

    ![Add-ins for Outlook dialog box in the new Outlook on the web with My add-ins selected.](../images/outlook-on-the-web-new-my-add-ins.png)

1. Locate the **Custom add-ins** section at the bottom of the dialog box. Select the **Add a custom add-in** link, and then select **Add from file**.

    ![Manage add-ins screenshot pointing to Add from a file option.](../images/outlook-sideload-desktop-add-from-file.png)

1. Locate the manifest file for your custom add-in and install it. Accept all prompts during the installation.

### Classic Outlook on the web

1. Go to [Outlook on the web](https://outlook.office.com).

1. Choose the gear icon in the top-right section of the toolbar and select **Manage add-ins**.

    ![Outlook on the web screenshot pointing to Manage add-ins option.](../images/outlook-sideload-web-manage-integrations.png)

1. On the **Manage add-ins** page, select **Add-Ins**, and then select **My add-ins**.

    ![Outlook on the web store dialog with My add-ins selected.](../images/outlook-sideload-store-select-add-ins.png)

1. Locate the **Custom add-ins** section at the bottom of the dialog box. Select the **Add a custom add-in** link, and then select **Add from file**.

    ![Manage add-ins screenshot pointing to Add from a file option.](../images/outlook-sideload-desktop-add-from-file.png)

1. Locate the manifest file for your custom add-in and install it. Accept all prompts during the installation.

### Outlook on the desktop

#### Outlook 2016 or later

1. Open Outlook 2016 or later on Windows or Mac.

1. Select the **Get Add-ins** button on the ribbon.

    ![Outlook 2016 ribbon pointing to Get Add-ins button.](../images/outlook-sideload-desktop-store.png)

    > [!IMPORTANT]
    > If you don't see the **Get Add-ins** button in your version of Outlook, select:
    >
    > - **Store** button on the ribbon, if available.
    >
    >   OR
    >
    > - **File** menu, then select the **Manage Add-ins** button on the **Info** tab to open the **Add-ins** dialog in Outlook on the web.<br>You can see more about the web experience in the previous section [Sideload an add-in in Outlook on the web](#outlook-on-the-web).

1. If there are tabs near the top of the dialog, ensure that the **Add-ins** tab is selected. Choose **My add-ins**.

    ![Outlook 2016 store dialog with My add-ins selected.](../images/outlook-sideload-store-select-add-ins.png)

1. Locate the **Custom add-ins** section at the bottom of the dialog. Select the **Add a custom add-in** link, and then select **Add from file**.

    ![Store screenshot pointing to Add from a file option.](../images/outlook-sideload-desktop-add-from-file.png)

1. Locate the manifest file for your custom add-in and install it. Accept all prompts during the installation.

#### Outlook 2013

1. Open Outlook 2013 on Windows.

1. Select the **File** menu, then select the **Manage Add-ins** button on the **Info** tab. Outlook will open the web version in a browser.

1. Follow the steps in the [Sideload an add-in in Outlook on the web](#outlook-on-the-web) section according to your version of Outlook on the web.

## Remove a sideloaded add-in

On all versions of Outlook, the key to removing a sideloaded add-in is the **My Add-ins** dialog which lists your installed add-ins. Choose the ellipsis (`...`) for the add-in then select **Remove**.

To navigate to the **My Add-ins** dialog box for your Outlook client, use the last steps listed for [manual sideloading](#sideload-manually) in the previous sections of this article.

To remove a sideloaded add-in from Outlook, use the steps previously described in this article to find the add-in in the **Custom add-ins** section of the dialog box that lists your installed add-ins. Choose the ellipsis (`...`) for the add-in then choose **Remove** to remove that specific add-in. Close the dialog.
