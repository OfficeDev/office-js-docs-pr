---
title: Sideload Outlook add-ins for testing
description: Use sideloading to install an Outlook add-in for testing without having to first put it in an add-in catalog.
ms.date: 06/24/2019
localization_priority: Normal
---

# Sideload Outlook add-ins for testing

You can use sideloading to install an Outlook add-in for testing without having to first put it in an add-in catalog.


## Sideload an add-in in Outlook in Office 365

The process for sideloading an add-in in Outlook in Office 365 depends upon whether you are using the new Outlook on the web or classic Outlook on the web.

- If your mailbox toolbar looks like the following image, see [Sideload an add-in in the new Outlook on the web](#sideload-an-add-in-in-the-new-outlook-on-the-web).

    ![partial screenshot of the new Outlook on the web toolbar](images/outlook-on-the-web-new-toolbar.png)

- If your mailbox toolbar looks like the following image, see [Sideload an add-in in classic Outlook on the web](#sideload-an-add-in-in-classic-outlook-on-the-web).

    ![partial screenshot of the classic Outlook on the web toolbar](images/outlook-on-the-web-classic-toolbar.png)

> [!NOTE]
> If your organization has included its logo in the mailbox toolbar, you might see something slightly different than shown in the preceding images.

### Sideload an add-in in the new Outlook on the web

1. Go to [Outlook in Office 365](https://outlook.office.com).

1. In Outlook on the web, create a new message.   

1. Choose **...** from the bottom of the new message and then select **Get Add-ins** from the menu that appears.

    ![Message compose window in the new Outlook on the web with Get Add-ins option highlighted](images/outlook-on-the-web-new-get-add-ins.png)

1. In the **Add-Ins for Outlook** dialog box, select **My add-ins**.

    ![Add-Ins for Outlook dialog box in the new Outlook on the web with My add-ins selected](images/outlook-on-the-web-new-my-add-ins.png)

1. Locate the **Custom add-ins** section at the bottom of the dialog box. Select the **Add a custom add-in** link, and then select **Add from file**.

    ![Manage add-ins screenshot pointing to Add from a file option](images/outlook-sideload-desktop-add-from-file.png)

1. Locate the manifest file for your custom add-in and install it. Accept all prompts during the installation.

### Sideload an add-in in classic Outlook on the web

1. Go to [Outlook in Office 365](https://outlook.office.com).

1. Choose the gear icon in the top-right section of the toolbar and select **Manage add-ins**.

    ![Outlook on the web screenshot pointing to Manage add-ins option](images/outlook-sideload-web-manage-integrations.png)

1. On the **Manage add-ins** page, select **Add-Ins**, and then select **My add-ins**.

    ![Outlook on the web store dialog with My add-ins selected](images/outlook-sideload-store-select-add-ins.png)

1. Locate the **Custom add-ins** section at the bottom of the dialog box. Select the **Add a custom add-in** link, and then select **Add from file**.

    ![Manage add-ins screenshot pointing to Add from a file option](images/outlook-sideload-desktop-add-from-file.png)

1. Locate the manifest file for your custom add-in and install it. Accept all prompts during the installation.

## Sideload an add-in in Outlook on the desktop

1. Open Outlook 2013 or later on Windows, or Outlook 2016 or later on Mac.

1. Select the **Get Add-ins** button on the ribbon.

    ![Outlook 2016 ribbon pointing to Store button](images/outlook-sideload-desktop-store.png)

    > [!NOTE]
    > If you don't see the **Get Add-ins** button in your version of Outlook, select the **Store** button on the ribbon instead.

1. Select **Add-Ins**, and then select **My add-ins**.

    ![Outlook 2016 store dialog with My add-ins selected](images/outlook-sideload-store-select-add-ins.png)

1. Locate the **Custom add-ins** section at the bottom of the dialog. Select the **Add a custom add-in** link, and then select **Add from file**.

    ![Store screenshot pointing to Add from a file option](images/outlook-sideload-desktop-add-from-file.png)

1. Locate the manifest file for your custom add-in and install it. Accept all prompts during the installation.
