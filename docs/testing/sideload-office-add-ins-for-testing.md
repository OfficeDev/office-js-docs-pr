---
title: Sideload Office Add-ins to Office on the web
description: Test your Office Add-in in Office on the web by sideloading.
ms.date: 12/03/2025
ms.localizationpriority: medium
---

# Sideload Office Add-ins to Office on the web

When you sideload an add-in, you're able to install the add-in without first putting it in an add-in catalog. This is useful when testing and developing your add-in because you can see how your add-in will appear and function.

> [!NOTE]
>
> - This article applies to **Excel**, **OneNote**, **PowerPoint**, and **Word** add-ins. For information on sideloading **Outlook** add-ins, see the article [Sideload Outlook add-ins for testing](../outlook/sideload-outlook-add-ins-for-testing.md).
> - This article applies to add-ins that use the add-in only manifest. For information about sideloading add-ins that use the [unified manifest for Microsoft 365](../develop/unified-manifest-overview.md), see [Sideload Office Add-ins that use the unified manifest for Microsoft 365](sideload-add-in-with-unified-manifest.md).

When you sideload an add-in on the web, the add-in's manifest is stored in the browser's local storage, so if you clear the browser's cache, or switch to a different browser, you have to sideload the add-in again.

The steps to sideload an add-in on the web vary based on the following factors.

- The host application (for example, Excel, Word, Outlook)
- What tool created the add-in project (for example, Visual Studio, Yeoman generator for Office Add-ins, or neither)
- Whether you are sideloading to Office on the web with a Microsoft account or with an account in a Microsoft 365 tenant

In the following list, go to the section or article that matches your scenario. Note the first scenario in the list applies to Outlook add-ins. The remaining scenarios apply to non-Outlook add-ins.

- If you're sideloading an Outlook add-in, see the article [Sideload Outlook add-ins for testing](../outlook/sideload-outlook-add-ins-for-testing.md).
- If you created the add-in using the [Yeoman generator for Office Add-ins](../develop/yeoman-generator-overview.md), see [Sideload a Yeoman-created add-in to Office on the web](#sideload-a-yeoman-created-add-in-to-office-on-the-web).
- If you created the add-in using Visual Studio, see [Sideload an add-in on the web when using Visual Studio](#sideload-an-add-in-on-the-web-when-using-visual-studio).
- For all other cases, see one of the following sections.

  - If you're sideloading to Office on the web with a Microsoft account, see [Manually sideload an add-in to Office on the web](#manually-sideload-an-add-in-to-office-on-the-web).
  - If you're sideloading to Office on the web with an account in a Microsoft 365 tenant, see [Sideload an add-in to Microsoft 365](#sideload-an-add-in-to-microsoft-365).

## Sideload a Yeoman-created add-in to Office on the web

This process is supported for **Excel**, **OneNote**, **PowerPoint**, and **Word** only. This example project assumes you're using a project created with the [Yeoman generator for Office Add-ins](../develop/yeoman-generator-overview.md).

1. Open [Office on the web](https://office.live.com/) or OneDrive. Using the **Create** option, make a document in **Excel**, **OneNote**, **PowerPoint**, or **Word**. In this new document, select **Share**, select **Copy Link**, and copy the URL.

1. Open a Command Prompt as an administrator. In the command line starting at the root directory of your project, run the following command. Replace "{url}" with the URL that you copied.

    [!INCLUDE [npm start on web command syntax](../includes/start-web-sideload-instructions.md)]

1. The first time you use this method to sideload an add-in on the web, you'll see a dialog asking you to enable developer mode. Select the checkbox for **Enable Developer Mode now** and select **OK**.

1. You'll see a second dialog box, asking if you wish to register an Office Add-in manifest from your computer. Select **Yes**.

1. Your add-in is installed. If it has an add-in command, it should appear on either the ribbon or the context menu. If it's a task pane add-in without any add-in commands, the task pane should appear.

## Sideload an add-in on the web when using Visual Studio

If you're using Visual Studio to develop your add-in, press <kbd>F5</kbd> to open an Office document in *desktop* Office, create a blank document, and sideload the add-in. When you want to sideload to *Office on the web*, the process to sideload is similar to manual sideloading to the web. The only difference is that you must update the value of the **SourceURL** element, and possibly other elements, in your manifest to include the full URL where the add-in is deployed.

1. In Visual Studio, choose **View** > **Properties Window**.

1. In the **Solution Explorer**, select the web project. This displays properties for the project in the **Properties** window.

1. In the Properties window, copy the **SSL URL**.

1. In the add-in project, open the manifest XML file. Be sure you're editing the source XML. For some project types, Visual Studio will open a visual view of the XML which won't work for the next step.

1. Search and replace all instances of **~remoteAppUrl/** with the SSL URL you just copied. You'll see several replacements depending on the project type, and the new URLs will appear similar to `https://localhost:44300/Home.html`.

1. **Save** the XML file.

1. In the **Solution Explorer**, open the context menu of the web project (for example, by right clicking on it) then choose **Debug** > **Start new instance**. This runs the web project without launching Office.

1. From Office on the web, sideload the add-in using steps described in [Manually sideload an add-in to Office on the web](#manually-sideload-an-add-in-to-office-on-the-web).

## Manually sideload an add-in to Office on the web

This method doesn't use the command line and can be accomplished using commands only within the host application (such as Excel).

1. Open [Office on the web](https://office.com/). Open a document in **Excel**, **OneNote**, **PowerPoint**, or  **Word**.

1. Select **Home** > **Add-ins**, then select **More Settings**.

1. On the **Office Add-ins** dialog, select **Upload My Add-in**.

1. **Browse** to the add-in manifest file, and then select **Upload**.

    :::image type="content" source="../images/upload-add-in.png" alt-text="The upload add-in dialog with buttons for browse, upload, and cancel.":::

1. Verify that your add-in is installed. For example, if it has an add-in command, it should appear on either the ribbon or the context menu. If it's a task pane add-in that has no add-in commands, the task pane should appear.

[!INCLUDE[Office settings tool not supported on Mac](../includes/tool-nonsupport-mac-note.md)]

## Sideload an add-in to Microsoft 365

1. Sign in to your Microsoft 365 account.

1. Open the App Launcher on the left end of the toolbar and select **Excel**, **OneNote**, **PowerPoint**, or **Word**, and then create a new document.

1. Follow steps 2 - 5 of the section [Manually sideload an add-in to Office on the web](#manually-sideload-an-add-in-to-office-on-the-web).

## Remove a sideloaded add-in

If you ran the `npm start` command and your add-in was automatically sideloaded, then run `npm stop` when you're ready to stop the dev server and uninstall your add-in.

Otherwise, to remove a sideloaded add-in, see [Uninstall add-ins under development](uninstall-add-in.md).

## See also

- [Sideload Office Add-ins on Mac](sideload-an-office-add-in-on-mac.md)
- [Sideload Office Add-ins on iPad](sideload-an-office-add-in-on-ipad.md)
- [Sideload Outlook add-ins for testing](../outlook/sideload-outlook-add-ins-for-testing.md)
- [Clear the Office cache](clear-cache.md)
